"""Fill missing vocabulary fields with a locally downloaded ModelScope model.

The script downloads (or reuses) a Qwen model from ModelScope once, then sends
one vocabulary item through the local model at a time.  The workbook is saved
after every processed row so a long run can be resumed safely.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any


DEFAULT_MODEL_ID = "Qwen/Qwen3.5-9B"

FIELD_HEADERS = {
    "translate": "translate",
    "part_of_speech": "part of speech",
    "transformation": "Transformation",
    "memory_techniques": "Memory techniques",
    "distinguishing_between_similar_words": (
        "Distinguishing between similar words"
    ),
}
RESPONSE_FIELDS = tuple(FIELD_HEADERS)

JSON_OUTPUT_FORMAT = """{
  "translate": "中文释义",
  "part_of_speech": "词性缩写",
  "transformation": "词形变化；没有特殊变化时填写无",
  "memory_techniques": "记忆方法",
  "distinguishing_between_similar_words": "近义词辨析；没有必要对比时填写无"
}"""

SYSTEM_PROMPT = """你是一名严谨的考研英语词汇教师和词典编辑。
你的任务是为一个英文单词补全词汇表字段。必须基于常见、可靠的英语用法，
面向中国考研英语学习者，内容准确、简洁、可直接放入 Excel 单元格。

只输出一个合法的 JSON 对象，不要 Markdown 代码围栏，不要解释 JSON 之外的内容。
JSON 必须严格按照下面的格式输出，键名必须完全一致，不能增加、删除或改名：
""" + JSON_OUTPUT_FORMAT + """

所有键都必须出现且只能出现一次；所有值都必须是非空字符串。
没有内容时填写“无”，不能填写 null、数组、对象或数字。

字段要求：
1. translate：给出最常见的中文释义；有多个重要词性时分号分隔，并标注词性。
2. part_of_speech：使用简洁英文缩写，例如 "n.; v."、"adj."、"prep."。
3. transformation：给出重要词形变化，如复数、过去式、过去分词、现在分词、
   比较级/最高级等；格式清晰。没有特殊变化时写 "无"，不要臆造词形。
4. memory_techniques：优先给出词根词缀、构词法或可靠联想；无法可靠拆分时，
   给出简短、合理的记忆联想。不要编造词源，控制在 1-2 句。
5. distinguishing_between_similar_words：列出最容易混淆的 1-3 个词，如adopt和adapt，说明核心
   区别和典型用法；没有必要对比时写 "无"。不要为了凑数列生僻词，更不要生成 "passage 与 passage 易混淆..." 这类同一个词的幻觉。

不要把频次编号当成释义，也不要输出例句、音标或额外字段。"""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Download a ModelScope Qwen model and fill missing vocabulary fields locally."
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=Path(os.getenv("VOCAB_INPUT", "考研生词.xlsx")),
        help="Input workbook path (default: 考研生词.xlsx).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output workbook path; defaults to updating the input file in place.",
    )
    parser.add_argument(
        "--sheet",
        default=None,
        help="Worksheet name; defaults to the active worksheet.",
    )
    parser.add_argument(
        "--request-interval",
        type=float,
        default=float(os.getenv("MODELSCOPE_REQUEST_INTERVAL", "0.5")),
        help="Seconds to wait between successful requests.",
    )
    parser.add_argument(
        "--max-retries",
        type=int,
        default=3,
        help="Retries for transient model/generation failures.",
    )
    parser.add_argument(
        "--model",
        dest="model_id",
        default=os.getenv("MODELSCOPE_MODEL", DEFAULT_MODEL_ID),
        help=(
            "ModelScope model ID (default: Qwen/Qwen3.5-9B, or "
            "MODELSCOPE_MODEL)."
        ),
    )
    parser.add_argument(
        "--model-cache-dir",
        "--model-dir",
        dest="model_cache_dir",
        type=Path,
        default=Path(os.getenv("MODELSCOPE_CACHE_DIR", "models")),
        help=(
            "Local ModelScope cache directory (default: ./models, or "
            "MODELSCOPE_CACHE_DIR)."
        ),
    )
    parser.add_argument(
        "--device",
        default=os.getenv("MODEL_DEVICE", "auto"),
        help="Inference device, such as auto, cpu, cuda, or cuda:0 (default: auto).",
    )
    parser.add_argument(
        "--max-new-tokens",
        type=int,
        default=int(os.getenv("MODEL_MAX_NEW_TOKENS", "1200")),
        help="Maximum number of tokens generated for each word (default: 1200).",
    )
    return parser.parse_args()


def normalise_header(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).lower()


def find_columns(sheet: Any) -> dict[str, int]:
    headers = {
        normalise_header(cell.value): cell.column
        for cell in sheet[1]
        if cell.value is not None
    }
    missing = [
        header
        for header in FIELD_HEADERS.values()
        if normalise_header(header) not in headers
    ]
    if "word" not in headers:
        missing.insert(0, "word")
    if missing:
        raise ValueError(
            "Workbook is missing required header(s): " + ", ".join(missing)
        )
    return {
        "word": headers["word"],
        "number": headers.get("number"),
        **{
            key: headers[normalise_header(header)]
            for key, header in FIELD_HEADERS.items()
        },
    }


def is_blank(value: Any) -> bool:
    return value is None or not str(value).strip()


def build_user_prompt(word: str, frequency: Any, missing_fields: list[str]) -> str:
    missing = ", ".join(missing_fields)
    frequency_text = str(frequency).strip() if not is_blank(frequency) else "未提供"
    return f"""请处理下面这个考研英语单词，并严格按系统要求输出 JSON。

单词：{word}
词表中的频次/编号：{frequency_text}
需要补全的字段：{missing}

即使只需要补全部分字段，也必须返回系统要求的全部 5 个键；无需补全的键仍给出
准确内容，但程序只会写入原本为空的单元格。

必须严格返回以下 JSON 格式（不要 Markdown 代码围栏或额外文字）：
{JSON_OUTPUT_FORMAT}
所有 5 个键都必须存在，所有值都必须是非空字符串。"""


def extract_json_object(text: str) -> dict[str, Any]:
    cleaned = text.strip()
    try:
        value = json.loads(cleaned, object_pairs_hook=_strict_object_pairs)
    except (json.JSONDecodeError, ValueError) as exc:
        raise ValueError(f"Invalid strict JSON returned by model: {text[:300]!r}") from exc
    if not isinstance(value, dict):
        raise ValueError("Model response JSON must be an object")
    return value


def _strict_object_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    """Reject duplicate keys instead of silently keeping the last value."""

    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"Duplicate JSON key: {key!r}")
        result[key] = value
    return result


def content_to_text(content: Any) -> str:
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts = []
        for item in content:
            if isinstance(item, str):
                parts.append(item)
            elif isinstance(item, dict) and isinstance(item.get("text"), str):
                parts.append(item["text"])
        return "".join(parts)
    return str(content)


def validate_completion_result(value: Any) -> dict[str, str]:
    """Validate and normalise the model's fixed JSON response schema."""

    if not isinstance(value, dict):
        raise ValueError("Model response JSON must be an object")

    expected_keys = set(RESPONSE_FIELDS)
    actual_keys = set(value)
    missing_keys = expected_keys - actual_keys
    extra_keys = actual_keys - expected_keys
    if missing_keys or extra_keys:
        details = []
        if missing_keys:
            details.append("missing: " + ", ".join(sorted(missing_keys)))
        if extra_keys:
            details.append("unexpected: " + ", ".join(sorted(extra_keys)))
        raise ValueError("Model response JSON has invalid keys (" + "; ".join(details) + ")")

    result: dict[str, str] = {}
    for key in RESPONSE_FIELDS:
        field_value = value[key]
        if not isinstance(field_value, str) or not field_value.strip():
            raise ValueError(f"Model response field {key!r} must be a non-empty string")
        result[key] = field_value.strip()
    return result


@dataclass
class LocalQwen:
    """Loaded tokenizer/model pair used for all workbook rows."""

    tokenizer: Any
    model: Any
    input_device: Any
    max_new_tokens: int


def download_model(model_id: str, cache_dir: Path) -> Path:
    """Download ``model_id`` into a local ModelScope cache and return its path."""

    try:
        from modelscope import snapshot_download
    except ImportError as exc:
        raise RuntimeError(
            "Local inference requires ModelScope. Install it with "
            "`pip install -r requirements.txt`."
        ) from exc

    cache_dir = cache_dir.expanduser().resolve()
    cache_dir.mkdir(parents=True, exist_ok=True)
    print(f"Downloading or reusing ModelScope model {model_id!r} in {cache_dir}...")
    try:
        model_path = snapshot_download(model_id, cache_dir=str(cache_dir))
    except Exception as exc:
        raise RuntimeError(
            f"Could not download ModelScope model {model_id!r} to {cache_dir}"
        ) from exc
    model_path = Path(model_path).expanduser().resolve()
    if not model_path.exists():
        raise RuntimeError(f"ModelScope returned a missing model path: {model_path}")
    print(f"Using local model at {model_path}")
    return model_path


def load_local_model(
    model_id: str,
    cache_dir: Path,
    device: str,
    max_new_tokens: int,
) -> LocalQwen:
    """Download and load a Qwen checkpoint for local text generation."""

    if max_new_tokens <= 0:
        raise ValueError("max_new_tokens must be greater than zero")

    model_path = download_model(model_id, cache_dir)
    try:
        import torch
        from transformers import AutoModelForCausalLM, AutoTokenizer
    except ImportError as exc:
        raise RuntimeError(
            "Local inference requires Transformers and PyTorch. Install them with "
            "`pip install -r requirements.txt`."
        ) from exc

    try:
        tokenizer = AutoTokenizer.from_pretrained(
            str(model_path),
            trust_remote_code=True,
            local_files_only=True,
        )
    except (OSError, RuntimeError, ValueError) as exc:
        raise RuntimeError(f"Could not load tokenizer from {model_path}") from exc
    if (
        getattr(tokenizer, "pad_token_id", None) is None
        and getattr(tokenizer, "eos_token_id", None) is not None
    ):
        tokenizer.pad_token = tokenizer.eos_token

    requested_device_name = (device or "auto").strip().lower()
    requested_device = requested_device_name
    if requested_device == "auto":
        requested_device = "cuda" if torch.cuda.is_available() else "cpu"
    if requested_device.startswith("cuda") and not torch.cuda.is_available():
        raise RuntimeError(
            f"Device {device!r} was requested, but this PyTorch installation has no CUDA."
        )

    model_kwargs: dict[str, Any] = {
        "trust_remote_code": True,
        "local_files_only": True,
        "torch_dtype": "auto",
    }
    # device_map=auto lets Accelerate place a 7B model across available GPUs.
    # CPU and an explicitly selected device avoid requiring Accelerate just to
    # move the model after loading.
    use_device_map = requested_device_name == "auto" and requested_device == "cuda"
    if use_device_map:
        model_kwargs["device_map"] = "auto"

    try:
        model = AutoModelForCausalLM.from_pretrained(str(model_path), **model_kwargs)
    except (OSError, RuntimeError, ValueError) as exc:
        raise RuntimeError(f"Could not load local model from {model_path}") from exc
    if not use_device_map:
        model = model.to(torch.device(requested_device))
    model.eval()

    try:
        input_device = next(model.parameters()).device
    except StopIteration as exc:
        raise RuntimeError("The loaded model has no parameters") from exc
    return LocalQwen(
        tokenizer=tokenizer,
        model=model,
        input_device=input_device,
        max_new_tokens=max_new_tokens,
    )


def generate_local_completion(model: LocalQwen, messages: list[dict[str, str]]) -> str:
    """Generate one assistant response from the local Qwen model."""

    tokenizer = model.tokenizer
    try:
        prompt = tokenizer.apply_chat_template(
            messages,
            tokenize=False,
            add_generation_prompt=True,
            enable_thinking=False,
        )
    except (AttributeError, TypeError):
        # ``enable_thinking`` is not available in older Qwen tokenizers.
        try:
            prompt = tokenizer.apply_chat_template(
                messages,
                tokenize=False,
                add_generation_prompt=True,
            )
        except AttributeError:
            prompt = "\n".join(
                f"{message['role']}: {message['content']}" for message in messages
            ) + "\nassistant:"

    model_inputs = tokenizer(prompt, return_tensors="pt")
    model_inputs = {
        key: value.to(model.input_device) for key, value in model_inputs.items()
    }
    input_length = model_inputs["input_ids"].shape[-1]

    # Importing torch here keeps --help and workbook validation usable without
    # installing the heavyweight local-inference dependencies.
    import torch

    with torch.inference_mode():
        generated_ids = model.model.generate(
            **model_inputs,
            max_new_tokens=model.max_new_tokens,
            do_sample=False,
            use_cache=True,
        )
    generated_text = tokenizer.decode(
        generated_ids[0][input_length:],
        skip_special_tokens=True,
    )
    return generated_text.strip()


def request_completion(
    word: str,
    frequency: Any,
    missing_fields: list[str],
    model: LocalQwen,
    max_retries: int,
) -> dict[str, str]:
    last_error: Exception | None = None
    retry_limit = max(0, max_retries)
    for attempt in range(retry_limit + 1):
        try:
            content = generate_local_completion(
                model,
                [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {
                        "role": "user",
                        "content": build_user_prompt(word, frequency, missing_fields),
                    },
                ],
            )
            raw_result = extract_json_object(content_to_text(content))
            return validate_completion_result(raw_result)
        except Exception as exc:
            last_error = exc
            if attempt >= retry_limit:
                break
            delay = min(30.0, 2**attempt)
            print(f"Generation failed for {word!r} ({exc}); retrying in {delay:g}s...")
            time.sleep(delay)
    raise RuntimeError(f"Could not process {word!r} after retries") from last_error


def save_workbook_atomic(workbook: Any, output_path: Path) -> None:
    output_path = output_path.resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_name: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=f"{output_path.stem}.",
            suffix=".tmp.xlsx",
            dir=output_path.parent,
            delete=False,
        ) as temp_file:
            temp_name = temp_file.name
        workbook.save(temp_name)
        os.replace(temp_name, output_path)
        temp_name = None
    finally:
        if temp_name:
            try:
                os.unlink(temp_name)
            except FileNotFoundError:
                pass


def fill_workbook(
    input_path: Path,
    output_path: Path,
    sheet_name: str | None,
    request_interval: float,
    max_retries: int,
    model: LocalQwen,
) -> None:
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise RuntimeError(
            "Excel processing requires openpyxl. Install dependencies with "
            "`pip install -r requirements.txt`."
        ) from exc

    workbook = load_workbook(input_path)
    sheet = workbook[sheet_name] if sheet_name else workbook.active
    columns = find_columns(sheet)
    processed = 0
    skipped = 0
    failed = 0

    for row_number in range(2, sheet.max_row + 1):
        word_value = sheet.cell(row_number, columns["word"]).value
        if is_blank(word_value):
            skipped += 1
            continue
        word = str(word_value).strip()
        missing_fields = [
            key
            for key in FIELD_HEADERS
            if is_blank(sheet.cell(row_number, columns[key]).value)
        ]
        if not missing_fields:
            skipped += 1
            continue

        number_column = columns.get("number")
        frequency = (
            sheet.cell(row_number, number_column).value
            if number_column is not None
            else None
        )
        print(f"[{row_number}/{sheet.max_row}] processing {word!r}...")
        try:
            result = request_completion(
                word, frequency, missing_fields, model=model, max_retries=max_retries
            )
        except RuntimeError as exc:
            failed += 1
            skipped += 1
            print(f"Skipping {word!r} after retry limit: {exc}")
            continue

        print(f"Output JSON for {word!r}:")
        print(json.dumps(result, ensure_ascii=False, indent=2))
        for key in missing_fields:
            value = result.get(key, "")
            if value:
                sheet.cell(row_number, columns[key]).value = value
        save_workbook_atomic(workbook, output_path)
        processed += 1
        if request_interval > 0:
            time.sleep(request_interval)

    print(
        f"Done. Processed {processed} word(s), skipped {skipped} row(s); "
        f"failed after retries: {failed}."
    )


def main() -> None:
    args = parse_args()
    if not args.input.exists():
        raise SystemExit(f"Input workbook not found: {args.input}")
    try:
        model = load_local_model(
            model_id=args.model_id,
            cache_dir=args.model_cache_dir,
            device=args.device,
            max_new_tokens=args.max_new_tokens,
        )
    except (RuntimeError, ValueError) as exc:
        raise SystemExit(str(exc)) from exc
    output_path = args.output or args.input
    fill_workbook(
        input_path=args.input,
        output_path=output_path,
        sheet_name=args.sheet,
        request_interval=args.request_interval,
        max_retries=max(0, args.max_retries),
        model=model,
    )


if __name__ == "__main__":
    main()
