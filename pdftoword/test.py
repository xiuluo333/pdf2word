"""Extract text covered by native PDF highlighter annotations.

The command first runs MinerU (unless ``--no-mineru`` is supplied), then uses
PyMuPDF to read the PDF's Highlight annotations and match them with words by
their page geometry.  This is important because MinerU's JSON output contains
text and layout, but normally does not preserve the PDF annotation objects.

Examples::

    python test.py paper.pdf --color '#ffff00'
    python test.py paper.pdf --eyedropper --output highlights.json
    python test.py paper.pdf --eyedropper --excel highlights.xlsx --review

``--eyedropper`` opens a small PDF page picker. Clicking a highlighted area
uses the annotation's original color; clicking elsewhere samples the rendered
page pixel as a fallback.
"""

from __future__ import annotations

import argparse
import json
import math
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Iterable, Sequence
from xml.sax.saxutils import escape


MINERU_TIMEOUT_SECONDS = 1800
DEFAULT_COLOR_TOLERANCE = 48.0
NAMED_COLORS = {
    "yellow": "#ffff00",
    "green": "#00ff00",
    "lime": "#00ff00",
    "cyan": "#00ffff",
    "blue": "#0000ff",
    "pink": "#ff69b4",
    "magenta": "#ff00ff",
    "orange": "#ffa500",
}


def _fitz():
    """Import PyMuPDF lazily so JSON/colour helpers remain dependency-free."""
    try:
        import pymupdf  # type: ignore
        return pymupdf
    except ImportError as exc:  # pragma: no cover - depends on local extras
        try:
            import fitz  # type: ignore
            return fitz
        except ImportError:
            raise RuntimeError("PyMuPDF is required. Install it with: pip install pymupdf") from exc


def parse_color(value: str | Sequence[float] | None) -> tuple[int, int, int]:
    """Return an RGB tuple in the 0..255 range.

    Accepted strings include common names, ``#rgb``, ``#rrggbb``,
    ``rgb(r,g,b)`` and comma separated values. PDF annotation colours are
    also accepted as float triples in the 0..1 range.
    """
    if value is None:
        raise ValueError("a colour is required")
    if not isinstance(value, str):
        values = list(value)
        if len(values) < 3:
            raise ValueError("colour must contain three channels")
        scale = 255 if max(float(channel) for channel in values[:3]) <= 1 else 1
        return tuple(max(0, min(255, round(float(channel) * scale))) for channel in values[:3])  # type: ignore[return-value]
    text = value.strip().lower()
    if text in NAMED_COLORS:
        text = NAMED_COLORS[text]
    if text.startswith("#"):
        text = text[1:]
        if len(text) == 3:
            text = "".join(channel * 2 for channel in text)
        if len(text) != 6:
            raise ValueError("hex colour must be #rgb or #rrggbb")
        try:
            return tuple(int(text[index:index + 2], 16) for index in (0, 2, 4))  # type: ignore[return-value]
        except ValueError as exc:
            raise ValueError("invalid hex colour") from exc
    if text.startswith("rgb(") and text.endswith(")"):
        text = text[4:-1]
    parts = [part.strip() for part in text.replace(" ", ",").split(",") if part.strip()]
    if len(parts) != 3:
        raise ValueError("colour must be hex, rgb(r,g,b), or r,g,b")
    return parse_color(tuple(float(part) for part in parts))


def color_hex(rgb: Sequence[float]) -> str:
    channels = parse_color(rgb)
    return "#" + "".join(f"{channel:02x}" for channel in channels)


def color_distance(first: Sequence[float], second: Sequence[float]) -> float:
    """Euclidean RGB distance, where 0 means an exact match."""
    left = parse_color(first)
    right = parse_color(second)
    return math.sqrt(sum((a - b) ** 2 for a, b in zip(left, right)))


def _number(value: Any, fallback: float = 0.0) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return fallback
    return number if math.isfinite(number) else fallback


def _bbox(value: Any) -> list[float] | None:
    if isinstance(value, dict):
        x0 = _number(value.get("x0", value.get("x", value.get("left"))))
        y0 = _number(value.get("y0", value.get("y", value.get("top"))))
        x1 = _number(value.get("x1", value.get("right")), x0 + _number(value.get("width")))
        y1 = _number(value.get("y1", value.get("bottom")), y0 + _number(value.get("height")))
        value = [x0, y0, x1, y1]
    if not isinstance(value, (list, tuple)):
        # fitz.Rect and similar geometry objects are iterable but are not
        # instances of list/tuple.
        try:
            value = list(value)
        except (TypeError, ValueError):
            return None
    if len(value) < 4:
        return None
    values = [_number(item) for item in value[:4]]
    x0, y0, x1, y1 = min(values[0], values[2]), min(values[1], values[3]), max(values[0], values[2]), max(values[1], values[3])
    if x1 <= x0 or y1 <= y0:
        return None
    return [x0, y0, x1, y1]


def _text(value: Any) -> str:
    if isinstance(value, str):
        return " ".join(value.split())
    if isinstance(value, (list, tuple)):
        return " ".join(part for part in (_text(item) for item in value) if part)
    if isinstance(value, dict):
        for key in ("text", "content", "value", "latex", "span_content", "paragraph_content"):
            if key in value:
                return _text(value[key])
    return "" if value is None else str(value).strip()


def _page_number(value: Any, zero_based: bool = False) -> int:
    try:
        number = int(value)
    except (TypeError, ValueError):
        return 1
    return number + 1 if zero_based and number >= 0 else max(1, number)


def _find_mineru(command: str | None = None) -> str | None:
    if command:
        path = Path(command)
        if path.is_file() and path.stat().st_mode & 0o111:
            return str(path)
        return shutil.which(command)
    candidates = [
        Path(__file__).resolve().parent / ".venv" / "bin" / "mineru",
        Path(sys.prefix) / "bin" / "mineru",
    ]
    return next((str(path) for path in candidates if path.is_file() and path.stat().st_mode & 0o111), None) or shutil.which("mineru")


def run_mineru(pdf_path: str | Path, output_dir: str | Path, command: str | None = None, timeout: int = MINERU_TIMEOUT_SECONDS) -> Path:
    """Run the official MinerU CLI and return its output directory."""
    source = Path(pdf_path).expanduser().resolve()
    destination = Path(output_dir).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(f"PDF not found: {source}")
    executable = _find_mineru(command)
    if not executable:
        raise RuntimeError("MinerU CLI was not found. Install mineru or pass --mineru-command.")
    destination.mkdir(parents=True, exist_ok=True)
    try:
        completed = subprocess.run(
            [executable, "-p", str(source), "-o", str(destination), "-b", "pipeline"],
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(f"MinerU timed out after {timeout} seconds") from exc
    except OSError as exc:
        raise RuntimeError(f"could not start MinerU: {exc}") from exc
    if completed.returncode:
        detail = (completed.stderr or completed.stdout).strip().splitlines()
        raise RuntimeError("MinerU failed" + (f": {detail[-1]}" if detail else ""))
    return destination


def _json_priority(path: Path) -> tuple[int, int, str]:
    name = path.name
    priority = 0 if name.endswith("_middle.json") else 1 if name.endswith("_content_list.json") else 2 if name.endswith("_content_list_v2.json") else 3
    return priority, len(path.parts), name


def _mineru_items_from_middle(page_infos: Iterable[Any]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []

    def visit(block: Any, page: int, page_width: float = 0, page_height: float = 0) -> None:
        if not isinstance(block, dict):
            return
        found_span = False
        for line in block.get("lines", []) or []:
            if not isinstance(line, dict):
                continue
            for span in line.get("spans", []) or []:
                if not isinstance(span, dict):
                    continue
                box = _bbox(span.get("bbox"))
                text = _text(span.get("content") or span.get("text") or span.get("value"))
                if box and text:
                    result.append({"page": page, "text": text, "bbox": box, "page_width": page_width, "page_height": page_height})
                    found_span = True
        for child in block.get("blocks", []) or []:
            visit(child, page, page_width, page_height)
        if not found_span:
            box = _bbox(block.get("bbox"))
            text = _text(block.get("text") or block.get("content"))
            if box and text:
                result.append({"page": page, "text": text, "bbox": box, "page_width": page_width, "page_height": page_height})

    for index, page_info in enumerate(page_infos or []):
        if not isinstance(page_info, dict):
            continue
        has_index = "page_idx" in page_info
        page = _page_number(page_info.get("page_idx", index), zero_based=has_index)
        size = page_info.get("page_size") or {}
        width = _number(size.get("width") if isinstance(size, dict) else (size[0] if size else 0))
        height = _number(size.get("height") if isinstance(size, dict) else (size[1] if size else 0))
        blocks = page_info.get("para_blocks") or page_info.get("preproc_blocks") or []
        for block in [*blocks, *(page_info.get("discarded_blocks", []) or [])]:
            visit(block, page, width, height)
    return result


def load_mineru_items(output_dir: str | Path) -> list[dict[str, Any]]:
    """Load text/bboxes from the first usable MinerU JSON result."""
    directory = Path(output_dir)
    for path in sorted(directory.rglob("*.json"), key=_json_priority):
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, UnicodeDecodeError, json.JSONDecodeError):
            continue
        if isinstance(payload, dict) and isinstance(payload.get("pdf_info"), list):
            items = _mineru_items_from_middle(payload["pdf_info"])
            if items:
                return items
        candidates = payload if isinstance(payload, list) else payload.get("content_list", payload.get("pages", [])) if isinstance(payload, dict) else []
        if isinstance(candidates, dict):
            candidates = candidates.get("items", candidates.get("content", []))
        items = []
        for index, item in enumerate(candidates or []):
            nested = item if isinstance(item, list) else [item]
            for candidate in nested:
                if not isinstance(candidate, dict):
                    continue
                page_index = candidate.get("page_idx", candidate.get("page_id", candidate.get("page", index)))
                page = _page_number(page_index, zero_based="page_idx" in candidate)
                box = _bbox(candidate.get("bbox") or candidate.get("text_bbox") or candidate.get("box"))
                text = _text(candidate.get("text") or candidate.get("content") or candidate.get("latex") or candidate.get("value"))
                if box and text:
                    items.append({"page": page, "text": text, "bbox": box, "page_width": _number(candidate.get("page_width", candidate.get("width_px"))), "page_height": _number(candidate.get("page_height", candidate.get("height_px")))})
        if items:
            return items
    return []


def _annotation_color(annotation: Any) -> tuple[int, int, int]:
    colors = getattr(annotation, "colors", {}) or {}
    value = colors.get("stroke") or colors.get("fill")
    return parse_color(value or (1, 1, 0))


def _annotation_boxes(annotation: Any) -> list[list[float]]:
    vertices = getattr(annotation, "vertices", None)
    points = []
    try:
        vertices = list(vertices) if vertices else []
    except TypeError:
        vertices = []
    if vertices and all(isinstance(point, (int, float)) for point in vertices):
        values = vertices
        points = [(_number(values[index]), _number(values[index + 1])) for index in range(0, len(values) - 1, 2)]
    elif vertices:
        for point in vertices:
            x, y = getattr(point, "x", None), getattr(point, "y", None)
            if x is None and isinstance(point, (tuple, list)) and len(point) >= 2:
                x, y = point[0], point[1]
            if x is not None and y is not None:
                points.append((_number(x), _number(y)))
    boxes = []
    for index in range(0, len(points) - 3, 4):
        quad = points[index:index + 4]
        boxes.append([min(point[0] for point in quad), min(point[1] for point in quad), max(point[0] for point in quad), max(point[1] for point in quad)])
    if boxes:
        return boxes
    box = _bbox(getattr(annotation, "rect", None))
    return [box] if box else []


def _intersection_ratio(first: Sequence[float], second: Sequence[float]) -> float:
    x0 = max(first[0], second[0])
    y0 = max(first[1], second[1])
    x1 = min(first[2], second[2])
    y1 = min(first[3], second[3])
    intersection = max(0.0, x1 - x0) * max(0.0, y1 - y0)
    area = max(1.0, (first[2] - first[0]) * (first[3] - first[1]))
    return intersection / area


def _matches_word(word_box: Sequence[float], boxes: Sequence[Sequence[float]]) -> bool:
    center = ((word_box[0] + word_box[2]) / 2, (word_box[1] + word_box[3]) / 2)
    for box in boxes:
        if _intersection_ratio(word_box, box) >= 0.08 or (box[0] <= center[0] <= box[2] and box[1] <= center[1] <= box[3]):
            return True
    return False


def _annotation_records(document: Any) -> list[dict[str, Any]]:
    records = []
    for page_index, page in enumerate(document):
        annotations = page.annots() or []
        for annotation_index, annotation in enumerate(annotations):
            raw_type = getattr(annotation, "type", (None,))
            annotation_type = raw_type[0] if isinstance(raw_type, (tuple, list)) else raw_type
            if annotation_type != 8:  # PDF annotation type 8 is Highlight.
                continue
            records.append({"page": page_index + 1, "annotation_index": annotation_index, "color": _annotation_color(annotation), "boxes": _annotation_boxes(annotation)})
    return records


def _extract_native_words(document: Any, records: list[dict[str, Any]], requested_color: tuple[int, int, int] | None, tolerance: float) -> list[dict[str, Any]]:
    output = []
    for record in records:
        if requested_color is not None and color_distance(record["color"], requested_color) > tolerance:
            continue
        page = document[record["page"] - 1]
        words = []
        for word in page.get_text("words", sort=True) or []:
            if len(word) < 5 or not str(word[4]).strip():
                continue
            box = _bbox(word[:4])
            if box and _matches_word(box, record["boxes"]):
                words.append({"text": str(word[4]).strip(), "bbox": box})
        if not words:
            continue
        bbox = [min(word["bbox"][0] for word in words), min(word["bbox"][1] for word in words), max(word["bbox"][2] for word in words), max(word["bbox"][3] for word in words)]
        output.append({"page": record["page"], "annotation_index": record["annotation_index"], "color": color_hex(record["color"]), "text": " ".join(word["text"] for word in words), "bbox": bbox, "words": words})
    return output


def _extract_mineru_items(items: list[dict[str, Any]], records: list[dict[str, Any]], requested_color: tuple[int, int, int] | None, tolerance: float) -> list[dict[str, Any]]:
    output = []
    for record in records:
        if requested_color is not None and color_distance(record["color"], requested_color) > tolerance:
            continue
        matched = [item for item in items if item.get("page") == record["page"] and _bbox(item.get("bbox")) and _matches_word(_bbox(item["bbox"]) or [], record["boxes"])]
        if not matched:
            continue
        bbox_values = [_bbox(item["bbox"]) for item in matched]
        boxes = [box for box in bbox_values if box]
        output.append({"page": record["page"], "annotation_index": record["annotation_index"], "color": color_hex(record["color"]), "text": " ".join(item["text"] for item in matched), "bbox": [min(box[0] for box in boxes), min(box[1] for box in boxes), max(box[2] for box in boxes), max(box[3] for box in boxes)], "words": [{"text": item["text"], "bbox": _bbox(item["bbox"])} for item in matched]})
    return output


def _mineru_box_to_page(item: dict[str, Any], page: Any) -> list[float] | None:
    """Convert MinerU coordinates to the PDF page coordinate system."""
    box = _bbox(item.get("bbox"))
    if not box:
        return None
    page_width, page_height = _number(item.get("page_width")), _number(item.get("page_height"))
    target_width, target_height = _number(getattr(page.rect, "width", 0)), _number(getattr(page.rect, "height", 0))
    if page_width > 0 and page_height > 0 and target_width > 0 and target_height > 0:
        # Middle JSON normally uses PDF points; content-list JSON may use a
        # 1000x1000 normalized canvas or image pixels.
        if abs(page_width - target_width) > 2 or abs(page_height - target_height) > 2:
            return [box[0] * target_width / page_width, box[1] * target_height / page_height, box[2] * target_width / page_width, box[3] * target_height / page_height]
    return box


def _pixmap_matches_color(pixmap: Any, box: Sequence[float], page: Any, target: tuple[int, int, int], tolerance: float) -> bool:
    page_width = max(1.0, _number(getattr(page.rect, "width", 0)))
    page_height = max(1.0, _number(getattr(page.rect, "height", 0)))
    x0 = max(0, min(pixmap.width - 1, math.floor(box[0] / page_width * pixmap.width)))
    y0 = max(0, min(pixmap.height - 1, math.floor(box[1] / page_height * pixmap.height)))
    x1 = max(x0 + 1, min(pixmap.width, math.ceil(box[2] / page_width * pixmap.width)))
    y1 = max(y0 + 1, min(pixmap.height, math.ceil(box[3] / page_height * pixmap.height)))
    step = max(1, int(min(x1 - x0, y1 - y0) / 30))
    samples = matches = 0
    for y in range(y0, y1, step):
        for x in range(x0, x1, step):
            pixel = pixmap.pixel(x, y)
            if len(pixel) < 3:
                continue
            samples += 1
            if color_distance(pixel[:3], target) <= tolerance:
                matches += 1
    # Text occupies a small part of a highlighted word box, so a modest
    # ratio avoids rejecting a marker merely because it crosses dark glyphs.
    return samples > 0 and matches / samples >= 0.12


def _extract_raster_words(document: Any, items: list[dict[str, Any]], requested_color: tuple[int, int, int], tolerance: float, excluded_pages: set[int] | None = None) -> list[dict[str, Any]]:
    """Match MinerU word boxes against coloured pixels in image-only PDFs."""
    fitz = _fitz()
    excluded_pages = excluded_pages or set()
    pages: dict[int, list[dict[str, Any]]] = {}
    for item in items:
        page_number = _page_number(item.get("page"))
        if page_number not in excluded_pages:
            pages.setdefault(page_number, []).append(item)
    output = []
    for page_number, page_items in pages.items():
        if page_number > len(document):
            continue
        page = document[page_number - 1]
        pixmap = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), annots=True, alpha=False)
        for index, item in enumerate(page_items):
            box = _mineru_box_to_page(item, page)
            text = _text(item.get("text"))
            if not box or not text or not _pixmap_matches_color(pixmap, box, page, requested_color, tolerance):
                continue
            output.append({"page": page_number, "annotation_index": f"raster-{index}", "color": color_hex(requested_color), "text": text, "bbox": box, "words": [{"text": text, "bbox": box}], "raster": True})
    return output


def extract_highlighted_text(pdf_path: str | Path, requested_color: str | Sequence[float] | None = None, tolerance: float = DEFAULT_COLOR_TOLERANCE, mineru_items: list[dict[str, Any]] | None = None) -> list[dict[str, Any]]:
    """Extract one result per Highlight annotation matching ``requested_color``.

    Native PDF text is preferred because it has the same coordinate system as
    annotations. MinerU items are used for scanned PDFs where no text layer is
    available.
    """
    fitz = _fitz()
    document = fitz.open(str(Path(pdf_path).expanduser().resolve()))
    try:
        records = _annotation_records(document)
        color = parse_color(requested_color) if requested_color is not None else None
        native = _extract_native_words(document, records, color, tolerance)
        if not mineru_items or color is None:
            return native
        # A mixed PDF may have a text layer on some pages and image-only
        # pages on others. Keep native matches and fill only missing
        # annotations from MinerU, avoiding duplicate rows.
        native_keys = {(item["page"], item["annotation_index"]) for item in native}
        fallback = [item for item in _extract_mineru_items(mineru_items, records, color, tolerance) if (item["page"], item["annotation_index"]) not in native_keys]
        # Native annotation geometry is authoritative when annotations exist;
        # pixel classification is the fallback for flattened/scanned PDFs.
        raster = [] if records else _extract_raster_words(document, mineru_items, color, tolerance)
        return [*native, *fallback, *raster]
    finally:
        document.close()


def pick_color_from_pdf(pdf_path: str | Path, page_number: int = 1) -> str:
    """Open a Tk eyedropper and return the selected colour as ``#rrggbb``."""
    fitz = _fitz()
    try:
        import tkinter as tk
    except ImportError as exc:  # pragma: no cover - platform dependent
        raise RuntimeError("Tkinter is required for --eyedropper") from exc
    document = fitz.open(str(Path(pdf_path).expanduser().resolve()))
    try:
        if not 1 <= page_number <= len(document):
            raise ValueError(f"page must be between 1 and {len(document)}")
        page = document[page_number - 1]
        max_width = 1200.0
        scale = min(1.0, max_width / max(1.0, page.rect.width))
        pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), annots=True, alpha=False)
        try:
            root = tk.Tk()
        except tk.TclError as exc:  # pragma: no cover - depends on display
            raise RuntimeError("the eyedropper needs a graphical display") from exc
        root.title("PDF colour eyedropper - click a highlight")
        root.resizable(False, False)
        photo = tk.PhotoImage(data=pixmap.tobytes("ppm"))
        canvas = tk.Canvas(root, width=pixmap.width, height=pixmap.height, highlightthickness=0)
        canvas.create_image(0, 0, image=photo, anchor="nw")
        canvas.pack()
        tk.Label(root, text="Click a highlighted area to select its colour; Esc cancels.").pack(fill="x")
        selected: list[str] = []
        annotations = []
        for annotation in page.annots() or []:
            raw_type = getattr(annotation, "type", (None,))
            annotation_type = raw_type[0] if isinstance(raw_type, (tuple, list)) else raw_type
            if annotation_type == 8:
                annotations.append(annotation)

        def choose(event: Any) -> None:
            pdf_x, pdf_y = event.x / scale, event.y / scale
            for annotation in annotations:
                if any(box[0] <= pdf_x <= box[2] and box[1] <= pdf_y <= box[3] for box in _annotation_boxes(annotation)):
                    selected.append(color_hex(_annotation_color(annotation)))
                    root.destroy()
                    return
            pixel = pixmap.pixel(max(0, min(pixmap.width - 1, event.x)), max(0, min(pixmap.height - 1, event.y)))
            selected.append(color_hex(pixel[:3]))
            root.destroy()

        canvas.bind("<Button-1>", choose)
        root.bind("<Escape>", lambda _event: root.destroy())
        root.mainloop()
        if not selected:
            raise RuntimeError("colour selection cancelled")
        return selected[0]
    finally:
        document.close()


def process_pdf(pdf_path: str | Path, color: str | Sequence[float] | None = None, output_dir: str | Path | None = None, mineru_command: str | None = None, run_ocr: bool = True, tolerance: float = DEFAULT_COLOR_TOLERANCE) -> list[dict[str, Any]]:
    """Run MinerU and extract highlighted text from ``pdf_path``."""
    temporary: tempfile.TemporaryDirectory[str] | None = None
    if run_ocr:
        if output_dir is None:
            temporary = tempfile.TemporaryDirectory(prefix="noteflow-mineru-")
            output_dir = temporary.name
        try:
            run_mineru(pdf_path, output_dir, command=mineru_command)
            mineru_items = load_mineru_items(output_dir)
        finally:
            if temporary is not None:
                temporary.cleanup()
    else:
        mineru_items = []
    return extract_highlighted_text(pdf_path, color, tolerance=tolerance, mineru_items=mineru_items)


def _xlsx_cell(value: Any) -> str:
    """Create an inline-string or numeric OOXML cell without openpyxl."""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c t="n"><v>{escape(str(value))}</v></c>'
    return f'<c t="inlineStr"><is><t>{escape(str(value or ""))}</t></is></c>'


def _xlsx_sheet(rows: Sequence[Sequence[Any]]) -> str:
    body = []
    for row_number, row in enumerate(rows, 1):
        cells = "".join(_xlsx_cell(value) for value in row)
        body.append(f'<row r="{row_number}">{cells}</row>')
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + "".join(body) + "</sheetData></worksheet>"


def export_excel(results: Sequence[dict[str, Any]], output_path: str | Path) -> Path:
    """Write extraction results to a dependency-free Excel ``.xlsx`` file."""
    destination = Path(output_path).expanduser().resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    summary_rows: list[list[Any]] = [["index", "page", "text", "color", "word_count", "x0", "y0", "x1", "y1", "raster"]]
    word_rows: list[list[Any]] = [["highlight_index", "page", "word", "color", "x0", "y0", "x1", "y1"]]
    for index, result in enumerate(results, 1):
        box = _bbox(result.get("bbox")) or ["", "", "", ""]
        words = result.get("words") if isinstance(result.get("words"), list) else []
        summary_rows.append([index, result.get("page", ""), result.get("text", ""), result.get("color", ""), len(words) or 1, *box, "yes" if result.get("raster") else "no"])
        for word in words:
            word_box = _bbox(word.get("bbox")) if isinstance(word, dict) else None
            if not word_box:
                continue
            word_rows.append([index, result.get("page", ""), word.get("text", ""), result.get("color", ""), *word_box])
    content_types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
    root_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Highlights" sheetId="1" r:id="rId1"/><sheet name="Words" sheetId="2" r:id="rId2"/></sheets></workbook>'
    workbook_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>'
    with zipfile.ZipFile(destination, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        archive.writestr("xl/worksheets/sheet1.xml", _xlsx_sheet(summary_rows))
        archive.writestr("xl/worksheets/sheet2.xml", _xlsx_sheet(word_rows))
    return destination


def create_review_pdf(pdf_path: str | Path, results: Sequence[dict[str, Any]], output_path: str | Path) -> Path:
    """Draw numbered boxes around extracted text for visual review."""
    fitz = _fitz()
    source = Path(pdf_path).expanduser().resolve()
    destination = Path(output_path).expanduser().resolve()
    if source == destination:
        raise ValueError("review PDF output must be different from the source PDF")
    destination.parent.mkdir(parents=True, exist_ok=True)
    document = fitz.open(str(source))
    try:
        for index, result in enumerate(results, 1):
            page_number = _page_number(result.get("page"))
            if page_number > len(document):
                continue
            page = document[page_number - 1]
            words = result.get("words") if isinstance(result.get("words"), list) else []
            boxes = [_bbox(word.get("bbox")) for word in words if isinstance(word, dict)]
            boxes = [box for box in boxes if box]
            if not boxes:
                box = _bbox(result.get("bbox"))
                boxes = [box] if box else []
            for box in boxes:
                page.draw_rect(fitz.Rect(*box), color=(0.05, 0.55, 0.35), width=1.2, overlay=True)
            if boxes:
                label_box = boxes[0]
                page.insert_text((label_box[0], max(8, label_box[1] - 2)), str(index), fontsize=7, color=(0.05, 0.35, 0.22), overlay=True)
        document.save(str(destination), garbage=4, deflate=True)
    finally:
        document.close()
    return destination


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Extract text covered by PDF highlighter annotations.")
    parser.add_argument("pdf", type=Path, help="input PDF")
    parser.add_argument("--color", help="highlight colour: name, #rrggbb, rgb(r,g,b), or r,g,b")
    parser.add_argument("--tolerance", type=float, default=DEFAULT_COLOR_TOLERANCE, help="RGB matching tolerance (default: %(default)s)")
    parser.add_argument("--eyedropper", action="store_true", help="pick the highlight colour by clicking a PDF page")
    parser.add_argument("--eyedropper-page", type=int, default=1, help="page shown by the eyedropper (default: 1)")
    parser.add_argument("--output", type=Path, help="write JSON results to this file")
    parser.add_argument("--mineru-output", type=Path, help="keep MinerU JSON in this directory")
    parser.add_argument("--mineru-command", help="MinerU executable or command name")
    parser.add_argument("--no-mineru", action="store_true", help="skip MinerU and use the native PDF text layer")
    parser.add_argument("--excel", "--output-excel", dest="excel", type=Path, help="write an .xlsx workbook with highlight and word tables")
    parser.add_argument("--review-pdf", type=Path, help="write a copy of the PDF with boxes around extracted text")
    parser.add_argument("--review", action="store_true", help="write <pdf-stem>_review.pdf with boxes around extracted text")
    args = parser.parse_args(argv)
    try:
        color = args.color
        if args.eyedropper:
            color = pick_color_from_pdf(args.pdf, args.eyedropper_page)
            print(f"Selected highlight colour: {color}", file=sys.stderr)
        results = process_pdf(args.pdf, color=color, output_dir=args.mineru_output, mineru_command=args.mineru_command, run_ocr=not args.no_mineru, tolerance=max(0.0, args.tolerance))
        if args.excel:
            print(f"Excel written to: {export_excel(results, args.excel)}", file=sys.stderr)
        if args.review or args.review_pdf:
            review_path = args.review_pdf or args.pdf.with_name(f"{args.pdf.stem}_review.pdf")
            print(f"Review PDF written to: {create_review_pdf(args.pdf, results, review_path)}", file=sys.stderr)
    except (FileNotFoundError, RuntimeError, ValueError, OSError) as exc:
        parser.error(str(exc))
    payload = json.dumps(results, ensure_ascii=False, indent=2)
    if args.output:
        args.output.write_text(payload + "\n", encoding="utf-8")
    elif not (args.excel or args.review or args.review_pdf):
        print(payload)
    print(f"Extracted {len(results)} highlighted annotation(s).", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
