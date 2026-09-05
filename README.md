# PDF2Word

把 PDF 中标记的单词整理成可编辑的 Excel，再用本地 Qwen 模型补全词汇信息，最后生成可直接打开的学习卡片网页。

## 工作流

1. **提取高亮**：`pdftoword/test.py` 读取 PDF 原生高亮标注，输出 JSON/Excel，并可生成复核 PDF。
2. **补全词汇**：`qwentoword/main.py` 从 ModelScope 下载或复用本地 Qwen 模型，逐行补全释义、词性、词形、记忆方法和近义辨析。每行处理后原子保存，可中断后继续。
3. **生成学习页**：`qwentoword/html.py` 为工作簿添加稳定 ID 和音标，并生成独立的 `*.html` 文件。网页支持搜索、朗读、进度记录和懒加载。

## 安装

需要 Python 3.12 或更高版本。建议在仓库根目录创建虚拟环境，然后安装基础依赖和所需功能：

```bash
cd qwentoword
python -m venv .venv
source .venv/bin/activate       # Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
pip install pymupdf              # PDF 提取功能
```

`pdf` 功能需要 PyMuPDF；`ai` 功能需要 PyTorch、Transformers、Accelerate 和 ModelScope。根据显卡环境选择合适的 PyTorch 安装方式。

## 使用示例

从高亮 PDF 提取单词：

```bash
python pdftoword/test.py input.pdf --color '#ffff00' \\
  --excel output/highlights.xlsx --output output/highlights.json \\
  --review-pdf output/checked.pdf
```

用本地模型补全 Excel（默认更新输入文件；建议先复制一份）：

```bash
python qwentoword/main.py --input output/highlights.xlsx \\
  --output output/vocabulary.xlsx --model-cache-dir models
```

生成可离线打开的学习页：

```bash
python qwentoword/html.py --input output/vocabulary.xlsx \\
  --words qwentoword/words.json --output output/vocabulary.html
```

所有脚本均支持 `--help` 查看完整参数。模型下载、推理和 MinerU 都可能需要较多时间与磁盘空间。

## 仓库内容

| 路径 | 说明 |
| --- | --- |
| `pdftoword/test.py` | PDF 高亮提取与复核 |
| `qwentoword/main.py` | 本地 Qwen 词汇补全 |
| `qwentoword/html.py` | Excel 元数据更新与 HTML 生成 |
| `qwentoword/words.json` | 本地音标数据 |
| `qwentoword/考研生词.xlsx` | 示例词汇工作簿 |

## 说明

模型输出会经过严格 JSON 校验，但仍建议人工复核生成的释义。PDF、Excel 和网页示例文件属于项目样例；请确认其中的内容和授权符合你的发布范围。项目采用 MIT 许可证，详见 [LICENSE](LICENSE)。
