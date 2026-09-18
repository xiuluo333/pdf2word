# PDF2Word · 考研词汇学习卡片

将 Excel 词表交给本地 Qwen 模型补全，再生成可直接在浏览器打开的学习卡片。支持中文释义、英译英、音标、日常例句、点击朗读、搜索、需巩固标记、卡片翻面和学习位置记录。

当前仓库的可用流程从 **Excel 词表** 开始。旧版文档中的 `pdftoword/test.py` 已不在当前目录中，暂不提供 PDF 高亮提取命令；如果词汇来自 PDF，请先整理成下述 Excel 格式。

## 快速开始

### 1. 安装环境

建议使用 Python 3.12，并在项目根目录运行命令。首次使用模型需要联网下载依赖和模型；推理在本机执行，无需配置远程模型 API Key。

Linux / macOS：

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -U pip
python -m pip install -U -r requirements.txt
```

Windows PowerShell：

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -U pip
python -m pip install -U -r requirements.txt
```

`requirements.txt` 包含 ModelScope、Transformers、Accelerate、PyTorch、openpyxl 和 PyMuPDF。当前 Excel → 网页流程不使用 PyMuPDF。使用 NVIDIA GPU 时，请安装与本机驱动兼容的 PyTorch 版本，可参考 [PyTorch 安装说明](https://pytorch.org/get-started/locally/)。

默认模型是 `Qwen/Qwen3.5-9B`，需要较多内存或显存及磁盘空间，实际占用取决于模型和运行配置。`auto` 优先使用可用的 CUDA，否则使用 CPU；CPU 推理通常较慢。当前自动设备选择不包含 Apple MPS。

### 2. 准备 Excel

使用 `.xlsx` 文件，第一行为表头，每行一个单词。供 `main.py` 使用的最小词表只需要 `word` 列，`number` 可选：

| word | number |
| --- | --- |
| voyage | 12 |
| passage | 20 |

请使用下方字段表中的英文表头，大小写及首尾空格不影响识别。缺少的生成字段会自动补列，空白单词行会跳过。默认读取活动工作表，也可通过 `--sheet` 指定。

### 3. 补全词汇并生成页面

使用仓库自带的 `考研生词.xlsx`：

```bash
python main.py
```

默认生成原表同目录下的两个文件：

- `考研生词_processed.xlsx`：补全后的词表。
- `考研生词_processed.html`：学习页面，双击即可打开。

默认模型缓存位于项目目录的 `models/`。原始输入文件默认保留；如果将 `--output` 指向输入文件本身，则会直接更新该文件。

使用自己的词表，并指定输出位置：

```bash
python main.py --input "my_words.xlsx" --output "output/vocabulary.xlsx" --words "words.json" --model-cache-dir "models"
```

以上命令会生成 `output/vocabulary.html`。示例使用单行命令，可直接用于 Bash 或 PowerShell。

## 字段与处理规则

| Excel 表头 | 模型字段 / 用途 | 内容 |
| --- | --- | --- |
| `word` | 输入 | 英文单词，必需 |
| `number` | 输入 | 可选的词频或编号，仅作参考，不参与自动排序 |
| `id` | 程序生成 | 卡片编号，如 `W000001`；保留已有非重复编号 |
| `translate` | `translate` | 中文释义 |
| `English definition` | `english_definition` | 用英文解释单词核心含义 |
| `phonetic` | `phonetic` | IPA 音标；模型补全时要求英式 `/.../` 格式 |
| `part of speech` | `part_of_speech` | 词性，如 `n.; v.` |
| `Transformation` | `transformation` | 词形变化 |
| `Memory techniques` | `memory_techniques` | 记忆方法 |
| `Distinguishing between similar words` | `distinguishing_between_similar_words` | 近义词或易混淆词辨析 |
| `example sentence` | `example_sentence` | 简短日常英语例句，可附句末括号说明 |

处理顺序如下：

1. 首次运行时，将输入表复制到输出路径。输出表已存在时直接继续处理该表，不重新复制输入表。
2. 添加缺失或重复的卡片编号，从本地 `words.json` 回填缺失音标，并用词典中的中文释义更新 `translate`。
3. 模型补全剩余空白字段，包括英译英、例句和仍缺失的音标。模型返回全部 8 个字段，但只将结果写入原本为空的单元格。
4. 每行成功后保存工作簿；全部处理结束后生成 HTML，除非指定 `--no-html`。

**中文释义以本地词典为准，可能覆盖已有 `translate`。** 音标只在为空时回填，已有音标不会自动转换成英式。页面默认英式朗读与表中音标的数据来源是两回事。

空白指空单元格或只有空格的内容；`无`、`待补充` 等文字不算空白。需要重新生成某项时，清空输出工作簿中对应单元格后再运行。

模型返回值会经过 JSON 字段、非空字符串、英文释义语言和音标格式检查；这些检查不能保证释义或 IPA 内容完全正确，学习时仍可结合词典复核。

### 中断、重试与续跑

按 `Ctrl+C` 中断后，已成功保存的行会保留。重新执行相同命令，使用同一个 `--output` 即可补全剩余空白字段。旧版输出缺少英译英等新列时，也会自动补列并处理。

默认每个单词失败后最多重试 3 次，即包含首次请求最多尝试 4 次。持续失败的行会跳过并在终端报告；其他行继续处理。结束日志中的 `failed after retries` 可用于判断是否仍有失败项。

如果修改了原始输入表、增加了单词，已有输出表不会自动合并这些修改。可选择新的 `--output` 从新版输入开始处理，或直接编辑已有输出表后续跑。

中断时 HTML 可能尚未生成或仍为旧版本，可用下节命令导出已保存的数据。

## 只生成或刷新网页

已有完整词表，或只想查看当前已处理结果时，运行：

```bash
python generate_html.py --input "考研生词_processed.xlsx" --words "words.json" --output "考研生词_processed.html"
```

此命令不加载 Qwen，也不会生成缺失的英译英或例句；空白字段显示“待补充”。HTML 内嵌词表数据，修改 Excel 后需要重新生成并刷新浏览器。

**生成器也会更新输入 Excel**：添加编号、回填缺失音标，并按词典更新中文释义。需要保留原表时，先复制工作簿再运行。

单独使用生成器不需要安装模型依赖。它优先使用 openpyxl；未安装时，会用 Python 标准库解析普通 `.xlsx` 词表。生成器还接受部分中文表头，例如 `单词`、`音标`、`英译英`、`英文释义`、`日常例句`；要同时用于 `main.py`，请统一使用上表中的英文表头。

### 本地词典

`--words` 指定本地词典 JSON。默认查找输入工作簿旁边的 `words.json`，自定义输入目录时建议显式传入路径。

支持如下列表格式：

```json
[
  {"word": "voyage", "phonetic": "/ˈvɔɪɪdʒ/", "meaning": "n. 航行；航程"}
]
```

释义字段也接受 `translate` 或 `translation`。仅提供音标时，也支持 `{"voyage": "/ˈvɔɪɪdʒ/"}`。匹配单词时忽略大小写。词典文件不存在时跳过词典回填；`main.py` 仍可调用模型补全空白字段。`KaoYan_2.json` 目前没有被脚本自动加载。

## 网页使用

- **单词朗读**：点击扬声器按钮；点击音标以更慢语速朗读单词。
- **英译英与日常例句**：点击文字即可朗读。末尾半角 `(...)` 或全角 `（...）` 括号补充内容不朗读，页面保留完整文字。支持连续多个末尾括号；句中括号内容保留在朗读中。
- **口音切换**：默认英式 `English (UK)`，可切换美式；选择会保存在当前浏览器中，下次打开恢复。
- **搜索**：输入单词、中文释义、词性、英译英或例句内容来筛选词卡；点击候选项或按回车定位。未找到时可转到有道词典，按 `Esc` 关闭候选列表。
- **查词**：双击单词标题打开有道词典，需要网络连接。
- **学习位置**：点击词卡、标记或翻页时自动保存学习位置，再次打开恢复到对应页，也可点击“继续上次学习”定位到具体词卡。旧版位置记录首次打开时会尝试迁移。
- **需巩固标记**：点击词卡底部“标记需巩固”开关，再次点击可取消。顶部“只看需巩固”可与搜索组合使用；取消当前标记后，该词会立即从筛选列表移除。
- **卡片翻面**：点击“翻到背面”，背面只显示单词和日常例句；点击“查看释义”返回。朗读和标记不会触发翻面，翻面状态仅在本次打开期间保留。
- **分页**：每页最多 36 张卡片，搜索和恢复学习位置直接跳转相应页，避免 8000 词时加载全部卡片。词卡保持工作簿顺序，若希望按词频学习，请先在 Excel 中排序。

### 学习记录保存与备份

标记和位置使用浏览器 `localStorage` 持久保存，每次改动立即逐条写入，并保留一份备用记录。主记录损坏时会尝试读取备用记录；保存失败时页面会明确提示，此时请在关闭页面前导出备份。

记录按规范化后的单词文字匹配，不依赖 Excel 行号或卡片编号，因此重新排序、重新生成 HTML 后仍可匹配。相同拼写的重复词条共享需巩固标记，不同词表中的同名单词在同一浏览器存储范围内也共享标记。更改单词拼写后会被视为新词。

- 点击“导出学习记录”，下载 JSON 文件，包含标记（含取消标记记录）与当前位置，不包含整份词表。
- 点击“导入学习记录”，选择此前导出的 JSON 文件。导入会先检查整个文件，再按单词合并；时间戳较新的标记优先，旧备份不会重新启用已取消的旧标记。跨设备导入前应确保设备时间正确。
- 同一网站存储范围内，多个标签页通过浏览器存储事件同步标记；每个单词独立写入，避免整个词表记录相互覆盖。

本地保存并非云端同步，备用记录也不能抵御浏览器清理、卸载或设备损坏。请定期导出 JSON，换浏览器、换设备或移动 HTML 文件前先备份。直接以 `file://` 打开时，不同浏览器对文件存储的支持和路径隔离方式可能不同；移动或重命名 HTML 后若记录不可见，可导入备份恢复。

若希望使用固定地址，可在项目目录运行 `python -m http.server 8000 --bind 127.0.0.1`，在浏览器打开 `http://127.0.0.1:8000/` 后选择学习页。继续使用相同浏览器、主机地址和端口；从本地文件切换到此地址时，先导出再导入记录。

页面内容可离线查看。朗读使用浏览器和系统提供的语音，实际口音及离线可用性取决于已安装的语音包；需要相应英语语音才能获得预期效果。

## 命令参数

### `main.py`

| 参数 | 默认值 | 说明 |
| --- | --- | --- |
| `--input` | 项目目录 `考研生词.xlsx` | 原始输入表 |
| `--output` | 输入表同目录 `<文件名>_processed.xlsx` | 输出及续跑工作簿 |
| `--sheet` | 活动工作表 | 指定工作表名称 |
| `--words` | 输入表旁的 `words.json` | 本地音标和释义词典 |
| `--model` | `Qwen/Qwen3.5-9B` | ModelScope 模型 ID，需与当前加载代码兼容 |
| `--model-cache-dir` / `--model-dir` | 项目目录 `models/` | ModelScope 缓存目录，不是直接指定已解压权重的入口 |
| `--device` | `auto` | 可用 `cpu`、`cuda`、`cuda:0` 等 |
| `--max-new-tokens` | `1200` | 每个单词的最大生成 token 数 |
| `--request-interval` | `0.5` | 每行成功后的等待秒数 |
| `--max-retries` | `3` | 单词失败后的重试次数 |
| `--html-output` | 输出工作簿同名 `.html` | 自定义网页路径 |
| `--no-html` | 不启用 | 只处理 Excel，不生成网页 |

只更新 Excel：

```bash
python main.py --input "my_words.xlsx" --output "output/vocabulary.xlsx" --words "words.json" --no-html
```

指定 CPU 并增加输出长度上限：

```bash
python main.py --device cpu --max-new-tokens 1800
```

### `generate_html.py`

| 参数 | 默认值 | 说明 |
| --- | --- | --- |
| `--input` | 项目目录 `考研生词.xlsx` | 要更新并导出的工作簿 |
| `--output` | 输入工作簿同名 `.html` | 网页路径 |
| `--sheet` | 活动工作表 | 指定工作表名称 |
| `--words` | 输入表旁的 `words.json` | 本地词典路径 |

两个脚本均支持 `--help`。相对路径以运行命令时的当前目录为准。

### 环境变量

以下环境变量可设置默认值，显式命令行参数优先：

| 环境变量 | 对应参数 | 适用脚本 |
| --- | --- | --- |
| `VOCAB_INPUT` | `--input` | 两个脚本 |
| `MODELSCOPE_MODEL` | `--model` | `main.py` |
| `MODELSCOPE_CACHE_DIR` | `--model-cache-dir` | `main.py` |
| `MODEL_DEVICE` | `--device` | `main.py` |
| `MODEL_MAX_NEW_TOKENS` | `--max-new-tokens` | `main.py` |
| `MODELSCOPE_REQUEST_INTERVAL` | `--request-interval` | `main.py` |

## 常见问题

| 现象 | 处理方式 |
| --- | --- |
| 提示缺少 `word` 列 | 确认第一行有英文 `word` 表头，并检查 `--sheet` 是否选对工作表。 |
| 输出表没有反映原表的新修改 | 已存在输出用于续跑，不重新复制输入；使用新输出路径或更新输出工作簿。 |
| 英译英、例句仍显示“待补充” | 先运行 `main.py` 补全，检查失败日志，再打开新生成的 `_processed.html`；单独生成 HTML 不调用模型。 |
| 缺少包、共享库错误或依赖导入失败 | 在同一虚拟环境执行 `python -m pip install -U -r requirements.txt`，根据日志原始错误检查 PyTorch、Transformers 版本及驱动。程序会输出所用 Python 路径。 |
| 模型类型不受支持 | 依赖文件是版本下限，可能不足以支持默认模型；更新 Transformers / ModelScope，并核对模型要求。 |
| 下载失败 | 检查到 ModelScope 的网络连接、磁盘空间及缓存目录权限；保留已有缓存后重试。 |
| CUDA 不可用或显存不足 | 检查 PyTorch CUDA 支持及驱动，关闭占用显存的程序，或指定 `--device cpu` 并确保系统内存足够。 |
| JSON 校验失败、输出截断 | 查看该词重试日志；可适当增加 `--max-new-tokens`，重新运行补全失败行。 |
| 音标仍是美式 | 已有或本地词典提供的音标会保留；模型只对剩余空白音标按英式要求补全。切换网页口音不会修改音标文字。 |
| 点击朗读无声或口音不符 | 确认浏览器支持语音合成、系统已安装对应英语语音，并检查音量；可尝试其他浏览器。 |
| 保存 Excel 失败 | 关闭 Excel 中正在占用该文件的窗口，检查输出目录写权限。 |

## 项目结构与检查

| 路径 | 用途 |
| --- | --- |
| `main.py` | 模型提示词、结果校验、Excel 补全与续跑入口 |
| `generate_html.py` | 工作簿元数据更新、HTML 模板和页面交互 |
| `words.json` | 默认本地音标与中文释义词典 |
| `考研生词.xlsx` / `考研生词.html` | 示例词表与学习页 |
| `tests/test_vocabulary.py` | 字段、旧表补列和续跑逻辑测试 |
| `tests/reader.mjs` | 页面渲染与朗读调用测试 |
| `tests/study-browser.mjs` | 8000 词、翻面、筛选、存储恢复、备份合并与移动布局的浏览器测试 |

修改页面交互时，应同步更新 `generate_html.py` 的模板和示例 HTML，避免重新生成时丢失修改。

运行已有检查：

```bash
python -m unittest discover -s tests -v
node tests/reader.mjs
python -m py_compile main.py generate_html.py
```

测试无需下载模型；页面脚本测试需要 Node.js。它使用模拟语音接口验证朗读内容与参数，不能代替浏览器中的实际发声检查，也不验证模型生成质量。

浏览器集成测试需要先使用独立测试配置启动 Chrome，例如 Linux：

```bash
google-chrome --headless=new --disable-gpu --remote-debugging-port=9224 --user-data-dir=/tmp/vocab-study-test-browser about:blank
```

在另一终端运行 `node tests/study-browser.mjs`。测试会在临时目录生成 8000 词页面并修改该测试页面的存储，因此请使用独立的浏览器测试配置。测试结束后关闭该 Chrome 进程。

## 许可证

项目代码采用 MIT 许可证，详见 [LICENSE](LICENSE)。
