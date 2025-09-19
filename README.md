# OfficeKing 一键扫描（Office/PDF → Markdown + OCR）

最小可运行入口，所有参数集中在根目录 `config.yml`。支持：
- Word / Excel / PPT（含老旧格式）→ 通过 MarkItDown 转 Markdown 文本
- PDF → 同时进行“嵌入文本提取 + Tesseract OCR”，并与 MarkItDown 的结果合并
- 搜索给定的学生名单（JSON，姓名→学号）
- 命中文件导出 Excel，连同源文件复制到 `Output/时间戳这个写得有问题吗` 目录

运行：`python index.py`

日志：追加写入 `log.txt`，同时输出到终端（不按时间切分）。

## 主要函数说明（简要）

- `index.py:setup_logging(log_path, level_name)`
  - 控制台 + `log.txt`，单一追加文件，重复运行不重复分片。
- `index.py:_coerce_paths(values)`
  - 将 `config.yml` 的相对/绝对路径解析为存在的 `Path` 列表。
- `index.py:collect_supported_files(inputs, exts)`
  - 递归收集受支持的文档类型（见 `SUPPORTED_EXTENSIONS`）。
- `index.py:_load_config_values()`
  - 从 `config.yml` 读取：`input_paths`、`output_root_dir`、`output_folder_format`、`report_filename`。
- `index.py:_ensure_output_folder(root, folder_fmt)`
  - 依据时间格式创建输出目录（支持 `strftime` 占位符）。
- `index.py:_search_hits(text, name_to_id)`
  - 在文本中搜索姓名/学号（含“去中点”姓名变体）。
- `index.py:main()`
  - 串联读取→提取→搜索→导出→复制命中文件的完整流程。

辅助/底层：
- `activity_scanner/extractors/read_text_from_path(path)`
  - Office 文件交给 MarkItDown；PDF 合并“矢量/OCR 文本 + MarkItDown 转换”。
- `activity_scanner/extractors/pdf_reader.py:read_pdf_text(path)`
  - PyMuPDF 渲染 + pytesseract OCR，并与向量文本合并。
- `activity_scanner/extractors/office_markdown.py:convert_office_to_markdown(path)`
  - 统一用 MarkItDown 转 Markdown，缓存到 `markdown_cache_dir` 下。

## 配置参数（config.yml）

以下参数均需在 `config.yml` 中设置（提供 `config.example.yml` 供参考）：

- `input_paths` (list[str])
  - 必填。要扫描的文件/目录（递归）。
- `output_root_dir` (str)
  - 必填。输出根目录，如 `Output`。
- `output_folder_format` (str)
  - 必填。时间戳目录命名，支持 `strftime`，例如 `%Y%m%d%H%M%S这个写得有问题吗`。
- `report_filename` (str)
  - 必填。导出 Excel 文件名，需以 `.xlsx` 结尾，例如 `report.xlsx`。
- `student_roster_path` (str)
  - 必填。学生名册 JSON 路径，结构：`{"students": {"姓名": "学号"}}`。
- `markdown_cache_dir` (str)
  - 必填。MarkItDown 生成的 Markdown 缓存目录（进程结束会清理）。
- `log_level` (str)
  - 必填。日志等级：`DEBUG` / `INFO` / `WARNING` / `ERROR` / `CRITICAL`。
- 其余 OCR 参数（`ocr_*`）保留用于 PDF OCR 的图片渲染与语言设置。

兼容旧入口（可选，不影响 index.py）：
- `pdf_paths`、`workers`、`timeout_sec` 供旧的并行 PDF 提取脚本使用。

## 输出内容

- 目录：`<output_root_dir>/<output_folder_format>`（如：`Output/20250919093000这个写得有问题吗`）
- 文件：
  - `report.xlsx`：包含三列 —— `命中的内容`、`文件名`、`保存路径`
  - 命中的源文件：复制到同一目录下，若重名会自动按 `_2`、`_3` 追加

## 运行前置

- 安装依赖：见 `requirements.txt`
- Tesseract OCR 已安装并可执行（Windows 常见路径：`C:\\Program Files\\Tesseract-OCR\\tesseract.exe`），或将其加入 PATH/设定 `TESSERACT_CMD` 环境变量。

## 其他说明

- 程序启动不接受命令行参数，所有参数从 `config.yml` 自动读取。
- 代码内函数不保留默认参数；所有可配置项都转移到 `config.yml`。
- 日志永远写入同一个 `log.txt`，同时打印到终端。
