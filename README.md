# OfficeKing 一键扫描（Office/PDF 转 Markdown + OCR）

最小化一键运行入口，所有参数集中在仓库根目录 `config.yml`。

- Word / Excel / PowerPoint（含老旧格式）→ MarkItDown 转 Markdown 文本
- PDF 同时进行“嵌入文本提取 + Tesseract OCR”，并与 MarkItDown 结果合并
- 搜索配置给定的学生名单（JSON：姓名→学号），命中则记录并导出 Excel
- 命中的源文件同时复制到输出目录中

运行命令：`python index.py`

日志：永远追加写入根目录 `log.txt`，并同步输出到终端（不按时间切分）。

## 主要函数说明（简要）

- `index.py:setup_logging(log_path, level_name)`
  - 配置终端与文件双通道日志；基于配置对第三方噪声 Warning 做过滤
- `index.py:_coerce_paths(values)`
  - 将 `config.yml` 的相对/绝对路径解析为存在的 `Path` 列表
- `index.py:collect_supported_files(inputs, exts)`
  - 递归收集受支持的文档类型（见 `SUPPORTED_EXTENSIONS`）
- `index.py:_load_config_values()`
  - 读取配置：`input_paths`、`output_root_dir`、`output_folder_format`、`report_filename`
- `index.py:_ensure_output_folder(root, folder_fmt)`
  - 按时间格式创建输出目录（支持 `strftime` 占位符）
- `index.py:_search_hits(text, name_to_id)`
  - 在文本中搜索姓名/学号（含“去中点”姓名变体）
- `index.py:main()`
  - 串联读取→提取→搜索→导出→复制命中文件的完整流程

辅助/底层：
- `activity_scanner/extractors/read_text_from_path(path)`
  - Office 交给 MarkItDown；PDF 合并“矢量/OCR 文本 + MarkItDown 转换”
- `activity_scanner/extractors/pdf_reader.py:read_pdf_text(path)`
  - PyMuPDF 渲染 + pytesseract OCR，并与向量文本合并
- `activity_scanner/extractors/office_markdown.py:convert_office_to_markdown(path)`
  - 使用 MarkItDown 转 Markdown，并写入缓存目录

## 配置参数（config.yml）

以下参数均需在 `config.yml` 中设置（提供 `config.example.yml` 供参考）。

- `input_paths` (list[str])
  - 必填。要扫描的文件/目录（递归）
- `output_root_dir` (str)
  - 必填。输出根目录，如 `Output`
- `output_folder_format` (str)
  - 必填。时间戳目录命名，支持 `strftime`，例：`%Y%m%d%H%M%S`
- `report_filename` (str)
  - 必填。导出的 Excel 文件名，需以 `.xlsx` 结尾，如 `report.xlsx`
- `student_roster_path` (str)
  - 必填。学生名册 JSON 文件路径，结构：`{"students": {"姓名": "学号"}}`
- `markdown_cache_dir` (str)
  - 必填。MarkItDown 生成的 Markdown 缓存目录（进程结束会清理）
- `log_level` (str)
  - 必填。日志等级：`DEBUG` / `INFO` / `WARNING` / `ERROR` / `CRITICAL`
- `log_suppressed_logger_prefixes` (list[str], 可选)
  - 将这些前缀的第三方 Logger 提升到 `ERROR`，用于屏蔽其 Warning（例如 `pdfminer`）
- `log_suppressed_message_contains` (list[str], 可选)
  - 若日志消息包含这些子串则直接丢弃（针对特定重复噪声，如 `FontBBox`）

OCR/Tesseract 相关：
- `ocr_skip_if_vector_text` (bool) 是否在向量文本足够多时跳过 OCR
- `ocr_vector_text_min_chars` (int) 跳过 OCR 的最少向量文本字符数
- `ocr_dpi` (int) 渲染 DPI，过大时会自动按 `ocr_max_side` 限制
- `ocr_use_gpu` (bool) 预留开关（OCR 走 CPU，图像处理可按需扩展）
- `ocr_lang` (str) 语言代码，支持复合如 `eng+chi_sim`
- `ocr_max_side` (int) 渲染图片最大边长上限

兼容旧入口（可选，不影响 index.py）：
- `pdf_paths`、`workers`、`timeout_sec` 供旧的并行 PDF 提取脚本使用

## 输出内容

- 目录：`<output_root_dir>/<output_folder_format>`（例如：`Output/20250101090000`）
- 文件：
  - `report.xlsx`：包含三列——`命中的内容`、`文件名`、`保存路径`
  - 命中的源文件：复制到同一目录下，若重名自动追加 `_2`、`_3` …

## 运行前置

- 安装依赖：见 `requirements.txt`
- Tesseract OCR 已安装并可执行（Windows 常见路径：`C:\\Program Files\\Tesseract-OCR\\tesseract.exe`），或将其加入 PATH/设置 `TESSERACT_CMD` 环境变量

## 其他说明

- 程序启动不接受命令行参数，所有参数从 `config.yml` 自动读取
- 代码中避免引入新的默认参数；所有可配置项均放入 `config.yml`
- 日志永远写入同一份 `log.txt`，同时打印到终端

