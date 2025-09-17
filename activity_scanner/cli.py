from __future__ import annotations

"""Command-line entry point for the activity scanner."""

import argparse
import logging
import os
import sys
from typing import List

from .config import DEFAULT_CLASS_KEYWORDS
from .document_scanner import (
    ScannableDocument,
    collect_supported_paths,
    derive_activity_title,
    scan_document_for_matches,
)
from .extractors import FAILED_PDFS, SUPPORTED_EXTENSIONS, open_failed_pdfs, read_text_from_path
from .report_builder import build_report_tables, write_report_workbook
from .roster_store import StudentDirectory
from .schema import (
    COLUMN_ACTIVITY_NAME,
    COLUMN_FILE_PATH,
    COLUMN_MATCH_COUNT,
    COLUMN_MATCH_TYPE,
    COLUMN_MATCH_VALUE,
    COLUMN_SNIPPET,
    COLUMN_STATUS,
    COLUMN_STUDENT_ID,
    COLUMN_STUDENT_NAME,
)

LOGGER = logging.getLogger(__name__)

DEFAULT_SCAN_ROOTS = ["input"]
DEFAULT_OUTPUT_FILE = "activity_hits.xlsx"


def parse_cli_arguments(argv: List[str] | None = None) -> argparse.Namespace:
    """Parse command-line arguments for the scanner."""

    parser = argparse.ArgumentParser(description="CSUST 班级观察活动资料扫描器")
    parser.add_argument(
        "--paths",
        nargs="+",
        help="待扫描的文件或文件夹路径，可填写多个；默认扫描 ./input",
    )
    parser.add_argument(
        "--out",
        default=DEFAULT_OUTPUT_FILE,
        help="输出 Excel 文件名，默认 activity_hits.xlsx",
    )
    parser.add_argument(
        "--class-tags",
        default=",".join(DEFAULT_CLASS_KEYWORDS),
        help="班级关键词，逗号分隔。默认使用 config.py 中的配置",
    )
    return parser.parse_args(argv)


def resolve_class_tags(raw_value: str | None) -> List[str]:
    """Normalise the raw class tag string into a list."""

    if raw_value is None or not raw_value.strip():
        return list(DEFAULT_CLASS_KEYWORDS)
    return [value.strip() for value in raw_value.split(",") if value.strip()]


def resolve_scan_paths(raw_paths: List[str] | None) -> List[str]:
    """Return the explicit paths or fall back to the default roots."""

    return raw_paths if raw_paths else list(DEFAULT_SCAN_ROOTS)


def run_cli(argv: List[str] | None = None) -> int:
    """Execute the scanner using command-line style arguments."""

    args = parse_cli_arguments(argv)

    try:
        roster = StudentDirectory.default()
    except Exception as exc:  # pragma: no cover - unexpected configuration failure
        print(f"初始化花名册失败: {exc}")
        return 1

    class_tags = resolve_class_tags(args.class_tags)
    targets = resolve_scan_paths(args.paths)

    files = collect_supported_paths(targets, SUPPORTED_EXTENSIONS)
    if not files:
        print("未找到可扫描的文件，支持 .docx .pdf .xlsx .xls .csv")
        return 4

    total = len(files)
    all_rows: List[dict] = []

    for index, path in enumerate(files, start=1):
        display_name = os.path.basename(path)
        print(f"[{index}/{total}] 正在扫描: {display_name}", flush=True)
        try:
            text = read_text_from_path(path)
        except ImportError as exc:
            missing = getattr(exc, "name", str(exc))
            print(f"缺少依赖: {missing}")
            return 1
        except Exception as exc:
            all_rows.append(
                {
                    COLUMN_FILE_PATH: path,
                    COLUMN_ACTIVITY_NAME: os.path.splitext(os.path.basename(path))[0],
                    COLUMN_STATUS: f"读取失败: {exc}",
                    COLUMN_MATCH_TYPE: "",
                    COLUMN_MATCH_VALUE: "",
                    COLUMN_STUDENT_ID: "",
                    COLUMN_STUDENT_NAME: "",
                    COLUMN_SNIPPET: "",
                    COLUMN_MATCH_COUNT: 0,
                }
            )
            continue

        document = ScannableDocument(
            path=path,
            text=text,
            activity=derive_activity_title(path, text),
        )
        try:
            rows = scan_document_for_matches(document, roster, class_tags)
        except Exception as exc:
            all_rows.append(
                {
                    COLUMN_FILE_PATH: path,
                    COLUMN_ACTIVITY_NAME: document.activity,
                    COLUMN_STATUS: f"扫描异常: {exc}",
                    COLUMN_MATCH_TYPE: "",
                    COLUMN_MATCH_VALUE: "",
                    COLUMN_STUDENT_ID: "",
                    COLUMN_STUDENT_NAME: "",
                    COLUMN_SNIPPET: "",
                    COLUMN_MATCH_COUNT: 0,
                }
            )
            continue
        all_rows.extend(rows)

    detail, per_activity, per_person, class_hits = build_report_tables(all_rows, roster)
    write_report_workbook(args.out, detail, per_activity, per_person, class_hits)

    if FAILED_PDFS:
        LOGGER.warning("[PDF] 以下文件文本提取失败，建议手动检查: %s", "、".join(sorted(FAILED_PDFS)))
        open_failed_pdfs(sorted(FAILED_PDFS))

    print(f"扫描完成，共处理 {total} 个文件，结果已写入 {args.out}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(run_cli())
