from __future__ import annotations

"""Command-line entry point for the activity scanner."""

import argparse
import logging
import os
import shutil
import sys
from pathlib import Path
from typing import List

from .config import DEFAULT_CLASS_KEYWORDS
from .document_scanner import (
    ScannableDocument,
    collect_supported_paths,
    derive_activity_title,
    scan_document_for_matches,
)
from .extractors import FAILED_PDFS, SUPPORTED_EXTENSIONS, read_text_from_path
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
ARCHIVE_ROOT = Path("scan_outputs")
SOURCE_SUBDIR = "source_files"


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


def _relative_to_targets(path: Path, targets: List[Path]) -> Path:
    """Build a relative path for archiving based on scan roots."""

    for root in targets:
        if root.is_file():
            if path == root:
                return Path(root.name)
        else:
            try:
                rel = path.relative_to(root)
                return Path(root.name) / rel
            except ValueError:
                continue
    return Path(path.name)


def bundle_run_artifacts(output_path: Path, files: List[Path], targets: List[Path]) -> Path:
    """Move the Excel output and copy source files into a bundle directory."""

    output_path = output_path.resolve()
    archive_root = ARCHIVE_ROOT.resolve()
    archive_root.mkdir(parents=True, exist_ok=True)

    run_folder = archive_root / output_path.stem
    suffix = 1
    while run_folder.exists():
        suffix += 1
        run_folder = archive_root / f"{output_path.stem}_{suffix}"
    run_folder.mkdir(parents=True, exist_ok=True)

    destination_output = run_folder / output_path.name
    if output_path.exists():
        shutil.move(str(output_path), destination_output)
    else:
        LOGGER.warning("预期的输出文件不存在: %s", output_path)

    sources_root = run_folder / SOURCE_SUBDIR
    for file_path in files:
        archive_rel = _relative_to_targets(file_path, targets)
        destination = sources_root / archive_rel
        destination.parent.mkdir(parents=True, exist_ok=True)
        try:
            shutil.copy2(file_path, destination)
        except Exception as exc:
            LOGGER.warning("复制源文件失败 %s -> %s: %s", file_path, destination, exc)

    return run_folder


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
    resolved_targets = [Path(target).resolve() for target in targets]

    files = collect_supported_paths(targets, SUPPORTED_EXTENSIONS)
    resolved_files = [Path(path).resolve() for path in files]
    if not resolved_files:
        print("未找到可扫描的文件，支持 .docx .pdf .xlsx .xls .csv")
        return 4

    total = len(resolved_files)
    all_rows: List[dict] = []
    matched_files: List[Path] = []

    for index, path in enumerate(resolved_files, start=1):
        print(f"[{index}/{total}] 正在扫描: {path.name}", flush=True)
        try:
            text = read_text_from_path(str(path))
        except ImportError as exc:
            missing = getattr(exc, "name", str(exc))
            print(f"缺少依赖: {missing}")
            return 1
        except Exception as exc:
            all_rows.append(
                {
                    COLUMN_FILE_PATH: str(path),
                    COLUMN_ACTIVITY_NAME: path.stem,
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
            path=str(path),
            text=text,
            activity=derive_activity_title(str(path), text),
        )
        try:
            rows = scan_document_for_matches(document, roster, class_tags)
        except Exception as exc:
            all_rows.append(
                {
                    COLUMN_FILE_PATH: str(path),
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

        # Track files that actually have matches (status == OK rows)
        if rows and any(r.get(COLUMN_STATUS) == "OK" for r in rows):
            matched_files.append(path)

    output_path = Path(args.out)
    detail, per_activity, per_person, class_hits = build_report_tables(all_rows, roster)
    write_report_workbook(str(output_path), detail, per_activity, per_person, class_hits)

    # Only copy matched files into the archive bundle
    archive_folder = bundle_run_artifacts(output_path, matched_files, resolved_targets)

    if FAILED_PDFS:
        LOGGER.warning("[PDF] 以下文件文本提取失败，建议手动检查: %s", "、".join(sorted(FAILED_PDFS)))
        

    print(f"扫描完成，共处理 {total} 个文件，结果已集中到 {archive_folder}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(run_cli())
