from __future__ import annotations

"""High-level utilities for scanning student activity documents."""

from .cli import run_cli
from .report_builder import build_report_tables, write_report_workbook
from .document_scanner import (
    ScannableDocument,
    collect_supported_paths,
    derive_activity_title,
    find_occurrences,
    scan_document_for_matches,
)
from .roster_store import StudentDirectory

__all__ = [
    "run_cli",
    "build_report_tables",
    "write_report_workbook",
    "ScannableDocument",
    "collect_supported_paths",
    "derive_activity_title",
    "find_occurrences",
    "scan_document_for_matches",
    "StudentDirectory",
]
