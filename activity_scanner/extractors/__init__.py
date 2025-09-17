from __future__ import annotations

"""Dispatch utilities for file text extraction."""

import os

from .docx_reader import read_docx_text
from .excel_reader import read_spreadsheet_text
from .pdf_reader import FAILED_PDFS, open_failed_pdfs, read_pdf_text

SUPPORTED_EXTENSIONS = {".docx", ".pdf", ".xlsx", ".xls", ".csv"}


def read_text_from_path(path: str) -> str:
    """Extract text from the given file based on its extension."""

    extension = os.path.splitext(path)[1].lower()
    if extension == ".docx":
        return read_docx_text(path)
    if extension == ".pdf":
        return read_pdf_text(path)
    if extension in {".xlsx", ".xls", ".csv"}:
        return read_spreadsheet_text(path)
    return ""

__all__ = [
    "read_text_from_path",
    "FAILED_PDFS",
    "open_failed_pdfs",
    "SUPPORTED_EXTENSIONS",
]
