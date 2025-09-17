from __future__ import annotations

import os
from typing import Iterable

from .docx_reader import extract_text_from_docx
from .excel_reader import extract_text_from_excel_like
from .pdf_reader import FAILED_PDFS, extract_text_from_pdf, open_failed_pdfs

SUPPORTED_EXTS = {".docx", ".pdf", ".xlsx", ".xls", ".csv"}


def extract_text(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return extract_text_from_docx(path)
    if ext == ".pdf":
        return extract_text_from_pdf(path)
    if ext in {".xlsx", ".xls", ".csv"}:
        return extract_text_from_excel_like(path)
    return ""

__all__ = [
    "extract_text",
    "FAILED_PDFS",
    "open_failed_pdfs",
    "SUPPORTED_EXTS",
]
