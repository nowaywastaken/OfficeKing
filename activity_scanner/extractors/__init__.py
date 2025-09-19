from __future__ import annotations

"""Dispatch utilities for file text extraction."""

import os
from pathlib import Path

from .pdf_reader import FAILED_PDFS, open_failed_pdfs, read_pdf_text
from .office_markdown import convert_office_to_markdown, cleanup_markdown_cache
from ..config import _load_text as _load_text_from_config  # type: ignore

SUPPORTED_EXTENSIONS = {
    ".pdf",
    # Word
    ".docx",
    ".doc",
    # Excel
    ".xlsx",
    ".xls",
    ".csv",
    # PowerPoint
    ".pptx",
    ".ppt",
    # Plain text
    ".txt",
    ".json",
}


def read_text_from_path(path: str) -> str:
    """Extract text from the given file based on its extension.

    For PDFs, combine vector/OCR text with MarkItDown's Markdown conversion
    to satisfy the requirement of processing a PDF with both strategies.
    """

    extension = os.path.splitext(path)[1].lower()
    if extension in {".docx", ".doc", ".xlsx", ".xls", ".csv", ".pptx", ".ppt"}:
        text, _ = convert_office_to_markdown(path)
        return text
    if extension == ".pdf":
        # Combine OCR + vector text with MarkItDown conversion
        base_text = read_pdf_text(path)
        md_text = ""
        try:
            md_text, _ = convert_office_to_markdown(path)
        except ImportError:
            # MarkItDown not installed; fall back to OCR/vector only
            md_text = ""
        except Exception:
            md_text = ""
        combined = "\n".join([t for t in [base_text, md_text] if t and t.strip()])
        return combined
    if extension in {".txt", ".json"}:
        try:
            return _load_text_from_config(Path(path))
        except Exception:
            try:
                return Path(path).read_text(encoding="utf-8", errors="ignore")
            except Exception:
                return ""
    return ""

__all__ = [
    "read_text_from_path",
    "FAILED_PDFS",
    "open_failed_pdfs",
    "SUPPORTED_EXTENSIONS",
    "cleanup_markdown_cache",
]
