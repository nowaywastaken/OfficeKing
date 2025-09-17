from __future__ import annotations

"""PDF extraction utilities built on PaddleOCR."""

import logging
import os
import subprocess
import sys
from typing import Iterable

FAILED_PDFS: set[str] = set()
_PADDLE_OCR = None


def open_failed_pdfs(paths: Iterable[str]) -> None:
    """Open problematic PDFs for manual inspection when the script finishes."""

    for path in paths:
        try:
            if hasattr(os, "startfile"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":  # type: ignore[name-defined]
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as exc:
            logging.warning("[PDF] Failed to auto-open %s -> %s", path, exc)


def _load_paddle() -> "PaddleOCR":
    """Instantiate PaddleOCR once and reuse it."""

    global _PADDLE_OCR
    if _PADDLE_OCR is None:
        from paddleocr import PaddleOCR  # type: ignore[import-not-found]

        _PADDLE_OCR = PaddleOCR(use_angle_cls=True, lang="ch")
    return _PADDLE_OCR


def _page_to_image_array(page: "fitz.Page") -> "np.ndarray":
    """Render a PDF page to a numpy array suitable for PaddleOCR."""

    import numpy as np  # type: ignore[import-not-found]
    from PIL import Image  # type: ignore[import-not-found]

    pixmap = page.get_pixmap(dpi=200)
    mode = "RGB" if pixmap.alpha == 0 else "RGBA"
    image = Image.frombytes(mode, (pixmap.width, pixmap.height), pixmap.samples)
    if image.mode == "RGBA":
        image = image.convert("RGB")
    return np.array(image)


def _extract_with_paddle(doc: "fitz.Document") -> list[str]:
    """Run PaddleOCR on every page and gather recognised lines."""

    ocr = _load_paddle()
    texts: list[str] = []
    for page in doc:
        try:
            image_array = _page_to_image_array(page)
        except Exception as exc:
            logging.warning("[PaddleOCR] Failed to render page: %s", exc)
            continue
        try:
            result = ocr.ocr(image_array, cls=True)
        except Exception as exc:
            logging.warning("[PaddleOCR] OCR failed: %s", exc)
            continue
        for item in result or []:
            if isinstance(item, list) and len(item) >= 2:
                text = item[1][0].strip()
                if text:
                    texts.append(text)
    return texts


def _extract_vector_text(doc: "fitz.Document") -> list[str]:
    """Collect textual content that already exists inside the PDF."""

    extracted: list[str] = []
    for page in doc:
        try:
            text = page.get_text("text") or ""
        except Exception as exc:
            logging.debug("[PDF] Failed to extract vector text: %s", exc)
            text = ""
        if text.strip():
            extracted.append(text.strip())
    return extracted


def read_pdf_text(path: str) -> str:
    """Extract text from a PDF using PaddleOCR plus native text if available."""

    try:
        import fitz  # type: ignore[import-not-found]
    except ImportError as exc:
        raise exc

    try:
        doc = fitz.open(path)
    except Exception as exc:
        logging.warning("[PDF] Failed to open %s: %s", path, exc)
        FAILED_PDFS.add(path)
        return ""

    vector_texts: list[str] = []
    ocr_texts: list[str] = []

    try:
        vector_texts = _extract_vector_text(doc)
    except Exception as exc:
        logging.warning("[PDF] Vector text extraction issue: %s", exc)

    try:
        ocr_texts = _extract_with_paddle(doc)
    except ImportError as exc:
        raise exc
    except Exception as exc:
        logging.warning("[PaddleOCR] OCR pipeline issue: %s", exc)

    combined_parts: list[str] = []
    if vector_texts:
        combined_parts.append("\n".join(vector_texts))
    if ocr_texts:
        combined_parts.append("\n".join(ocr_texts))

    combined_text = "\n".join(combined_parts)
    if not combined_text.strip():
        FAILED_PDFS.add(path)
        logging.warning("[PDF] Text extraction failed for %s", path)
    return combined_text
