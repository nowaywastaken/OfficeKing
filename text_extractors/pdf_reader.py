from __future__ import annotations

import sys

import logging
import os
import shutil
import subprocess
import tempfile
from typing import Iterable, Tuple

FAILED_PDFS: set[str] = set()


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


def _try_pdfplumber(path: str) -> str:
    import pdfplumber  # type: ignore[import-not-found]

    texts: list[str] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            try:
                txt = page.extract_text(x_tolerance=2, y_tolerance=3) or ""
            except Exception:
                txt = ""
            if txt.strip():
                texts.append(txt)
    return "\n".join(texts)


def _try_pymupdf(path: str) -> Tuple[str, bool, bool]:
    """Return extracted text, whether it looks image-heavy, and success flag."""
    try:
        import fitz  # type: ignore[import-not-found]
    except Exception:
        return "", False, False
    try:
        doc = fitz.open(path)
    except Exception:
        return "", False, False
    texts: list[str] = []
    image_pages = 0
    for page in doc:
        txt = page.get_text("text") or ""
        if not txt.strip():
            xobjs = page.get_images(full=True)
            drawings = page.get_drawings()
            if xobjs or drawings:
                image_pages += 1
        else:
            texts.append(txt)
    is_image_heavy = len(texts) == 0 and image_pages > 0
    return "\n".join(texts), is_image_heavy, True


def _ocr_with_ocrmypdf(src_pdf: str) -> str | None:
    """Attempt OCR via the ocrmypdf CLI, returning the OCR'd file path."""
    if shutil.which("ocrmypdf") is None:
        return None
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix="_ocr.pdf")
    tmp.close()
    cmd = [
        "ocrmypdf",
        "--force-ocr",
        "--skip-text",
        "--quiet",
        src_pdf,
        tmp.name,
    ]
    try:
        subprocess.run(cmd, check=True)
        return tmp.name
    except Exception:
        try:
            os.unlink(tmp.name)
        except Exception:
            pass
        return None


def _ocr_with_pytesseract(src_pdf: str) -> str:
    try:
        import fitz  # type: ignore[import-not-found]
        import pytesseract  # type: ignore[import-not-found]
        from PIL import Image  # type: ignore[import-not-found]
    except Exception:
        return ""
    try:
        doc = fitz.open(src_pdf)
    except Exception:
        return ""
    texts: list[str] = []
    for page in doc:
        pix = page.get_pixmap(dpi=200)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        try:
            txt = pytesseract.image_to_string(img, lang="chi_sim+eng")
        except Exception:
            txt = ""
        if txt.strip():
            texts.append(txt)
    return "\n".join(texts)


def extract_text_from_pdf(path: str) -> str:
    """Try several extraction strategies before falling back to OCR."""
    try:
        txt = _try_pdfplumber(path)
        if txt.strip():
            return txt
    except Exception as exc:
        logging.warning("[pdfplumber] extraction failed: %s", exc)

    txt2, image_heavy, ok = "", False, False
    try:
        txt2, image_heavy, ok = _try_pymupdf(path)
        if txt2.strip():
            return txt2
    except Exception as exc:
        logging.warning("[PyMuPDF] extraction failed: %s", exc)
        image_heavy, ok = False, False

    ocr_pdf = None
    if (not txt2.strip()) or image_heavy or (not ok):
        ocr_pdf = _ocr_with_ocrmypdf(path)
        if ocr_pdf:
            try:
                try:
                    t = _try_pdfplumber(ocr_pdf)
                    if t.strip():
                        return t
                except Exception:
                    pass
                t2, _, _ = _try_pymupdf(ocr_pdf)
                if t2.strip():
                    return t2
            finally:
                try:
                    os.unlink(ocr_pdf)
                except Exception:
                    pass

    try:
        t3 = _ocr_with_pytesseract(path)
        if t3.strip():
            return t3
    except Exception as exc:
        logging.warning("[pytesseract] OCR failed: %s", exc)

    FAILED_PDFS.add(path)
    logging.warning(
        "[PDF] Text extraction failed for %s; possible scans/images needing OCR.",
        path,
    )
    return ""
