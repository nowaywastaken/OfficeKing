from __future__ import annotations

"""PDF text extraction using Tesseract OCR (pytesseract).

Workflow:
- Render each PDF page to an RGB numpy array via PyMuPDF (fitz)
- Recognize text with Tesseract via `pytesseract.image_to_string`
- Merge vector text (if present) with OCR results

Public API remains stable: `read_pdf_text`, `open_failed_pdfs`, `FAILED_PDFS`.
"""

import logging
import os
import subprocess
import sys
from typing import Iterable, List

from ..config import (
    OCR_DPI,
    OCR_LANG,
    OCR_SKIP_IF_VECTOR_TEXT,
    OCR_VECTOR_TEXT_MIN_CHARS,
    OCR_MAX_SIDE,
)

# Collect paths for which no text could be extracted, for post-run inspection
FAILED_PDFS: set[str] = set()

_OCR_INSTANCE = None  # Deprecated; retained for API stability only


def open_failed_pdfs(paths: Iterable[str]) -> None:
    """Open problematic PDFs for manual inspection when the script finishes."""

    for path in paths:
        try:
            if hasattr(os, "startfile"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as exc:
            logging.warning("[PDF] Failed to auto-open %s -> %s", path, exc)


def _resolve_tesseract_cmd() -> str | None:
    """Resolve a tesseract executable path and configure pytesseract if found.

    Returns the resolved command path if available, otherwise None.
    """

    try:
        import pytesseract  # type: ignore[import-not-found]
    except Exception:
        return None

    # 1) Respect environment variable if provided
    env_cmd = os.environ.get("TESSERACT_CMD")
    if env_cmd and os.path.exists(env_cmd):
        try:
            pytesseract.pytesseract.tesseract_cmd = env_cmd
            return env_cmd
        except Exception:
            pass

    # 2) Common Windows installation path
    win_default = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
    if os.name == "nt" and os.path.exists(win_default):
        try:
            pytesseract.pytesseract.tesseract_cmd = win_default
            return win_default
        except Exception:
            pass

    # 3) Leave as None to use PATH
    return None


def _lang_to_tesseract(lang: str) -> str:
    """Map project language codes to Tesseract language codes."""

    if not isinstance(lang, str) or not lang:
        return "eng"
    # Allow composite like "eng+chi_sim"
    parts = [p.strip() for p in lang.split("+") if p.strip()]
    mapped: list[str] = []
    for p in parts:
        key = p.lower()
        if key in {"ch", "zh", "zh_cn", "ch_sim", "zh_sim", "chi_sim"}:
            mapped.append("chi_sim")
        elif key in {"ch_tra", "zh_tra", "chi_tra"}:
            mapped.append("chi_tra")
        elif key in {"en", "eng"}:
            mapped.append("eng")
        else:
            # Pass-through for advanced users (e.g., 'osd', 'equ')
            mapped.append(p)
    return "+".join(dict.fromkeys(mapped)) or "eng"


def _render_page_to_array(page: "fitz.Page") -> "np.ndarray":
    """Render a fitz page to a numpy RGB image with size safeguards.

    Includes robust fallbacks for certain malformed PDFs (e.g. ExtGState issues)
    by disabling alpha and using a DisplayList-based rendering path when needed.
    """

    import numpy as np  # type: ignore[import-not-found]
    from PIL import Image  # type: ignore[import-not-found]

    dpi = int(OCR_DPI) if isinstance(OCR_DPI, int) and OCR_DPI > 0 else 150

    # Adjust DPI dynamically if page would exceed `OCR_MAX_SIDE`
    try:
        rect = getattr(page, "rect")
        width_px = float(rect.width) / 72.0 * dpi
        height_px = float(rect.height) / 72.0 * dpi
        max_px = max(width_px, height_px)
        cap = int(OCR_MAX_SIDE) if isinstance(OCR_MAX_SIDE, int) and OCR_MAX_SIDE > 0 else 4000
        if max_px > cap:
            scale = cap / max_px
            dpi = max(72, int(dpi * scale))
    except Exception:
        cap = int(OCR_MAX_SIDE) if isinstance(OCR_MAX_SIDE, int) and OCR_MAX_SIDE > 0 else 4000

    # Primary rendering path: disable alpha to avoid some ExtGState lookups
    try:
        pix = page.get_pixmap(dpi=dpi, alpha=False)
    except Exception as exc_primary:
        # Fallback: render via DisplayList with explicit matrix and colorspace
        try:
            import fitz  # type: ignore[import-not-found]
            zoom = dpi / 72.0
            matrix = fitz.Matrix(zoom, zoom)
            dlist = page.get_displaylist()
            pix = dlist.get_pixmap(matrix=matrix, alpha=False, colorspace=fitz.csRGB)
            logging.warning("[PDF] Page pixmap fallback used (alpha=False, displaylist): %s", exc_primary)
        except Exception as exc_dl:
            # Last try: reduce DPI to 96 and attempt again
            try:
                pix = page.get_pixmap(dpi=96, alpha=False)
                logging.warning("[PDF] Page pixmap fallback used (dpi=96): %s", exc_dl)
            except Exception as exc_last:
                # Give up: return a small blank image to keep pipeline moving
                logging.warning("[PDF] Page render failed; returning blank image: %s", exc_last)
                w = h = 256
                image = Image.new("RGB", (w, h), color=(255, 255, 255))
                return np.array(image)

    mode = "RGB" if pix.alpha == 0 else "RGBA"
    image = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
    if image.mode != "RGB":
        image = image.convert("RGB")

    # Final bounding to `OCR_MAX_SIDE` for safety
    cap = int(OCR_MAX_SIDE) if isinstance(OCR_MAX_SIDE, int) and OCR_MAX_SIDE > 0 else 4000
    if max(image.size) > cap:
        image.thumbnail((cap, cap), Image.LANCZOS)
    return np.array(image)


def _extract_vector_text(doc: "fitz.Document") -> List[str]:
    """Extract embedded (vector) text from the PDF when available."""

    blocks: List[str] = []
    for page in doc:
        try:
            txt = page.get_text("text") or ""
        except Exception as exc:
            logging.debug("[PDF] Vector extraction failed: %s", exc)
            txt = ""
        if txt.strip():
            blocks.append(txt.strip())
    return blocks


def _extract_with_ocr(doc: "fitz.Document") -> List[str]:
    """Run Tesseract on each page image and collect recognized line texts."""

    import numpy as np  # type: ignore[import-not-found]
    try:
        import pytesseract  # type: ignore[import-not-found]
    except ImportError as exc:
        raise exc

    # Configure tesseract command if possible
    _resolve_tesseract_cmd()

    tess_lang = _lang_to_tesseract(OCR_LANG)
    config = "--psm 6"

    lines: List[str] = []
    for page in doc:
        try:
            img: np.ndarray = _render_page_to_array(page)
        except Exception as exc:
            logging.warning("[Tesseract] Page render failed: %s", exc)
            continue
        try:
            text = pytesseract.image_to_string(img, lang=tess_lang, config=config)
        except Exception as exc:
            logging.warning("[Tesseract] OCR failed: %s", exc)
            continue

        try:
            for line in (text or "").splitlines():
                if line and line.strip():
                    lines.append(line.strip())
        except Exception:
            continue
    return lines


def read_pdf_text(path: str) -> str:
    """Extract text from a PDF by combining vector text and Tesseract results."""

    try:
        import fitz  # type: ignore[import-not-found]
    except ImportError as exc:
        raise exc

    try:
        with fitz.open(path) as doc:
            vector_blocks: List[str] = []
            ocr_lines: List[str] = []

            try:
                vector_blocks = _extract_vector_text(doc)
            except Exception as exc:
                logging.warning("[PDF] Vector text issue: %s", exc)

            vector_text = "\n".join(vector_blocks).strip()
            skip_ocr = bool(OCR_SKIP_IF_VECTOR_TEXT) and len(vector_text) >= int(OCR_VECTOR_TEXT_MIN_CHARS)

            if not skip_ocr:
                try:
                    ocr_lines = _extract_with_ocr(doc)
                except ImportError as exc:
                    raise exc
                except Exception as exc:
                    logging.warning("[Tesseract] OCR pipeline issue: %s", exc)

            parts: List[str] = []
            if vector_text:
                parts.append(vector_text)
            if ocr_lines:
                parts.append("\n".join(ocr_lines))

            combined = "\n".join(parts)
            if not combined.strip():
                FAILED_PDFS.add(path)
                logging.warning("[PDF] Text extraction failed for %s", path)
            return combined
    except Exception as exc:
        logging.warning("[PDF] Failed to open %s: %s", path, exc)
        FAILED_PDFS.add(path)
        return ""
