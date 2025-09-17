#!/usr/bin/env python3
"""Tesseract/pytesseract environment smoke test.

Usage examples:
  - Basic environment check (no OCR run):
      py test_paddle_env.py
  - Try to run a tiny OCR inference:
      py test_paddle_env.py --run-ocr
  - Specify tesseract.exe path explicitly:
      py test_paddle_env.py --tesseract-cmd "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
"""
from __future__ import annotations

import argparse
import os
import sys


def print_header(title: str) -> None:
    print(f"\n=== {title} ===")


def try_import_pytesseract(tess_cmd: str | None) -> dict:
    info: dict = {}
    try:
        import pytesseract  # type: ignore

        if tess_cmd:
            try:
                pytesseract.pytesseract.tesseract_cmd = tess_cmd
            except Exception as exc:
                info["set_cmd_error"] = str(exc)

        info["pytesseract_version"] = getattr(pytesseract, "__version__", "unknown")
        try:
            ver = pytesseract.get_tesseract_version()
            info["tesseract_version"] = str(ver)
        except Exception as exc:
            info["tesseract_version_error"] = str(exc)
    except Exception as exc:
        info["error"] = f"pytesseract import failed: {exc}"
    return info


def maybe_run_ocr(tess_cmd: str | None, lang: str) -> dict:
    out: dict = {}
    try:
        import pytesseract  # type: ignore
        from PIL import Image, ImageDraw  # type: ignore

        if tess_cmd:
            try:
                pytesseract.pytesseract.tesseract_cmd = tess_cmd
            except Exception:
                pass

        img = Image.new("RGB", (320, 120), "white")
        draw = ImageDraw.Draw(img)
        draw.text((10, 40), "TEST 123", fill="black")

        text = pytesseract.image_to_string(img, lang=lang or "eng", config="--psm 6")
        out["text"] = text.strip()
    except Exception as exc:
        out["error"] = f"OCR run failed: {exc}"
    return out


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Tesseract/pytesseract environment test")
    parser.add_argument("--run-ocr", action="store_true", help="Run a tiny OCR inference")
    parser.add_argument("--lang", default="eng", help="Language code for Tesseract (e.g., eng, chi_sim)")
    parser.add_argument("--tesseract-cmd", dest="tess_cmd", default=os.environ.get("TESSERACT_CMD", ""), help="Path to tesseract executable")
    args = parser.parse_args(argv)

    print_header("Python")
    print(sys.version)

    print_header("pytesseract")
    info = try_import_pytesseract(args.tess_cmd or None)
    for k, v in info.items():
        print(f"{k}: {v}")

    if "error" in info:
        return 1

    if args.run_ocr:
        print_header("Run OCR")
        ocr_out = maybe_run_ocr(args.tess_cmd or None, args.lang)
        for k, v in ocr_out.items():
            print(f"{k}: {v}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
