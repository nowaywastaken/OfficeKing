#!/usr/bin/env python3
"""Convenience launcher that scans the ./input directory and writes results."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
import shutil
import os

from activity_scanner.cli import run_cli
from activity_scanner.concurrency import calibrate_pdf_workers, extract_pdfs_concurrently
from activity_scanner.extractors import FAILED_PDFS


def cleanup_paddlex_cache() -> None:
    """Delete PaddleX official models cache before each run (if present)."""

    # Prefer resolving from user home to be robust across machines
    home = Path(os.path.expanduser("~"))
    target = home / ".paddlex" / "official_models"
    # Also ensure the explicit path provided by the user is covered
    explicit = Path(r"C:\Users\17969\.paddlex\official_models")
    for folder in {target, explicit}:
        try:
            if folder.exists():
                shutil.rmtree(folder, ignore_errors=True)
        except Exception:
            # Ignore failures; continue with the run
            pass


def build_output_filename() -> str:
    """Create a timestamped Excel filename to avoid accidental overwrites."""

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"activity_hits_{timestamp}.xlsx"


def _gather_pdfs(root: Path) -> list[str]:
    """Recursively collect absolute PDF file paths under `root`."""

    pdfs: list[str] = []
    try:
        for r, _dirs, files in os.walk(root):
            for name in files:
                if os.path.splitext(name)[1].lower() == ".pdf":
                    pdfs.append(str((Path(r) / name).resolve()))
    except Exception:
        pass
    return sorted(set(pdfs))


def _enable_adaptive_pdf_cache(root: Path) -> None:
    """Pre-extract PDFs concurrently and patch the extractor to use cached text.

    - Calibrates worker count automatically (can override with env PDF_WORKERS)
    - Extracts PDF texts in parallel; records empty/failed as FAILED_PDFS
    - Monkey-patches `activity_scanner.extractors.read_pdf_text` to serve cache
    """

    pdfs = _gather_pdfs(root)
    if not pdfs:
        return

    try:
        workers = calibrate_pdf_workers()
    except Exception:
        workers = max(1, os.cpu_count() or 1)

    print(f"[PDF] 自适应并发预提取: {len(pdfs)} 个PDF（{workers} 进程）", flush=True)
    try:
        texts, errors = extract_pdfs_concurrently(pdfs, max_workers=workers)
    except ImportError as exc:
        missing = getattr(exc, "name", str(exc))
        print(f"[PDF] 依赖缺失，回退到串行: {missing}")
        return
    except Exception as exc:
        print(f"[PDF] 并发预提取失败，回退到串行: {exc}")
        return

    # Update FAILED_PDFS and patch the extractor to use cache
    try:
        FAILED_PDFS.update({p for p in pdfs if p in errors or not (texts.get(p, "").strip())})
    except Exception:
        pass

    try:
        import activity_scanner.extractors as ex

        _orig_read_pdf_text = ex.read_pdf_text

        cache = {str(Path(p).resolve()): (texts.get(p, "") or "") for p in pdfs}

        def _cached_read_pdf_text(path: str) -> str:  # type: ignore[override]
            key = str(Path(path).resolve())
            if key in cache:
                return cache[key]
            return _orig_read_pdf_text(path)

        ex.read_pdf_text = _cached_read_pdf_text  # type: ignore[assignment]
        print(f"[PDF] 预提取缓存就绪（命中 {sum(1 for v in cache.values() if v.strip())} / {len(cache)}）", flush=True)
    except Exception as exc:
        print(f"[PDF] 启用缓存失败，回退到串行: {exc}")
        return


def launch_scan() -> int:
    # Clean up PaddleX models cache as requested
    cleanup_paddlex_cache()

    root = Path("input")
    if not root.exists():
        print(f"未找到输入目录: {root}")
        return 1
    if not root.is_dir():
        print(f"输入路径不是文件夹: {root}")
        return 1
    if not any(root.iterdir()):
        print(f"输入目录为空: {root}")
        return 1

    # Adaptive mode can be disabled via env ADAPTIVE_PDF=0
    adaptive_enabled = os.environ.get("ADAPTIVE_PDF", "1") != "0"
    if adaptive_enabled:
        _enable_adaptive_pdf_cache(root)

    output = build_output_filename()
    args = ["--out", output]
    token = os.environ.get("SCAN_CONTAINS")
    if token:
        args += ["--contains", token]
    return run_cli(args)


if __name__ == "__main__":
    raise SystemExit(launch_scan())
