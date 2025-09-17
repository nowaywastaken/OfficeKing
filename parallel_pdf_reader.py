#!/usr/bin/env python3
"""Concurrent PDF text reader with adaptive worker calibration.

Usage examples:
  python parallel_pdf_reader.py --paths input docs --timeout 300
  python parallel_pdf_reader.py --paths file1.pdf file2.pdf --workers 8

This script focuses on high‑throughput concurrent PDF reading. It:
- Calibrates a good process count for CPU‑bound OCR/render workloads
- Extracts text from PDFs in parallel using the project PDF reader
- Surfaces errors, timeouts, and empty‑text conditions cleanly

Notes:
- Uses processes, not threads, to fully utilize multiple CPU cores
- Respects env overrides: PDF_WORKERS / OCR_WORKERS
- Safe defaults provided; tune with --workers/--timeout for special cases
"""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

from activity_scanner.concurrency import (
    calibrate_pdf_workers,
    extract_pdfs_concurrently,
)


def collect_pdf_paths(inputs: Iterable[str]) -> List[str]:
    """Collect all .pdf files from given files or directories recursively.

    - Ignores non‑existing paths with a warning printed to stdout
    - Follows directory trees and picks files ending with .pdf (case‑insensitive)
    """

    result: List[str] = []
    for raw in inputs:
        p = Path(raw)
        if p.is_file():
            if p.suffix.lower() == ".pdf":
                result.append(str(p.resolve()))
        elif p.is_dir():
            for found in p.rglob("*.pdf"):
                result.append(str(found.resolve()))
        else:
            print(f"[WARN] 未找到路径: {raw}")
    # Stable, deterministic order
    result = sorted(set(result))
    return result


def run(paths: List[str], workers: int | None, timeout: float | None) -> int:
    """Execute concurrent extraction and print a short JSON summary.

    Returns a shell exit code (0 on success, 1 on fatal error).
    """

    pdfs = collect_pdf_paths(paths)
    if not pdfs:
        print("[INFO] 未找到任何 PDF 文件。")
        return 0

    try:
        max_workers = workers if (workers and workers > 0) else calibrate_pdf_workers()
        print(f"[INFO] 准备并发提取 {len(pdfs)} 份PDF（{max_workers} 个进程）")
        texts, errors = extract_pdfs_concurrently(pdfs, max_workers=max_workers, per_file_timeout_sec=timeout)
    except ImportError as exc:
        missing = getattr(exc, "name", str(exc))
        print(f"[ERROR] 依赖缺失: {missing}")
        return 1
    except Exception as exc:
        print(f"[ERROR] 并发提取失败: {exc}")
        return 1

    ok_count = sum(1 for p in pdfs if texts.get(p, "").strip())
    err_count = sum(1 for p in pdfs if p in errors)

    # Emit a machine‑readable summary for downstream tooling/tests
    summary = {
        "total": len(pdfs),
        "success": ok_count,
        "errors": err_count,
        "error_details": errors,
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


def parse_args(argv: List[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="并发读取PDF文本（自适应进程数）")
    parser.add_argument("--paths", nargs="+", help="要扫描的PDF文件或目录（可多选）")
    parser.add_argument("--workers", type=int, default=0, help="固定进程数；默认自适应")
    parser.add_argument("--timeout", type=float, default=300.0, help="单文件超时（秒），默认300s")
    return parser.parse_args(argv)


if __name__ == "__main__":
    args = parse_args()
    code = run(args.paths, args.workers, args.timeout)
    raise SystemExit(code)

