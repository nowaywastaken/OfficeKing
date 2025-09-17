from __future__ import annotations

"""Concurrency helpers for high‑throughput PDF text extraction.

This module provides:
- A lightweight CPU micro‑benchmark to estimate suitable worker counts
- A robust, timeout‑aware concurrent executor for extracting PDF text

Design notes:
- We parallelize across files (processes) because OCR and rendering are CPU‑bound
  and Python threads would contend on the GIL. PyMuPDF and Tesseract both release
  the GIL in C for heavy work, but combining Python‑level orchestration and OCR is
  still best handled by processes for predictable scaling across cores.
- We keep the public API small and self‑contained so callers in cli.py can use it
  without refactoring business logic.
"""

import os
import time
import math
import zlib
import random
from dataclasses import dataclass
from typing import Dict, Iterable, List, Tuple
from concurrent.futures import ProcessPoolExecutor, as_completed


# ------------------------------
# Worker calibration
# ------------------------------


def _cpu_probe_work(n_bytes: int = 8 * 1024 * 1024, loops: int = 200) -> int:
    """Do a chunk of CPU work (crc32) to simulate OCR‑like CPU intensity.

    Returns a simple integer score representing the amount of work done. This is
    intentionally deterministic within a call and fast (few hundred milliseconds).
    """

    rng = random.Random(42)
    # Create a fixed buffer; mutate a few bytes each loop to avoid trivial caching
    buf = bytearray(rng.getrandbits(8) for _ in range(n_bytes))
    score = 0
    for i in range(loops):
        # Flip a byte to vary input minimally and prevent warm cache tricks
        idx = (i * 2654435761) % n_bytes
        buf[idx] ^= (i & 0xFF)
        score ^= zlib.crc32(buf)
    return score


def _bench_workers_once(workers: int) -> float:
    """Measure aggregate throughput for a given worker count.

    We run identical CPU probes across processes and measure wall time. The score
    is (#tasks / elapsed). Higher is better.
    """

    if workers <= 0:
        return 0.0
    tasks = max(4, workers * 2)
    started = time.perf_counter()
    with ProcessPoolExecutor(max_workers=workers) as pool:
        futs = [pool.submit(_cpu_probe_work) for _ in range(tasks)]
        for f in as_completed(futs):
            # Ensure exceptions are surfaced in calibration
            _ = f.result()
    elapsed = max(1e-6, time.perf_counter() - started)
    return float(tasks) / elapsed


def calibrate_pdf_workers(max_cap: int | None = None) -> int:
    """Heuristically determine a good worker count for PDF extraction.

    Strategy:
    - Respect explicit overrides via env vars: `PDF_WORKERS` or `OCR_WORKERS`.
    - Try a few candidates around CPU count (½x, 1x, 1.5x) and pick the best.
    - Cap at a sensible upper bound to avoid thrashing on small systems.
    """

    # Environment overrides for power users
    for var in ("PDF_WORKERS", "OCR_WORKERS"):
        raw = os.environ.get(var)
        if raw and raw.isdigit():
            val = max(1, int(raw))
            return val

    cpu = os.cpu_count() or 1
    cap = max_cap if (isinstance(max_cap, int) and max_cap > 0) else max(4, min(24, cpu * 2))

    # Candidate set near CPU count; deduplicate and clamp to [1..cap]
    candidates = sorted({
        1,
        max(1, cpu // 2),
        cpu,
        min(cap, math.ceil(cpu * 1.5)),
    })

    best_workers = max(1, min(cap, cpu))
    best_score = -1.0
    for w in candidates:
        try:
            score = _bench_workers_once(w)
        except Exception:
            # Fall back to CPU count on any calibration issue
            score = 0.0
        if score > best_score:
            best_score = score
            best_workers = w

    return int(best_workers)


# ------------------------------
# Concurrent PDF extraction
# ------------------------------


@dataclass(frozen=True)
class PdfExtractionResult:
    path: str
    text: str
    error: str | None = None


def _extract_single_pdf(path: str) -> PdfExtractionResult:
    """Worker entry: extract text from a single PDF path.

    Runs in a separate process to leverage multiple CPU cores. Any exception is
    captured and returned as a message instead of propagating across processes.
    """

    try:
        from .extractors import pdf_reader
        text = pdf_reader.read_pdf_text(path)
        # Treat empty text as a soft failure so the parent can record it
        if not text.strip():
            return PdfExtractionResult(path=path, text=text, error="empty text")
        return PdfExtractionResult(path=path, text=text, error=None)
    except Exception as exc:  # pragma: no cover - depends on external libs/runtime
        return PdfExtractionResult(path=path, text="", error=f"{exc}")


def extract_pdfs_concurrently(
    paths: Iterable[str],
    *,
    max_workers: int | None = None,
    per_file_timeout_sec: float | None = 300.0,
) -> Tuple[Dict[str, str], Dict[str, str]]:
    """Extract text for many PDFs concurrently using processes.

    Returns a tuple: (texts, errors)
    - texts: mapping path -> extracted text (non‑empty may still be noisy OCR)
    - errors: mapping path -> error message (includes timeouts/empty text)

    Error handling:
    - Exceptions inside workers are captured and reported via `errors`.
    - If a task exceeds `per_file_timeout_sec`, it's marked as timeout.
    """

    pdf_paths = list(paths)
    if not pdf_paths:
        return {}, {}

    worker_count = calibrate_pdf_workers() if max_workers in (None, 0) else max(1, int(max_workers))

    texts: Dict[str, str] = {}
    errors: Dict[str, str] = {}

    # Submit all tasks then collect with timeout enforcement per future
    with ProcessPoolExecutor(max_workers=worker_count) as pool:
        futures = {pool.submit(_extract_single_pdf, p): p for p in pdf_paths}
        for fut in as_completed(futures, timeout=None):
            path = futures[fut]
            try:
                res = fut.result(timeout=per_file_timeout_sec)
            except Exception as exc:  # Timeout or worker crash
                errors[path] = f"worker error/timeout: {exc}"
                continue

            if res.error:
                errors[path] = res.error
            texts[path] = res.text or ""

    return texts, errors


__all__ = [
    "PdfExtractionResult",
    "calibrate_pdf_workers",
    "extract_pdfs_concurrently",
]

