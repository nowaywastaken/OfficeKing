from __future__ import annotations

"""Office document to Markdown conversion via MarkItDown with caching.

Responsibilities:
- Convert Word/Excel/PowerPoint (including legacy formats) to Markdown
- Save converted Markdown into a cache folder
- Read back Markdown text and return it to callers
"""

import hashlib
import logging
import os
from pathlib import Path
from typing import Tuple

from ..config import MARKDOWN_CACHE_DIR
import atexit


LOGGER = logging.getLogger(__name__)


def _ensure_cache_dir() -> Path:
    MARKDOWN_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    return MARKDOWN_CACHE_DIR


def _hashed_filename(src_path: str) -> str:
    p = Path(src_path)
    stat = p.stat()
    key = f"{p.resolve()}|{stat.st_size}|{int(stat.st_mtime)}".encode("utf-8", errors="ignore")
    digest = hashlib.sha256(key).hexdigest()[:16]
    return f"{p.stem}.{digest}.md"


def convert_office_to_markdown(src_path: str) -> Tuple[str, Path]:
    """Convert an Office document to Markdown and cache the .md file.

    Returns a tuple: (markdown_text, cached_file_path)
    """

    try:
        from markitdown import MarkItDown  # type: ignore[import-not-found]
    except Exception as exc:  # pragma: no cover - environment dependent
        raise ImportError("markitdown") from exc

    cache_dir = _ensure_cache_dir()
    out_name = _hashed_filename(src_path)
    out_path = cache_dir / out_name

    LOGGER.info("[MarkItDown] Convert -> %s", src_path)
    md = MarkItDown()
    result = md.convert(src_path)

    text = (getattr(result, "text_content", None) or getattr(result, "markdown", "") or "")
    try:
        out_path.write_text(text, encoding="utf-8")
        LOGGER.info("[MarkItDown] Cached: %s", out_path)
    except Exception as exc:
        LOGGER.warning("[MarkItDown] Cache write failed %s: %s", out_path, exc)

    # Read back from the cache file to match the requested workflow strictly
    try:
        read_back = out_path.read_text(encoding="utf-8")
    except Exception:
        read_back = text
    return read_back, out_path


def cleanup_markdown_cache() -> None:
    """Delete the entire Markdown cache directory if it exists."""

    try:
        if MARKDOWN_CACHE_DIR.exists():
            for child in MARKDOWN_CACHE_DIR.glob("**/*"):
                try:
                    if child.is_file() or child.is_symlink():
                        child.unlink(missing_ok=True)
                except Exception:
                    pass
            # Remove directories bottom-up
            for child in sorted(MARKDOWN_CACHE_DIR.glob("**/*"), reverse=True):
                if child.is_dir():
                    try:
                        child.rmdir()
                    except Exception:
                        pass
            try:
                MARKDOWN_CACHE_DIR.rmdir()
            except Exception:
                pass
            LOGGER.info("[MarkItDown] Cache cleared: %s", MARKDOWN_CACHE_DIR)
    except Exception as exc:  # pragma: no cover - defensive
        LOGGER.warning("[MarkItDown] Cache cleanup issue: %s", exc)


# Ensure cache is cleared when the process finishes
atexit.register(cleanup_markdown_cache)
