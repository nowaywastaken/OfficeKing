from __future__ import annotations

"""Configuration values for the activity scanner.

All runtime settings are sourced from the repository root `config.yml`.
This module provides a stable constant-like API for other packages while
loading the actual values at import time from YAML.
"""

import json
import logging
from pathlib import Path
from typing import Dict, List, Tuple

import yaml  # type: ignore[import-not-found]


# Root paths
CONFIG_ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = CONFIG_ROOT / "config.yml"


def _load_text(path: Path) -> str:
    """Load text using common encodings with BOM handling."""

    raw_bytes = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "gbk"):
        try:
            return raw_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError(
        "activity_scanner.config",
        raw_bytes,
        0,
        len(raw_bytes),
        "Unable to decode file using known encodings",
    )


def _load_yaml(path: Path) -> Dict:
    """Load YAML file into a Python dictionary."""

    if not path.exists():
        raise FileNotFoundError(f"Missing required config file: {path}")
    text = _load_text(path)
    data = yaml.safe_load(text) or {}
    if not isinstance(data, dict):
        raise ValueError("config.yml must contain a mapping at the top level")
    return data


def _coerce_str_list(value: object) -> List[str]:
    """Coerce a value to a list[str] (empty list if None)."""

    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        return [str(x) for x in value]
    return [str(value)]


def _load_student_roster(path: Path) -> Dict[str, str]:
    """Load the student roster from a JSON file.

    Returns an empty mapping if the file is missing or invalid. This keeps
    consumers robust when a roster is not needed for the current run.
    """

    try:
        if not path.exists():
            logging.debug("[config] Student roster not found: %s", path)
            return {}
        raw_text = _load_text(path)
        data = json.loads(raw_text)
        students = data.get("students", {})
        if not isinstance(students, dict):
            logging.warning("[config] 'students' must be a mapping in %s", path)
            return {}
        return {str(name): str(student_id) for name, student_id in students.items()}
    except Exception as exc:
        logging.warning("[config] Failed to load student roster %s: %s", path, exc)
        return {}


# Read YAML config once
_CFG = _load_yaml(CONFIG_PATH)

# Public values exported for other modules
DEFAULT_ACTIVITY_KEYWORDS: Tuple[str, ...] = tuple(
    _coerce_str_list(_CFG.get("default_activity_keywords"))
)
DEFAULT_CLASS_KEYWORDS: List[str] = _coerce_str_list(_CFG.get("default_class_keywords"))

OCR_SKIP_IF_VECTOR_TEXT: bool = bool(_CFG.get("ocr_skip_if_vector_text"))
OCR_VECTOR_TEXT_MIN_CHARS: int = int(_CFG.get("ocr_vector_text_min_chars"))
OCR_DPI: int = int(_CFG.get("ocr_dpi"))
OCR_USE_GPU: bool = bool(_CFG.get("ocr_use_gpu"))
OCR_LANG: str = str(_CFG.get("ocr_lang"))
OCR_MAX_SIDE: int = int(_CFG.get("ocr_max_side"))

# Additional application-level settings used by entrypoints
PDF_INPUT_PATHS: List[str] = _coerce_str_list(_CFG.get("pdf_paths"))
PDF_WORKERS: int = int(_CFG.get("workers"))
PDF_TIMEOUT_SEC: float = float(_CFG.get("timeout_sec"))
LOG_LEVEL: str = str(_CFG.get("log_level", "INFO"))

# Paths
STUDENT_ROSTER_PATH = (CONFIG_ROOT / str(_CFG.get("student_roster_path"))).resolve()
STUDENT_ID_MAP: Dict[str, str] = _load_student_roster(STUDENT_ROSTER_PATH)

# Cache dir for Office->Markdown conversions
MARKDOWN_CACHE_DIR = (CONFIG_ROOT / str(_CFG.get("markdown_cache_dir"))).resolve()

__all__ = [
    "CONFIG_ROOT",
    "CONFIG_PATH",
    "STUDENT_ROSTER_PATH",
    "STUDENT_ID_MAP",
    "DEFAULT_ACTIVITY_KEYWORDS",
    "DEFAULT_CLASS_KEYWORDS",
    "OCR_SKIP_IF_VECTOR_TEXT",
    "OCR_VECTOR_TEXT_MIN_CHARS",
    "OCR_DPI",
    "OCR_USE_GPU",
    "OCR_LANG",
    "OCR_MAX_SIDE",
    "PDF_INPUT_PATHS",
    "PDF_WORKERS",
    "PDF_TIMEOUT_SEC",
    "LOG_LEVEL",
    "MARKDOWN_CACHE_DIR",
]
