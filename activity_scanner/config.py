from __future__ import annotations

"""Configuration values for the activity scanner."""

import json
from pathlib import Path
from typing import Dict, List, Tuple

CONFIG_ROOT = Path(__file__).resolve().parent.parent
STUDENT_ROSTER_PATH = CONFIG_ROOT / "student_roster.json"


def _load_text(path: Path) -> str:
    """Load text using common encodings with BOM handling."""

    raw_bytes = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "gbk"):
        try:
            return raw_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue

    raise UnicodeDecodeError("activity_scanner.config", raw_bytes, 0, len(raw_bytes),
                              "Unable to decode file using known encodings")


def _load_student_roster(path: Path) -> Dict[str, str]:
    """Load the student roster from the JSON file."""

    if not path.exists():
        raise FileNotFoundError(f"Student roster JSON not found: {path}")

    raw_text = _load_text(path)
    try:
        data = json.loads(raw_text)
    except json.JSONDecodeError as exc:
        raise ValueError(f"Invalid JSON format in {path}: {exc}") from exc

    students = data.get("students", {})
    if not isinstance(students, dict):
        raise ValueError("The 'students' key must contain a mapping of name to ID.")

    return {str(name): str(student_id) for name, student_id in students.items()}


STUDENT_ID_MAP: Dict[str, str] = _load_student_roster(STUDENT_ROSTER_PATH)

DEFAULT_ACTIVITY_KEYWORDS: Tuple[str, ...] = ()

DEFAULT_CLASS_KEYWORDS: List[str] = []

OCR_SKIP_IF_VECTOR_TEXT: bool = False
OCR_VECTOR_TEXT_MIN_CHARS: int = 120
OCR_DPI: int = 150
OCR_USE_GPU: bool = True
OCR_LANG: str = "ch"
OCR_MAX_SIDE: int = 10000
__all__ = [
    "CONFIG_ROOT",
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
]




