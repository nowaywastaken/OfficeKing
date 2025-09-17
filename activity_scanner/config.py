from __future__ import annotations

"""Configuration values for the activity scanner."""

import json
from pathlib import Path
from typing import Dict, List, Tuple

CONFIG_ROOT = Path(__file__).resolve().parent.parent
STUDENT_ROSTER_PATH = CONFIG_ROOT / "student_roster.json"


def _load_text(path: Path) -> str:
    """Load text using UTF-8 first with graceful fallbacks."""

    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        try:
            return path.read_text(encoding="utf-8-sig")
        except UnicodeDecodeError:
            return path.read_text(encoding="gbk")


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

DEFAULT_ACTIVITY_KEYWORDS: Tuple[str, ...] = (
    "活动",
    "通知",
    "志愿",
    "比赛",
    "通告",
    "会议",
    "证明",
    "培训",
    "总结",
    "名单",
    "提示",
)

DEFAULT_CLASS_KEYWORDS: List[str] = [
    "高铁2401",
    "交通运输学院高铁2401班",
]

__all__ = [
    "CONFIG_ROOT",
    "STUDENT_ROSTER_PATH",
    "STUDENT_ID_MAP",
    "DEFAULT_ACTIVITY_KEYWORDS",
    "DEFAULT_CLASS_KEYWORDS",
]
