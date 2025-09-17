from __future__ import annotations

"""Core logic for scanning documents and collecting matches."""

import os
from dataclasses import dataclass
from typing import Dict, Iterable, List

from .config import DEFAULT_ACTIVITY_KEYWORDS
from .roster_store import StudentDirectory
from .schema import (
    CLASS_TAG_LABEL,
    COLUMN_ACTIVITY_NAME,
    COLUMN_FILE_PATH,
    COLUMN_MATCH_COUNT,
    COLUMN_MATCH_TYPE,
    COLUMN_MATCH_VALUE,
    COLUMN_SNIPPET,
    COLUMN_STATUS,
    COLUMN_STUDENT_ID,
    COLUMN_STUDENT_NAME,
)


@dataclass(frozen=True)
class ScannableDocument:
    """Lightweight representation of a file being scanned."""

    path: str
    text: str
    activity: str


def collect_supported_paths(inputs: Iterable[str], supported_extensions: Iterable[str]) -> List[str]:
    """Collect every supported file located within *inputs*."""

    supported = {ext.lower() for ext in supported_extensions}
    result: List[str] = []
    for raw_path in inputs:
        if os.path.isfile(raw_path):
            if os.path.splitext(raw_path)[1].lower() in supported:
                result.append(raw_path)
        elif os.path.isdir(raw_path):
            for root, _, filenames in os.walk(raw_path):
                for filename in filenames:
                    if os.path.splitext(filename)[1].lower() in supported:
                        result.append(os.path.join(root, filename))
        else:
            print(f"[提醒] 未找到路径: {raw_path}")
    return result


def find_occurrences(text: str, token: str) -> List[int]:
    """Return every index where *token* occurs in *text*."""

    if not token:
        return []
    indices: List[int] = []
    start = 0
    while True:
        index = text.find(token, start)
        if index == -1:
            break
        indices.append(index)
        start = index + len(token)
    return indices


def get_text_context(text: str, index: int, length: int, span: int = 20) -> str:
    """Return text surrounding a match, with newlines replaced by spaces."""

    start = max(0, index - span)
    end = min(len(text), index + length + span)
    return text[start:end].replace("\n", " ")


def derive_activity_title(path: str, text: str) -> str:
    """Derive a short title for the activity based on file text or path."""

    first_non_empty_line = ""
    for line in text.splitlines():
        candidate = line.strip()
        if candidate:
            first_non_empty_line = candidate
            break
    filename = os.path.splitext(os.path.basename(path))[0]
    if any(keyword in first_non_empty_line for keyword in DEFAULT_ACTIVITY_KEYWORDS):
        title = first_non_empty_line[:80]
    else:
        title = filename[:80]
    parent_directory = os.path.basename(os.path.dirname(path))
    if parent_directory and parent_directory not in title and len(parent_directory) <= 40:
        title = f"{title}（{parent_directory}）"
    return title


def create_empty_result_row(path: str, activity: str) -> Dict[str, object]:
    """Create a placeholder row when no hit is detected."""

    return {
        COLUMN_FILE_PATH: path,
        COLUMN_ACTIVITY_NAME: activity,
        COLUMN_STATUS: "未命中",
        COLUMN_MATCH_TYPE: "",
        COLUMN_MATCH_VALUE: "",
        COLUMN_STUDENT_ID: "",
        COLUMN_STUDENT_NAME: "",
        COLUMN_SNIPPET: "",
        COLUMN_MATCH_COUNT: 0,
    }


def scan_document_for_matches(document: ScannableDocument, roster: StudentDirectory, class_tags: List[str]) -> List[Dict[str, object]]:
    """Scan a document for student identifiers and class-tag keywords."""

    text = document.text
    collected_rows: List[Dict[str, object]] = []
    student_hits: Dict[str, Dict[str, object]] = {}

    def ensure_student_bucket(key: str) -> Dict[str, object]:
        if key not in student_hits:
            student_hits[key] = {
                COLUMN_FILE_PATH: document.path,
                COLUMN_ACTIVITY_NAME: document.activity,
                COLUMN_STATUS: "OK",
                COLUMN_MATCH_TYPE: "",
                COLUMN_MATCH_VALUE: "",
                COLUMN_STUDENT_ID: "",
                COLUMN_STUDENT_NAME: "",
                COLUMN_SNIPPET: [],
                COLUMN_MATCH_COUNT: 0,
                "_types": set(),
                "_values": [],
            }
        return student_hits[key]

    for student_id, name in roster.id_to_name.items():
        for index in find_occurrences(text, student_id):
            snippet = get_text_context(text, index, len(student_id))
            bucket = ensure_student_bucket(f"sid:{student_id}")
            bucket["_types"].add("学号")
            if student_id not in bucket["_values"]:
                bucket["_values"].append(student_id)
            bucket[COLUMN_MATCH_COUNT] = int(bucket[COLUMN_MATCH_COUNT]) + 1
            bucket[COLUMN_STUDENT_ID] = student_id
            bucket[COLUMN_STUDENT_NAME] = name
            snippets: List[str] = bucket[COLUMN_SNIPPET]
            if snippet not in snippets:
                snippets.append(snippet)

    for name in set(roster.searchable_names):
        if not (2 <= len(name) <= 10):
            continue
        for index in find_occurrences(text, name):
            student_id = roster.find_student_id(name) or ""
            display_name = (
                name
                if "·" not in name
                else next((candidate for candidate in roster.student_names() if candidate.replace("·", "") == name), name)
            )
            key = f"sid:{student_id}" if student_id else f"name:{name}"
            bucket = ensure_student_bucket(key)
            bucket["_types"].add("姓名")
            if name not in bucket["_values"]:
                bucket["_values"].append(name)
            bucket[COLUMN_MATCH_COUNT] = int(bucket[COLUMN_MATCH_COUNT]) + 1
            if student_id and not bucket[COLUMN_STUDENT_ID]:
                bucket[COLUMN_STUDENT_ID] = student_id
            if not bucket[COLUMN_STUDENT_NAME]:
                bucket[COLUMN_STUDENT_NAME] = display_name
            snippet = get_text_context(text, index, len(name))
            snippets = bucket[COLUMN_SNIPPET]
            if snippet not in snippets:
                snippets.append(snippet)

    for record in student_hits.values():
        record[COLUMN_MATCH_TYPE] = "+".join(sorted(record["_types"]))
        record[COLUMN_MATCH_VALUE] = "、".join(record["_values"])
        record[COLUMN_SNIPPET] = "\n---\n".join(record[COLUMN_SNIPPET])
        record.pop("_types")
        record.pop("_values")
        collected_rows.append(record)

    for raw_keyword in class_tags:
        keyword = raw_keyword.strip()
        if not keyword:
            continue
        for index in find_occurrences(text, keyword):
            collected_rows.append(
                {
                    COLUMN_FILE_PATH: document.path,
                    COLUMN_ACTIVITY_NAME: document.activity,
                    COLUMN_STATUS: "OK",
                    COLUMN_MATCH_TYPE: CLASS_TAG_LABEL,
                    COLUMN_MATCH_VALUE: keyword,
                    COLUMN_STUDENT_ID: "",
                    COLUMN_STUDENT_NAME: "",
                    COLUMN_SNIPPET: get_text_context(text, index, len(keyword)),
                    COLUMN_MATCH_COUNT: 1,
                }
            )

    # If no hits were found, return an empty list so callers can
    # treat this file as "unmatched" and exclude it from outputs.
    return collected_rows
