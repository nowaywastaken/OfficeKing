from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Dict, Iterable, List

from columns import (
    ACTIVITY_NAME,
    CLASS_TAG_LABEL,
    FILE_PATH,
    MATCH_COUNT,
    MATCH_TYPE,
    MATCH_VALUE,
    SNIPPET,
    STATUS,
    STUDENT_ID,
    STUDENT_NAME,
)
from roster import StudentRoster


@dataclass(frozen=True)
class Document:
    path: str
    text: str
    activity: str


def gather_input_files(paths: Iterable[str], supported_exts: Iterable[str]) -> List[str]:
    supported = {ext.lower() for ext in supported_exts}
    result: List[str] = []
    for p in paths:
        if os.path.isfile(p):
            if os.path.splitext(p)[1].lower() in supported:
                result.append(p)
        elif os.path.isdir(p):
            for root, _, files in os.walk(p):
                for file in files:
                    if os.path.splitext(file)[1].lower() in supported:
                        result.append(os.path.join(root, file))
        else:
            print(f"[����] �Ҳ���·����{p}")
    return result


def find_all_indices(text: str, sub: str) -> List[int]:
    if not sub:
        return []
    idxs: List[int] = []
    start = 0
    while True:
        idx = text.find(sub, start)
        if idx == -1:
            break
        idxs.append(idx)
        start = idx + len(sub)
    return idxs


def get_context(text: str, idx: int, length: int, span: int = 20) -> str:
    start = max(0, idx - span)
    end = min(len(text), idx + length + span)
    return text[start:end].replace("\n", " ")


def infer_activity_name(path: str, text: str) -> str:
    first_line = ""
    for line in text.splitlines():
        stripped = line.strip()
        if stripped:
            first_line = stripped
            break
    filename = os.path.splitext(os.path.basename(path))[0]
    keywords = ("�", "����", "־Ը", "����", "֪ͨ", "��", "����", "֤��", "����", "��ѵ", "����", "�嵥", "��ʾ")
    if any(k in first_line for k in keywords):
        title = first_line[:80]
    else:
        title = filename[:80]
    parent = os.path.basename(os.path.dirname(path))
    if parent and parent not in title and len(parent) <= 40:
        title = f"{title}��{parent}��"
    return title


def _empty_row(path: str, activity: str) -> Dict[str, object]:
    return {
        FILE_PATH: path,
        ACTIVITY_NAME: activity,
        STATUS: "δ����",
        MATCH_TYPE: "",
        MATCH_VALUE: "",
        STUDENT_ID: "",
        STUDENT_NAME: "",
        SNIPPET: "",
        MATCH_COUNT: 0,
    }


def scan_document(document: Document, roster: StudentRoster, class_tags: List[str]) -> List[Dict[str, object]]:
    text = document.text
    rows: List[Dict[str, object]] = []
    student_hits: Dict[str, Dict[str, object]] = {}

    def ensure_student_row(key: str) -> Dict[str, object]:
        if key not in student_hits:
            student_hits[key] = {
                FILE_PATH: document.path,
                ACTIVITY_NAME: document.activity,
                STATUS: "OK",
                MATCH_TYPE: "",
                MATCH_VALUE: "",
                STUDENT_ID: "",
                STUDENT_NAME: "",
                SNIPPET: [],
                MATCH_COUNT: 0,
                "_types": set(),
                "_values": [],
            }
        return student_hits[key]

    for sid, name in roster.id_to_name.items():
        for idx in find_all_indices(text, sid):
            context = get_context(text, idx, len(sid))
            row = ensure_student_row(f"sid:{sid}")
            row["_types"].add("ѧ��")
            values = row["_values"]
            if sid not in values:
                values.append(sid)
            row[MATCH_COUNT] = int(row[MATCH_COUNT]) + 1
            row[STUDENT_ID] = sid
            row[STUDENT_NAME] = name
            snippets = row[SNIPPET]
            if context not in snippets:
                snippets.append(context)

    for name in set(roster.searchable_names):
        if not (2 <= len(name) <= 10):
            continue
        for idx in find_all_indices(text, name):
            sid_match = roster.find_student_id(name) or ""
            display_name = name if "��" not in name else next(
                (nm for nm in roster.student_names() if nm.replace("��", "") == name),
                name,
            )
            key = f"sid:{sid_match}" if sid_match else f"name:{name}"
            row = ensure_student_row(key)
            row["_types"].add("����")
            values = row["_values"]
            if name not in values:
                values.append(name)
            row[MATCH_COUNT] = int(row[MATCH_COUNT]) + 1
            if sid_match and not row[STUDENT_ID]:
                row[STUDENT_ID] = sid_match
            if not row[STUDENT_NAME]:
                row[STUDENT_NAME] = display_name
            context = get_context(text, idx, len(name))
            snippets = row[SNIPPET]
            if context not in snippets:
                snippets.append(context)

    for row in student_hits.values():
        row[MATCH_TYPE] = "+".join(sorted(row["_types"]))
        row[MATCH_VALUE] = "��".join(row["_values"])
        row[SNIPPET] = "\n---\n".join(row[SNIPPET])
        row.pop("_types")
        row.pop("_values")
        rows.append(row)

    for keyword in class_tags:
        kw = keyword.strip()
        if not kw:
            continue
        for idx in find_all_indices(text, kw):
            rows.append(
                {
                    FILE_PATH: document.path,
                    ACTIVITY_NAME: document.activity,
                    STATUS: "OK",
                    MATCH_TYPE: CLASS_TAG_LABEL,
                    MATCH_VALUE: kw,
                    STUDENT_ID: "",
                    STUDENT_NAME: "",
                    SNIPPET: get_context(text, idx, len(kw)),
                    MATCH_COUNT: 1,
                }
            )

    if not rows:
        rows.append(_empty_row(document.path, document.activity))
    return rows
