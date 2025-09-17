from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Iterable, Optional

DEFAULT_NAME_TO_ID: Dict[str, str] = {"��׼��": "202401150102", "�Ž���": "202401150103", "������": "202401150104", "�ŬB��": "202401150105", "֣����": "202401150106", "�����": "202401150107", "��С��": "202401150109", "�����": "202401150110", "������": "202401150111", "��˼�": "202401150112", "������": "202401150115", "������": "202401150116", "������": "202401150117", "����": "202401150118", "������": "202401150119", "����": "202401150121", "��ϣ��": "202401150122", "������": "202401150123", "���γ�": "202401150124", "Ҷ�Ϻ�": "202401150125", "������": "202401150126", "������": "202401150127", "���ݷ�": "202401150128", "����": "202401150129", "����": "202401150131", "�����": "202401150133", "��÷��": "202401150134", "�շƷ�": "202401150137", "�����": "202401150139", "����ͥ": "202401150141", "���������Ρ�������": "202401150142", "�������������߶�": "202401150143"}


def _build_variant_lookup(name_to_id: Dict[str, str]) -> Dict[str, str]:
    variants: Dict[str, str] = {}
    for name, sid in name_to_id.items():
        variants.setdefault(name, sid)
        if "��" in name:
            variants.setdefault(name.replace("��", ""), sid)
    return variants


def _build_searchable_names(name_to_id: Dict[str, str]) -> list[str]:
    names = list(name_to_id.keys())
    extras = []
    for name in names:
        if "��" in name:
            extra = name.replace("��", "")
            if extra not in names:
                extras.append(extra)
    return names + extras


@dataclass(frozen=True)
class StudentRoster:
    """Simple container for roster data and lookup helpers."""

    name_to_id: Dict[str, str]
    id_to_name: Dict[str, str]
    searchable_names: list[str]
    variant_lookup: Dict[str, str]

    @classmethod
    def from_mapping(cls, name_to_id: Dict[str, str]) -> "StudentRoster":
        id_to_name = {sid: name for name, sid in name_to_id.items()}
        return cls(
            name_to_id=dict(name_to_id),
            id_to_name=id_to_name,
            searchable_names=_build_searchable_names(name_to_id),
            variant_lookup=_build_variant_lookup(name_to_id),
        )

    @classmethod
    def default(cls) -> "StudentRoster":
        return cls.from_mapping(DEFAULT_NAME_TO_ID)

    def find_student_id(self, name: str) -> Optional[str]:
        return self.variant_lookup.get(name)

    def student_ids(self) -> Iterable[str]:
        return self.id_to_name.keys()

    def student_names(self) -> Iterable[str]:
        return self.name_to_id.keys()

    def resolve_name(self, student_id: str) -> Optional[str]:
        return self.id_to_name.get(student_id)
