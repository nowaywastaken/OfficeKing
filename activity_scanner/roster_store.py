from __future__ import annotations

"""Roster helpers for mapping student names to identifiers."""

from dataclasses import dataclass
from typing import Dict, Iterable, Optional

from .config import STUDENT_ID_MAP


def build_variant_lookup(student_map: Dict[str, str]) -> Dict[str, str]:
    """Allow fuzzy lookups for names that contain middle dots."""

    variants: Dict[str, str] = {}
    for name, student_id in student_map.items():
        variants.setdefault(name, student_id)
        if "·" in name:
            variants.setdefault(name.replace("·", ""), student_id)
    return variants


def build_search_candidates(student_map: Dict[str, str]) -> list[str]:
    """Return a list of search names including dotless variants."""

    names = list(student_map.keys())
    extras: list[str] = []
    for name in names:
        if "·" in name:
            stripped = name.replace("·", "")
            if stripped not in names:
                extras.append(stripped)
    return names + extras


@dataclass(frozen=True)
class StudentDirectory:
    """Container for student roster lookups."""

    name_to_id: Dict[str, str]
    id_to_name: Dict[str, str]
    searchable_names: list[str]
    variant_lookup: Dict[str, str]

    @classmethod
    def from_mapping(cls, student_map: Dict[str, str]) -> "StudentDirectory":
        identifier_to_name = {student_id: name for name, student_id in student_map.items()}
        return cls(
            name_to_id=dict(student_map),
            id_to_name=identifier_to_name,
            searchable_names=build_search_candidates(student_map),
            variant_lookup=build_variant_lookup(student_map),
        )

    @classmethod
    def default(cls) -> "StudentDirectory":
        return cls.from_mapping(STUDENT_ID_MAP)

    def find_student_id(self, name: str) -> Optional[str]:
        return self.variant_lookup.get(name)

    def student_ids(self) -> Iterable[str]:
        return self.id_to_name.keys()

    def student_names(self) -> Iterable[str]:
        return self.name_to_id.keys()

    def resolve_name(self, student_id: str) -> Optional[str]:
        return self.id_to_name.get(student_id)
