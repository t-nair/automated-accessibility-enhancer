"""Typed intermediate records shared by every remediation stage."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from enum import StrEnum
from typing import Any


class PageRoute(StrEnum):
    DIGITAL = "digital"
    SCANNED = "scanned"
    HYBRID = "hybrid"
    UNCERTAIN = "uncertain"


@dataclass(slots=True)
class PageInspection:
    page_number: int
    width: float
    height: float
    text_characters: int
    words: int
    image_count: int
    raster_coverage: float
    route: PageRoute
    reasons: list[str] = field(default_factory=list)


@dataclass(slots=True)
class StructureInspection:
    marked: bool
    has_structure_tree: bool
    role_counts: dict[str, int]
    figures_with_alt: int


@dataclass(slots=True)
class DocumentInspection:
    source: str
    title: str
    language: str | None
    pages: list[PageInspection]
    encrypted: bool
    signed: bool
    structure: StructureInspection
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)
