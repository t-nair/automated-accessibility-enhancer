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
    column_count: int = 1
    table_count: int = 0
    drawing_count: int = 0
    complex_content: list[str] = field(default_factory=list)
    needs_review: bool = False


@dataclass(slots=True)
class StructureInspection:
    marked: bool
    has_structure_tree: bool
    role_counts: dict[str, int]
    figures_with_alt: int
    structure_elements: int = 0
    marked_content_references: int = 0
    object_references: int = 0
    figures_missing_alt: int = 0
    has_parent_tree: bool = False
    pages_with_struct_parents: int = 0
    issues: list[str] = field(default_factory=list)


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
