"""Provider contracts for OCR and alt-text drafting.

The local OCR implementation returns line evidence and review state. Alt text
is drafted by whoever calls the pipeline (pdf_runner.py asks Claude on Bedrock),
so the pipeline itself never needs AWS.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal, Protocol


@dataclass(slots=True)
class BoundingBox:
    x0: float
    y0: float
    x1: float
    y1: float


@dataclass(slots=True)
class OcrWord:
    text: str
    bounds: BoundingBox
    engine_confidence: float | None


@dataclass(slots=True)
class OcrLine:
    text: str
    bounds: BoundingBox
    engine_confidence: float


@dataclass(slots=True)
class OcrPage:
    page_number: int
    words: list[OcrWord]
    language: str | None
    needs_review: bool
    lines: list[OcrLine] = field(default_factory=list)
    review_reasons: list[str] = field(default_factory=list)
    reading_order_method: str = "unassigned"
    column_count: int = 1
    layout_warnings: list[str] = field(default_factory=list)


@dataclass(slots=True)
class AltTextDraft:
    classification: Literal["decorative", "informative", "functional", "complex"]
    text: str | None
    evidence: list[str]
    needs_review: bool = True
    long_description: str | None = None


class OcrProvider(Protocol):
    def recognize(self, page_number: int, image_path: str, language_hints: list[str]) -> OcrPage: ...


class AltTextProvider(Protocol):
    def draft(self, image_path: str, caption: str | None, nearby_text: str) -> AltTextDraft: ...


class ProviderUnavailable(RuntimeError):
    """Raised when a route requires a provider that has not been configured."""
