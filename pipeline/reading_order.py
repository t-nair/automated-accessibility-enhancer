"""Deterministic layout helpers for OCR and embedded PDF text.

The implementation recognizes a conservative two-column pattern. It records
the inference as review evidence instead of claiming that geometric order is a
verified logical reading order.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Generic, Iterable, TypeVar

from .providers import BoundingBox


T = TypeVar("T")


@dataclass(slots=True)
class LayoutLine:
    text: str
    bounds: BoundingBox


@dataclass(slots=True)
class ReadingOrderResult(Generic[T]):
    items: list[T]
    method: str
    column_count: int
    warnings: list[str]


def _bounds(item: Any) -> BoundingBox:
    return item.bounds


def _top_left(items: Iterable[T]) -> list[T]:
    return sorted(items, key=lambda item: (_bounds(item).y0, _bounds(item).x0))


def _has_overlapping_vertical_ranges(left: list[T], right: list[T]) -> bool:
    left_top = min(_bounds(item).y0 for item in left)
    left_bottom = max(_bounds(item).y1 for item in left)
    right_top = min(_bounds(item).y0 for item in right)
    right_bottom = max(_bounds(item).y1 for item in right)
    return min(left_bottom, right_bottom) > max(left_top, right_top)


def _order_band(items: list[T], page_width: float) -> tuple[list[T], int]:
    if len(items) < 4 or page_width <= 0:
        return _top_left(items), 1

    midpoint = page_width / 2
    left = [item for item in items if (_bounds(item).x0 + _bounds(item).x1) / 2 < midpoint]
    right = [item for item in items if item not in left]
    if len(left) < 2 or len(right) < 2 or not _has_overlapping_vertical_ranges(left, right):
        return _top_left(items), 1
    return _top_left(left) + _top_left(right), 2


def order_layout_items(items: Iterable[T], page_width: float) -> ReadingOrderResult[T]:
    """Order line-like objects, preserving full-width lines as band separators."""

    source = _top_left(list(items))
    if not source:
        return ReadingOrderResult([], "no-text", 0, [])

    ordered: list[T] = []
    band: list[T] = []
    maximum_columns = 1

    def flush() -> None:
        nonlocal maximum_columns
        if not band:
            return
        band_order, columns = _order_band(band, page_width)
        ordered.extend(band_order)
        maximum_columns = max(maximum_columns, columns)
        band.clear()

    for item in source:
        bounds = _bounds(item)
        width = max(0.0, bounds.x1 - bounds.x0)
        crosses_middle = bounds.x0 < page_width * 0.43 and bounds.x1 > page_width * 0.57
        if width >= page_width * 0.60 or crosses_middle:
            flush()
            ordered.append(item)
        else:
            band.append(item)
    flush()

    if maximum_columns > 1:
        return ReadingOrderResult(
            ordered,
            "geometric-two-column-bands",
            maximum_columns,
            [
                "Two-column reading order was inferred geometrically; compare it with the source page."
            ],
        )
    return ReadingOrderResult(ordered, "top-to-bottom-left-to-right", 1, [])


def lines_from_words(
    words: list[dict[str, Any]],
    y_tolerance: float = 3.0,
    column_gap_tolerance: float = 72.0,
) -> list[LayoutLine]:
    """Group pdfplumber word boxes into visual text lines."""

    usable = [word for word in words if str(word.get("text", "")).strip()]
    usable.sort(key=lambda word: (float(word.get("top", 0)), float(word.get("x0", 0))))
    groups: list[list[dict[str, Any]]] = []
    for word in usable:
        top = float(word.get("top", 0))
        if not groups:
            groups.append([word])
            continue
        current_top = sum(float(item.get("top", 0)) for item in groups[-1]) / len(groups[-1])
        if abs(top - current_top) <= y_tolerance:
            groups[-1].append(word)
        else:
            groups.append([word])

    line_groups: list[list[dict[str, Any]]] = []
    for group in groups:
        group.sort(key=lambda word: float(word.get("x0", 0)))
        current: list[dict[str, Any]] = []
        for word in group:
            if current:
                gap = float(word.get("x0", 0)) - float(current[-1].get("x1", 0))
                if gap > column_gap_tolerance:
                    line_groups.append(current)
                    current = []
            current.append(word)
        if current:
            line_groups.append(current)

    lines: list[LayoutLine] = []
    for group in line_groups:
        lines.append(
            LayoutLine(
                " ".join(str(word["text"]).strip() for word in group),
                BoundingBox(
                    min(float(word.get("x0", 0)) for word in group),
                    min(float(word.get("top", 0)) for word in group),
                    max(float(word.get("x1", 0)) for word in group),
                    max(float(word.get("bottom", word.get("top", 0))) for word in group),
                ),
            )
        )
    return lines
