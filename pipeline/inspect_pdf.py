"""Deterministic, full-document PDF inspection and page routing."""

from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Any

import pdfplumber
from pypdf import PdfReader

from .models import DocumentInspection, PageInspection, PageRoute, StructureInspection
from .reading_order import lines_from_words, order_layout_items


MIN_USABLE_TEXT_CHARACTERS = 80
FULL_PAGE_RASTER_COVERAGE = 0.65


def _route_page(text_characters: int, raster_coverage: float) -> tuple[PageRoute, list[str]]:
    has_text = text_characters >= MIN_USABLE_TEXT_CHARACTERS
    has_page_raster = raster_coverage >= FULL_PAGE_RASTER_COVERAGE

    if has_text and not has_page_raster:
        return PageRoute.DIGITAL, ["usable embedded text", "no dominant page-sized raster"]
    if not has_text and has_page_raster:
        return PageRoute.SCANNED, ["sparse embedded text", "dominant raster content"]
    if has_text and has_page_raster:
        return PageRoute.HYBRID, ["usable embedded text", "dominant raster content"]
    return PageRoute.UNCERTAIN, ["sparse embedded text", "no dominant raster detected"]


def _raster_coverage(page: pdfplumber.page.Page) -> float:
    page_width = max(float(page.width), 0.0)
    page_height = max(float(page.height), 0.0)
    page_area = max(page_width * page_height, 1.0)
    rectangles: list[tuple[float, float, float, float]] = []
    for image in page.images:
        # Clip image bounds to the page and measure their union. Summing each
        # image independently can double-count overlaps and misroute a page.
        x0 = min(max(float(image.get("x0", 0)), 0.0), page_width)
        x1 = min(max(float(image.get("x1", 0)), 0.0), page_width)
        y0 = min(max(float(image.get("top", 0)), 0.0), page_height)
        y1 = min(max(float(image.get("bottom", 0)), 0.0), page_height)
        if x1 > x0 and y1 > y0:
            rectangles.append((x0, y0, x1, y1))

    x_edges = sorted({edge for rectangle in rectangles for edge in (rectangle[0], rectangle[2])})
    image_area = 0.0
    for left, right in zip(x_edges, x_edges[1:]):
        intervals = sorted(
            (top, bottom)
            for x0, top, x1, bottom in rectangles
            if x0 < right and x1 > left
        )
        covered_height = 0.0
        if intervals:
            start, end = intervals[0]
            for top, bottom in intervals[1:]:
                if top <= end:
                    end = max(end, bottom)
                else:
                    covered_height += end - start
                    start, end = top, bottom
            covered_height += end - start
        image_area += (right - left) * covered_height
    return image_area / page_area


def _structure_inspection(reader: PdfReader) -> StructureInspection:
    root = reader.trailer["/Root"]
    mark_info = root.get("/MarkInfo")
    marked = bool(mark_info and mark_info.get_object().get("/Marked"))
    structure_root = root.get("/StructTreeRoot")
    if not structure_root:
        return StructureInspection(marked, False, {}, 0, issues=["No structure tree is present."])

    role_counts: Counter[str] = Counter()
    figures_with_alt = 0
    figures_missing_alt = 0
    structure_elements = 0
    marked_content_references = 0
    object_references = 0
    stack: list[Any] = [structure_root]
    visited: set[tuple[int, int] | int] = set()

    while stack:
        item = stack.pop()
        if item is None:
            continue
        if isinstance(item, int):
            marked_content_references += 1
            continue
        if isinstance(item, (float, str, bytes)):
            continue
        key: tuple[int, int] | int
        if hasattr(item, "idnum"):
            key = (item.idnum, item.generation)
        else:
            key = id(item)
        if key in visited:
            continue
        visited.add(key)
        try:
            obj = item.get_object()
        except Exception:
            continue
        if not isinstance(obj, dict):
            continue
        role = obj.get("/S")
        if role:
            structure_elements += 1
            role_name = str(role).lstrip("/")
            role_counts[role_name] += 1
            if role_name == "Figure":
                if obj.get("/Alt"):
                    figures_with_alt += 1
                else:
                    figures_missing_alt += 1
        object_type = str(obj.get("/Type") or "")
        if object_type == "/MCR" or obj.get("/MCID") is not None:
            marked_content_references += 1
            continue
        if object_type == "/OBJR":
            object_references += 1
            continue
        children = obj.get("/K")
        if isinstance(children, list):
            stack.extend(children)
        elif children is not None:
            stack.append(children)

    structure_object = structure_root.get_object()
    has_parent_tree = bool(structure_object.get("/ParentTree"))
    pages_with_struct_parents = sum(1 for page in reader.pages if page.get("/StructParents") is not None)
    issues: list[str] = []
    if not marked:
        issues.append("The document is not marked as tagged.")
    if not role_counts:
        issues.append("The structure tree contains no semantic roles.")
    if not has_parent_tree:
        issues.append("The structure tree has no parent tree.")
    if marked_content_references + object_references == 0:
        issues.append("No structure element points to marked page content or an object reference.")
    if pages_with_struct_parents != len(reader.pages):
        issues.append("One or more pages are not connected to the structure parent tree.")
    if figures_missing_alt:
        issues.append(f"{figures_missing_alt} Figure tag(s) have no alternative text.")

    return StructureInspection(
        marked=marked,
        has_structure_tree=True,
        role_counts=dict(sorted(role_counts.items())),
        figures_with_alt=figures_with_alt,
        structure_elements=structure_elements,
        marked_content_references=marked_content_references,
        object_references=object_references,
        figures_missing_alt=figures_missing_alt,
        has_parent_tree=has_parent_tree,
        pages_with_struct_parents=pages_with_struct_parents,
        issues=issues,
    )


def _table_count(page: pdfplumber.page.Page) -> int:
    try:
        return len(page.find_tables() or [])
    except Exception:
        return 0


def _complex_content(page: pdfplumber.page.Page, words: list[dict[str, Any]]) -> tuple[int, int, int, list[str]]:
    order = order_layout_items(lines_from_words(words), float(page.width))
    tables = _table_count(page)
    drawings = len(page.lines) + len(page.rects) + len(page.curves)
    signals: list[str] = []
    if order.column_count > 1:
        signals.append("multi-column-layout")
    if tables:
        signals.append("table")
    if page.images:
        signals.append("embedded-image")
    if drawings >= 20:
        signals.append("dense-vector-graphics")
    return order.column_count or 1, tables, drawings, signals


def _has_signatures(reader: PdfReader) -> bool:
    fields = reader.get_fields() or {}
    return any(str(field.get("/FT")) == "/Sig" for field in fields.values())


def inspect_pdf(path: str | Path) -> DocumentInspection:
    source = Path(path).resolve()
    with source.open("rb") as stream:
        if b"%PDF-" not in stream.read(1024):
            raise ValueError("The file does not contain a PDF signature in its first 1024 bytes.")

    reader = PdfReader(str(source), strict=False)
    if reader.is_encrypted:
        raise ValueError("Encrypted PDFs require an explicit password-handling route.")

    pages: list[PageInspection] = []
    with pdfplumber.open(str(source)) as document:
        for page_number, page in enumerate(document.pages, start=1):
            text = (page.extract_text() or "").strip()
            words = page.extract_words() or []
            coverage = _raster_coverage(page)
            route, reasons = _route_page(len(text), coverage)
            columns, tables, drawings, complex_content = _complex_content(page, words)
            pages.append(
                PageInspection(
                    page_number=page_number,
                    width=round(float(page.width), 2),
                    height=round(float(page.height), 2),
                    text_characters=len(text),
                    words=len(words),
                    image_count=len(page.images),
                    raster_coverage=round(coverage, 4),
                    route=route,
                    reasons=reasons,
                    column_count=columns,
                    table_count=tables,
                    drawing_count=drawings,
                    complex_content=complex_content,
                    needs_review=bool(complex_content),
                )
            )

    metadata = reader.metadata or {}
    title = str(metadata.get("/Title") or source.stem)
    root = reader.trailer["/Root"]
    language = root.get("/Lang")
    if language:
        language = str(language).lstrip("/") or None
    warnings: list[str] = []
    structure = _structure_inspection(reader)
    if structure.has_structure_tree and not any(
        role in structure.role_counts for role in ("H1", "H2", "H3", "P", "L", "Table")
    ):
        warnings.append("The PDF is marked as tagged, but its tag tree lacks common semantic roles.")
    if any(page.route != PageRoute.DIGITAL for page in pages):
        warnings.append("One or more pages require OCR, hybrid-region handling, or human review.")
    if any(page.image_count for page in pages):
        warnings.append("Images require decorative/informative classification and alt-text review.")
    if any("multi-column-layout" in page.complex_content for page in pages):
        warnings.append("Multi-column reading order was inferred geometrically and requires review.")
    if any(set(page.complex_content) & {"table", "dense-vector-graphics"} for page in pages):
        warnings.append("Tables or dense graphics require a specialized remediation path and human review.")
    signed = _has_signatures(reader)
    if signed:
        warnings.append("The source contains a signature; the original must remain immutable.")

    return DocumentInspection(
        source=str(source),
        title=title,
        language=str(language) if language else None,
        pages=pages,
        encrypted=False,
        signed=signed,
        structure=structure,
        warnings=warnings,
    )
