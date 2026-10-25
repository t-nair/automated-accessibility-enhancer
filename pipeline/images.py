"""Evidence-backed image triage and approval workflow for digital PDF pages."""

from __future__ import annotations

import base64
import json
import math
import re
from contextlib import closing
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Literal

import pdfplumber

from .models import DocumentInspection, PageRoute
from .ocr import RapidOcrProvider
from .providers import AltTextProvider, ProviderUnavailable
from .reading_order import lines_from_words


ImagePurpose = Literal["decorative", "informative", "functional", "complex"]
PURPOSES = frozenset({"decorative", "informative", "functional", "complex"})
COMPLEX_TERMS = re.compile(
    r"\b(chart|graph|plot|diagram|flowchart|map|table|matrix|timeline|equation|formula|axis|legend|series)\b",
    re.IGNORECASE,
)
FUNCTIONAL_TERMS = re.compile(
    r"\b(qr|barcode|scan|click|tap|button|link|open|download|submit|sign[ -]?in)\b",
    re.IGNORECASE,
)
CAPTION_RE = re.compile(r"^(figure|fig\.?|image|chart|graph|diagram|map|photo|table)\b", re.IGNORECASE)
NUMBER_RE = re.compile(r"(?<![\w.])-?\d+(?:[.,]\d+)*(?:%|\b)")


@dataclass(slots=True)
class ImageRemediation:
    id: str
    page_number: int
    image_number: int
    bounds: dict[str, float]
    page_coverage: float
    image_path: str
    detected_type: str
    classification: ImagePurpose
    caption: str | None
    nearby_text: str
    extracted_text: list[str]
    extracted_values: list[str]
    alt_text: str | None
    long_description: str | None
    accessible_data_table: list[list[str]] | None
    evidence: list[str]
    reviewer_verified_values: list[str] = field(default_factory=list)
    unsupported_claims: list[str] = field(default_factory=list)
    reviewer_notes: str = ""
    approved: bool = False
    needs_review: bool = True
    image_data: str = field(default="", repr=False)

    def review_dict(self) -> dict[str, Any]:
        record = asdict(self)
        record.pop("image_data", None)
        return record


def _clip_bounds(image: dict[str, Any], width: float, height: float) -> tuple[float, float, float, float]:
    x0 = min(max(float(image.get("x0", 0)), 0.0), width)
    x1 = min(max(float(image.get("x1", 0)), 0.0), width)
    top = min(max(float(image.get("top", 0)), 0.0), height)
    bottom = min(max(float(image.get("bottom", 0)), 0.0), height)
    return x0, top, x1, bottom


def _nearby_lines(words: list[dict[str, Any]], bounds: tuple[float, float, float, float]) -> tuple[str | None, str]:
    x0, top, x1, bottom = bounds
    candidates = []
    for word in words:
        word_x0 = float(word.get("x0", 0))
        word_x1 = float(word.get("x1", 0))
        word_top = float(word.get("top", 0))
        word_bottom = float(word.get("bottom", word_top))
        horizontal_overlap = min(x1, word_x1) > max(x0, word_x0)
        vertically_near = word_bottom >= top - 60 and word_top <= bottom + 72
        outside_image = word_bottom <= top + 3 or word_top >= bottom - 3
        if horizontal_overlap and vertically_near and outside_image:
            candidates.append(word)
    lines = lines_from_words(candidates)
    text_lines = [line.text.strip() for line in lines if line.text.strip()]
    caption = next((line for line in text_lines if CAPTION_RE.match(line)), None)
    if caption is None:
        below = [line for line in lines if line.bounds.y0 >= bottom - 3]
        if below:
            caption = below[0].text.strip()
    return caption, " ".join(text_lines)[:1000]


def _detected_type(evidence_text: str) -> str:
    lowered = evidence_text.lower()
    for kind, terms in (
        ("chart", ("chart", "graph", "axis", "legend", "series")),
        ("table", ("table", "matrix")),
        ("diagram", ("diagram", "flowchart")),
        ("map", ("map",)),
        ("mathematical-content", ("equation", "formula")),
        ("functional-code", ("qr", "barcode")),
    ):
        if any(term in lowered for term in terms):
            return kind
    return "unclassified-image"


def _classify(
    coverage: float,
    caption: str | None,
    nearby_text: str,
    extracted_text: list[str],
) -> tuple[ImagePurpose, str]:
    evidence_text = " ".join(filter(None, [caption or "", nearby_text, *extracted_text]))
    detected_type = _detected_type(evidence_text)
    if FUNCTIONAL_TERMS.search(evidence_text):
        return "functional", detected_type
    if COMPLEX_TERMS.search(evidence_text) or (coverage >= 0.12 and len(extracted_text) >= 4):
        return "complex", detected_type
    if coverage <= 0.005 and not extracted_text and not caption:
        return "decorative", detected_type
    return "informative", detected_type


def _draft(
    page_number: int,
    classification: ImagePurpose,
    detected_type: str,
    caption: str | None,
    nearby_text: str,
    extracted_text: list[str],
) -> tuple[str | None, str | None, list[str]]:
    visible = " ".join(extracted_text).strip()
    evidence: list[str] = []
    if caption:
        evidence.append("Draft uses the nearby caption.")
    if visible:
        evidence.append("Draft uses text recovered from inside the image.")
    if classification == "decorative":
        return None, None, ["Small image with no recovered text or caption; decorative status requires approval."]
    source = caption or visible or nearby_text
    if classification == "functional":
        alt = f"Functional image on page {page_number}."
        if source:
            alt += f" Nearby instruction or visible text: {source[:220]}"
        return alt, None, evidence
    if classification == "complex":
        label = detected_type.replace("-", " ") if detected_type != "unclassified-image" else "image"
        alt = f"Complex {label} on page {page_number}."
        if caption:
            alt += f" {caption[:220]}"
        long_description = "Review required before approval."
        if visible:
            long_description += f" Text recovered from the visual: {visible[:1200]}"
        return alt, long_description, evidence
    if caption:
        return caption[:300], None, evidence
    if visible:
        return f"Image containing visible text: {visible[:260]}", None, evidence
    return f"Informative image on page {page_number}; description required.", None, evidence


def _unsupported_numeric_claims(item: ImageRemediation) -> list[str]:
    table_text = " ".join(str(value) for row in (item.accessible_data_table or []) for value in row)
    approved_text = " ".join(filter(None, [item.alt_text or "", item.long_description or "", table_text]))
    evidence_text = " ".join(filter(None, [item.caption or "", item.nearby_text, *item.extracted_text]))
    evidence_numbers = {match.group(0).replace(",", "") for match in NUMBER_RE.finditer(evidence_text)}
    evidence_numbers.update(str(value).replace(",", "") for value in item.reviewer_verified_values)
    claims = {match.group(0).replace(",", "") for match in NUMBER_RE.finditer(approved_text)}
    # The generated page number is structural context, not a statement about the visual.
    claims.discard(str(item.page_number))
    return sorted(claims - evidence_numbers)


def _validate_approval(item: ImageRemediation) -> list[str]:
    issues: list[str] = []
    if not item.approved:
        issues.append("Human approval is required.")
    if item.classification not in PURPOSES:
        issues.append("Classification must be decorative, informative, functional, or complex.")
    if item.classification == "decorative" and item.alt_text:
        issues.append("Decorative images must use empty alternative text.")
    if item.classification in {"informative", "functional", "complex"} and len((item.alt_text or "").strip()) < 5:
        issues.append("Approved non-decorative images require alternative text.")
    if item.classification == "informative" and (item.alt_text or "").startswith("Informative image on page"):
        issues.append("Replace the generated informative-image placeholder before approval.")
    if item.classification == "functional" and (item.alt_text or "").strip() == f"Functional image on page {item.page_number}.":
        issues.append("Replace the generated functional-image placeholder before approval.")
    if item.classification == "complex":
        has_long_description = len((item.long_description or "").strip()) >= 20
        has_table = bool(item.accessible_data_table)
        if not has_long_description and not has_table:
            issues.append("Complex images require a long description or accessible data table.")
        if (item.long_description or "").startswith("Review required before approval"):
            issues.append("Replace the generated complex-image placeholder before approval.")
    item.unsupported_claims = _unsupported_numeric_claims(item)
    if item.unsupported_claims:
        issues.append("Description contains numeric claims absent from extracted evidence: " + ", ".join(item.unsupported_claims))
    if item.reviewer_verified_values and not item.reviewer_notes.strip():
        issues.append("Reviewer-verified values require notes explaining how they were checked against the image.")
    return issues


def _apply_decisions(items: list[ImageRemediation], decisions_path: Path | None) -> None:
    if decisions_path is None:
        return
    payload = json.loads(decisions_path.read_text(encoding="utf-8"))
    decisions = payload.get("images") if isinstance(payload, dict) else None
    if not isinstance(decisions, list):
        raise ValueError("Image decisions must contain an 'images' list.")
    by_id = {str(decision.get("id")): decision for decision in decisions if isinstance(decision, dict)}
    for item in items:
        decision = by_id.get(item.id)
        if not decision:
            continue
        classification = str(decision.get("classification", item.classification))
        if classification in PURPOSES:
            item.classification = classification  # type: ignore[assignment]
        item.alt_text = decision.get("alt_text", decision.get("altText", item.alt_text))
        item.long_description = decision.get("long_description", decision.get("longDescription", item.long_description))
        item.accessible_data_table = decision.get("accessible_data_table", decision.get("accessibleDataTable", item.accessible_data_table))
        verified_values = decision.get("reviewer_verified_values", decision.get("reviewerVerifiedValues", []))
        item.reviewer_verified_values = [str(value) for value in verified_values] if isinstance(verified_values, list) else []
        item.reviewer_notes = str(decision.get("reviewer_notes", decision.get("reviewerNotes", "")))
        item.approved = decision.get("approved") is True


def analyze_digital_images(
    source: Path,
    inspection: DocumentInspection,
    image_directory: Path,
    decisions_path: Path | None = None,
    alt_provider: AltTextProvider | None = None,
) -> tuple[dict[int, list[ImageRemediation]], list[str]]:
    """Extract digital-page figures, draft evidence, and apply explicit approvals."""

    target_pages = {page.page_number for page in inspection.pages if page.route == PageRoute.DIGITAL and page.image_count}
    if not target_pages:
        return {}, []
    import pypdfium2 as pdfium

    image_directory.mkdir(parents=True, exist_ok=True)
    items: list[ImageRemediation] = []
    ocr_provider = None
    try:
        ocr_provider = RapidOcrProvider()
    except ProviderUnavailable:
        pass

    with pdfplumber.open(str(source)) as plumber, closing(pdfium.PdfDocument(str(source))) as document:
        for page_number in sorted(target_pages):
            page = plumber.pages[page_number - 1]
            width, height = float(page.width), float(page.height)
            words = page.extract_words() or []
            with closing(document[page_number - 1]) as pdf_page:
                bitmap = pdf_page.render(scale=2.0)
                try:
                    rendered = bitmap.to_pil().convert("RGB")
                finally:
                    bitmap.close()
            seen: set[tuple[int, int, int, int]] = set()
            for image_number, image in enumerate(page.images, start=1):
                bounds = _clip_bounds(image, width, height)
                x0, top, x1, bottom = bounds
                rounded = tuple(round(value) for value in bounds)
                if rounded in seen or x1 <= x0 or bottom <= top:
                    continue
                seen.add(rounded)
                crop_box = (
                    max(0, math.floor(x0 * 2)),
                    max(0, math.floor(top * 2)),
                    min(rendered.width, math.ceil(x1 * 2)),
                    min(rendered.height, math.ceil(bottom * 2)),
                )
                crop = rendered.crop(crop_box)
                image_id = f"p{page_number}-i{image_number}"
                image_path = image_directory / f"{image_id}.png"
                crop.save(image_path, "PNG")
                extracted_text: list[str] = []
                if ocr_provider is not None:
                    recognized = ocr_provider.recognize(page_number, str(image_path), [])
                    extracted_text = [line.text for line in recognized.lines]
                caption, nearby_text = _nearby_lines(words, bounds)
                coverage = ((x1 - x0) * (bottom - top)) / max(width * height, 1)
                classification, detected_type = _classify(coverage, caption, nearby_text, extracted_text)
                alt_text, long_description, evidence = _draft(
                    page_number, classification, detected_type, caption, nearby_text, extracted_text
                )
                if alt_provider is not None:
                    provider_draft = alt_provider.draft(str(image_path), caption, nearby_text)
                    classification = provider_draft.classification
                    alt_text = provider_draft.text
                    evidence.extend(provider_draft.evidence)
                items.append(
                    ImageRemediation(
                        id=image_id,
                        page_number=page_number,
                        image_number=image_number,
                        bounds={"x0": x0, "top": top, "x1": x1, "bottom": bottom},
                        page_coverage=round(coverage, 4),
                        image_path=str(image_path.resolve()),
                        detected_type=detected_type,
                        classification=classification,
                        caption=caption,
                        nearby_text=nearby_text,
                        extracted_text=extracted_text,
                        extracted_values=sorted({match.group(0) for match in NUMBER_RE.finditer(" ".join(extracted_text))}),
                        alt_text=alt_text,
                        long_description=long_description,
                        accessible_data_table=None,
                        evidence=evidence,
                        image_data=base64.b64encode(image_path.read_bytes()).decode("ascii"),
                    )
                )
                crop.close()
            rendered.close()

    _apply_decisions(items, decisions_path)
    issues: list[str] = []
    for item in items:
        item_issues = _validate_approval(item)
        item.needs_review = bool(item_issues)
        issues.extend(f"{item.id}: {issue}" for issue in item_issues)
    by_page: dict[int, list[ImageRemediation]] = {}
    for item in items:
        by_page.setdefault(item.page_number, []).append(item)
    return by_page, issues


def write_image_review(path: Path, images: dict[int, list[ImageRemediation]], issues: list[str]) -> Path:
    records = [item.review_dict() for page_items in images.values() for item in page_items]
    payload = {
        "schemaVersion": 1,
        "instructions": (
            "Review each extracted image. Correct classification and descriptions, set approved to true, "
            "then rerun with --image-decisions pointing to this file. Complex images need a long description "
            "or accessible data table. Numeric claims must occur in extracted evidence or be listed in "
            "reviewer_verified_values with reviewer_notes explaining the visual check."
        ),
        "images": records,
        "issues": issues,
        "approved": bool(records) and not issues,
    }
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    return path
