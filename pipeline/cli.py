"""Command-line entry point for the non-scoring remediation prototype."""

from __future__ import annotations

import argparse
import hashlib
import json
from datetime import UTC, datetime
from pathlib import Path
from typing import Callable

from pypdf import PdfReader

from .chrome_pdf import export_tagged_pdf
from .images import analyze_digital_images, write_image_review
from .inspect_pdf import _structure_inspection, inspect_pdf
from .rebuild_html import build_semantic_html
from .ocr import recognize_raster_pages


ProgressCallback = Callable[[str, str, str], None]


def _report(
    callback: ProgressCallback | None,
    name: str,
    status: str,
    message: str,
) -> None:
    if callback:
        callback(name, status, message)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _verify_output(path: Path) -> dict[str, object]:
    reader = PdfReader(str(path), strict=False)
    root = reader.trailer["/Root"]
    structure = _structure_inspection(reader)
    language = str(root.get("/Lang") or "")
    title = str((reader.metadata or {}).get("/Title") or "")
    issues = list(structure.issues)
    if not language:
        issues.append("The document catalog has no language.")
    if not title:
        issues.append("The document metadata has no title.")
    semantic_roles = sum(structure.role_counts.get(role, 0) for role in ("Document", "Sect", "H1", "H2", "H3", "P", "L", "LI", "Table", "Figure"))
    passed = all((
        structure.marked,
        structure.has_structure_tree,
        structure.structure_elements > 0,
        semantic_roles > 0,
        structure.has_parent_tree,
        structure.marked_content_references + structure.object_references > 0,
        structure.pages_with_struct_parents == len(reader.pages),
        bool(language),
        bool(title),
        structure.figures_missing_alt == 0,
    ))
    return {
        "reopened": True,
        "pages": len(reader.pages),
        "language": language,
        "marked": structure.marked,
        "hasStructureTree": structure.has_structure_tree,
        "title": title,
        "roleCounts": structure.role_counts,
        "structureElementCount": structure.structure_elements,
        "markedContentReferences": structure.marked_content_references,
        "objectReferences": structure.object_references,
        "figuresMissingAlt": structure.figures_missing_alt,
        "hasParentTree": structure.has_parent_tree,
        "pagesWithStructParents": structure.pages_with_struct_parents,
        "issues": issues,
        "passed": passed,
    }


def run(
    source: Path,
    output_dir: Path,
    on_stage: ProgressCallback | None = None,
    image_decisions: Path | None = None,
) -> dict[str, object]:
    output_dir.mkdir(parents=True, exist_ok=True)
    _report(on_stage, "inspect", "running", "Inspecting every page and its existing structure.")
    inspection = inspect_pdf(source)
    _report(on_stage, "inspect", "complete", f"Inspected {len(inspection.pages)} pages.")
    stem = source.stem
    html_path = output_dir / f"{stem}.accessible.html"
    pdf_path = output_dir / f"{stem}.accessible.pdf"
    manifest_path = output_dir / f"{stem}.remediation.json"
    image_review_path = output_dir / f"{stem}.image-review.json"
    image_directory = output_dir / f"{stem}.images"

    manifest: dict[str, object] = {
        "schemaVersion": 1,
        "createdAt": datetime.now(UTC).isoformat(),
        "mode": "non-scoring semantic reconstruction",
        "source": {
            "path": str(source.resolve()),
            "sha256": _sha256(source),
        },
        "inspection": inspection.to_dict(),
        "stages": [],
        "outputs": {},
        "limitations": [
            "The reconstruction may change visual pagination and formatting.",
            "Two-column reading order is inferred geometrically and still requires human review.",
            "OCR drafts require human review. Uncertain pages and unsupported tables or graphics remain blocked.",
            "Image descriptions are drafts until explicitly approved; complex images require a long description or accessible data table.",
            "Tag verification checks structure connectivity and common roles; it is not a conformance certification.",
            "This workflow improves structure but does not claim PDF/UA or WCAG conformance.",
        ],
    }

    stages = manifest["stages"]
    assert isinstance(stages, list)
    stages.append({"name": "inspect", "status": "complete", "pages": len(inspection.pages)})
    try:
        if inspection.signed:
            raise RuntimeError("Signed documents require explicit review before reconstruction.")
        manifest["ocr"] = {}
        ocr_pages = recognize_raster_pages(
            source, inspection, manifest["ocr"],
            lambda name, status, message: _report(on_stage, name, status, message),
        )
        if ocr_pages:
            stages.append({"name": "ocr", "status": "complete", "pages": list(ocr_pages), "needsReview": True})
        _report(on_stage, "image-remediation", "running", "Extracting image regions and building evidence-backed review records.")
        image_remediations, image_issues = analyze_digital_images(
            source, inspection, image_directory, image_decisions
        )
        if image_remediations:
            write_image_review(image_review_path, image_remediations, image_issues)
            manifest["imageRemediation"] = {
                "reviewFile": str(image_review_path.resolve()),
                "imageDirectory": str(image_directory.resolve()),
                "images": sum(len(items) for items in image_remediations.values()),
                "issues": image_issues,
                "approved": not image_issues,
            }
            stages.append({
                "name": "image-remediation",
                "status": "complete" if not image_issues else "review-required",
                "images": sum(len(items) for items in image_remediations.values()),
                "issues": image_issues,
            })
            if image_issues:
                raise RuntimeError(
                    "Image remediation requires approval. Review " + str(image_review_path.resolve())
                    + " and rerun with --image-decisions pointing to the edited file."
                )
        else:
            manifest["imageRemediation"] = {"images": 0, "issues": [], "approved": True}
            stages.append({"name": "image-remediation", "status": "not-needed", "images": 0})
        _report(on_stage, "image-remediation", "complete", "Image evidence and approvals checked.")
        _report(on_stage, "semantic-html", "running", "Rebuilding extracted content with semantic HTML roles.")
        build_semantic_html(source, inspection, html_path, ocr_pages, image_remediations)
        stages.append({"name": "semantic-html", "status": "complete"})
        _report(on_stage, "semantic-html", "complete", "Semantic HTML reconstruction created.")
        _report(on_stage, "tagged-pdf-export", "running", "Exporting the reconstruction through the tagged-PDF print path.")
        export_tagged_pdf(html_path, pdf_path)
        _report(on_stage, "verify", "running", "Reopening the generated PDF and checking its structure markers.")
        verification = _verify_output(pdf_path)
        if not verification["passed"]:
            raise RuntimeError("Exported PDF did not pass the required tagged-structure checks.")
        stages.append({"name": "tagged-pdf-export", "status": "complete", "verification": verification})
        _report(on_stage, "tagged-pdf-export", "complete", "Tagged PDF reconstruction exported.")
        _report(on_stage, "verify", "complete", "Generated PDF reopened with structure verification.")
        manifest["outputs"] = {
            "html": str(html_path.resolve()),
            "pdf": str(pdf_path.resolve()),
            "pdfSha256": _sha256(pdf_path),
        }
        if image_review_path.exists():
            manifest["outputs"]["imageReview"] = str(image_review_path.resolve())
    except Exception as error:
        stages.append({"name": "rebuild", "status": "blocked", "reason": str(error)})
        manifest["status"] = "review-required"
        manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
        _report(on_stage, "rebuild", "blocked", str(error))
        raise

    remediated_image_pages = set(image_remediations)
    unresolved_layout_review = any(
        page.needs_review
        and (
            page.page_number not in remediated_image_pages
            or bool(set(page.complex_content) - {"embedded-image"})
        )
        for page in inspection.pages
    )
    review_required = bool(ocr_pages) or unresolved_layout_review
    manifest["status"] = "review-required" if review_required else "reconstruction-created"
    manifest["reviewRequired"] = review_required
    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
    return manifest


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("--output-dir", type=Path, default=Path("output/remediation"))
    parser.add_argument(
        "--image-decisions",
        type=Path,
        help="Reviewed image-remediation JSON produced by an earlier run.",
    )
    parser.add_argument("--progress-jsonl", action="store_true")
    arguments = parser.parse_args()

    def emit_progress(name: str, status: str, message: str) -> None:
        print(
            json.dumps({"event": "stage", "name": name, "status": status, "message": message}),
            flush=True,
        )

    manifest = run(
        arguments.source,
        arguments.output_dir,
        emit_progress if arguments.progress_jsonl else None,
        arguments.image_decisions,
    )
    print(
        json.dumps(
            {"event": "result", "status": manifest["status"], "outputs": manifest["outputs"]},
            indent=None if arguments.progress_jsonl else 2,
        ),
        flush=True,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
