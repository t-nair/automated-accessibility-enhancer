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
from .inspect_pdf import inspect_pdf
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
    structure = root.get("/StructTreeRoot")
    return {
        "reopened": True,
        "pages": len(reader.pages),
        "language": str(root.get("/Lang") or ""),
        "marked": bool(root.get("/MarkInfo") and root["/MarkInfo"].get_object().get("/Marked")),
        "hasStructureTree": bool(structure),
        "title": str((reader.metadata or {}).get("/Title") or ""),
    }


def run(
    source: Path,
    output_dir: Path,
    on_stage: ProgressCallback | None = None,
) -> dict[str, object]:
    output_dir.mkdir(parents=True, exist_ok=True)
    _report(on_stage, "inspect", "running", "Inspecting every page and its existing structure.")
    inspection = inspect_pdf(source)
    _report(on_stage, "inspect", "complete", f"Inspected {len(inspection.pages)} pages.")
    stem = source.stem
    html_path = output_dir / f"{stem}.accessible.html"
    pdf_path = output_dir / f"{stem}.accessible.pdf"
    manifest_path = output_dir / f"{stem}.remediation.json"

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
            "Reading order and inferred semantic roles still require human review.",
            "OCR drafts require human review. Uncertain and digital image-bearing pages remain blocked.",
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
        _report(on_stage, "semantic-html", "running", "Rebuilding extracted content with semantic HTML roles.")
        build_semantic_html(source, inspection, html_path, ocr_pages)
        stages.append({"name": "semantic-html", "status": "complete"})
        _report(on_stage, "semantic-html", "complete", "Semantic HTML reconstruction created.")
        _report(on_stage, "tagged-pdf-export", "running", "Exporting the reconstruction through the tagged-PDF print path.")
        export_tagged_pdf(html_path, pdf_path)
        _report(on_stage, "verify", "running", "Reopening the generated PDF and checking its structure markers.")
        verification = _verify_output(pdf_path)
        if not verification["marked"] or not verification["hasStructureTree"]:
            raise RuntimeError("Exported PDF did not pass the required tagged-structure checks.")
        stages.append({"name": "tagged-pdf-export", "status": "complete", "verification": verification})
        _report(on_stage, "tagged-pdf-export", "complete", "Tagged PDF reconstruction exported.")
        _report(on_stage, "verify", "complete", "Generated PDF reopened with structure verification.")
        manifest["outputs"] = {
            "html": str(html_path.resolve()),
            "pdf": str(pdf_path.resolve()),
            "pdfSha256": _sha256(pdf_path),
        }
    except Exception as error:
        stages.append({"name": "rebuild", "status": "blocked", "reason": str(error)})
        manifest["status"] = "review-required"
        manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
        _report(on_stage, "rebuild", "blocked", str(error))
        raise

    manifest["status"] = "review-required" if ocr_pages else "reconstruction-created"
    manifest["reviewRequired"] = bool(ocr_pages)
    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
    return manifest


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("--output-dir", type=Path, default=Path("output/remediation"))
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
