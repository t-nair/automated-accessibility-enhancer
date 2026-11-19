# Non-scoring PDF remediation prototype

This Python module improves text-based PDFs by rebuilding their extracted text as
semantic HTML and exporting that HTML through Chrome's tagged-PDF print path. It
does not calculate an accessibility score or claim conformance.

## Run

Use Python 3.11 or 3.12. On Windows run `./pipeline/setup.ps1` from the project
directory once to install the local environment and pinned English OCR model.
See [the OCR pathway guide](../docs/ocr-pathway.md) for exact processing rules.

```powershell
./.venv/Scripts/python.exe -m pipeline.cli "input.pdf" --output-dir output/remediation
```

The command writes an accessible HTML companion, a tagged PDF reconstruction,
and a JSON remediation manifest. The source PDF is never overwritten.

## Safety boundary

Every page is inspected. Digital text pages can be reconstructed. Scanned and
hybrid pages use local OCR and produce review-required drafts. A conservative
two-column heuristic orders embedded and OCR text while recording the inference
for review. The inspector separately flags columns, tables, embedded images,
and dense vector graphics. Uncertain, encrypted, signed, unsupported complex,
or digital image-bearing documents stop for additional processing. Layout ML,
fully automatic chart interpretation and exact-layout tag repair remain future work.

`images.py` extracts figures from digital pages, collects nearby text and local
OCR evidence, classifies likely purpose, and writes an editable image-review
file. The first run stops until each image is approved. Rerun with the reviewed
file to place decorative images as artifacts, add approved alternatives to
informative or functional images, and add a long description or accessible data
table for complex images. Scans are retained as visual references, not claimed
to have reviewed image descriptions.

```powershell
# First run creates <name>.image-review.json when figures need decisions.
./.venv/Scripts/python.exe -m pipeline.cli "input.pdf" --output-dir output/remediation

# After reviewing and setting approved=true for every image:
./.venv/Scripts/python.exe -m pipeline.cli "input.pdf" --output-dir output/remediation `
  --image-decisions output/remediation/input.image-review.json
```

The approval gate rejects generic placeholders, complex images without a long
description or data table, and numeric claims absent from the extracted caption,
nearby text, or image OCR evidence.

The generated document may reflow and change pagination. Human review is still
required for inferred reading order, heading choices, lists, tables, charts,
and non-text content. Output verification checks semantic roles, marked-content
connections, the parent tree, page structure-parent links, language, title, and
missing Figure alternatives. It does not certify PDF/UA or WCAG conformance.

## How ready it is

The 30-test PDF suite passes, a real image-only fixture completes OCR and
reconstruction, and a browser-generated PDF passes the tag audit, which rejects
empty tag trees and checks roles, content references, parent-tree and page
connections, language, title and Figure alt text.

It is not yet a general-purpose or conformance-certified PDF remediator. Before
relying on it, it needs broader document testing, resource limits, broader image
evaluation, automatic chart-data recovery, richer layout understanding and a
defined human review workflow. PDF/UA or WCAG claims need independent validation.

## Reading the code

1. `inspect_pdf.py` and `models.py`: the page routes and thresholds.
2. `reading_order.py`, `ocr.py` and [the OCR pathway guide](../docs/ocr-pathway.md):
   ordering evidence and human review.
3. `images.py` and its review JSON: figure text and approval.
4. `rebuild_html.py`, `chrome_pdf.py` and `cli.py`: reconstruction, export,
   verification and manifests.
