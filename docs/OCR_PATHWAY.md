# PDF routing and local OCR pathway

This document defines the command-line PDF pathway that is ready for the
proposed branch. It is deliberately separate from the existing PowerPoint
pipeline and does not require or add a website interface.

## What is included

The pathway inspects every page, makes a deterministic routing decision, runs
offline English OCR where that decision permits it, creates a semantic HTML
reconstruction, exports a tagged companion PDF through Chrome/Chromium/Edge,
and writes a JSON evidence manifest. The source PDF is never overwritten.

RapidOCR 1.4.4 runs the PaddleOCR PP-OCRv4 detector, 0/180-degree crop
classifier, and English recognizer through ONNX Runtime on the CPU. It needs no
API key, subscription, or GPU, and it does not upload documents. The pinned
English recognition model is stored at
`pipeline/ocr_models/en_PP-OCRv4_rec_mobile.onnx`; its license and checksum are
included with the code.

## Decision pathway

For each page, the authoritative Python inspection counts stripped embedded
text and measures the clipped union of image rectangles. It applies two fixed
thresholds:

- usable embedded text: at least 80 characters;
- dominant raster content: at least 65% of page area.

Those signals produce four routes:

| Route | Evidence | Action |
|---|---|---|
| Digital | usable text, no dominant raster | rebuild from embedded text |
| Scanned | sparse text, dominant raster | run full-page OCR; require review |
| Hybrid | usable text, dominant raster | run full-page OCR; compare OCR with the embedded layer; require review |
| Uncertain | sparse text, no dominant raster | stop for human review |

The inspection also records likely multi-column layouts, detected tables,
embedded images, and dense vector graphics. Multi-column text can proceed with
a review flag. Digital images enter the image-remediation review pathway.
Digital tables and dense graphics that are not handled as extracted images stop.
Signed and encrypted documents also stop.

For each digital image, the pipeline renders the exact page region, preserves
its coordinates, gathers nearby text and a likely caption, runs local OCR inside
the image when available, and classifies it as decorative, informative,
functional, or complex. It records the extracted evidence and a conservative
draft in `<name>.image-review.json`. Output remains blocked until a reviewer
approves every decision. Complex images require a concise alternative plus a
long description or accessible data table. The gate rejects unsupported numeric
claims and generic placeholders. Approved images are inserted according to their
geometric relationship with surrounding text.

## OCR and evidence

PDFium renders each scanned or hybrid page as RGB at a requested 300 DPI. Scale
is limited to approximately 16 million pixels. Pages are processed sequentially
and temporary full-resolution images are removed when the OCR stage exits.

Every nonempty OCR line is retained, including low-confidence lines. The
manifest records text, line score, bounds, source route, rendered dimensions,
effective DPI, model hashes, and engine/runtime versions. Line results are
ordered using horizontal bands and a conservative two-column test. Full-width
headings remain between bands; within a two-column band, the left column is
read before the right. The selected method, inferred column count, and warning
are recorded in the manifest. No word coordinates are invented from line-level
evidence, and the inferred order always requires comparison with the source.

Every OCR page is marked for review. Additional reasons are recorded when a
line score is below 0.80, the mean line score is below 0.90, or no text is
recognized. Scores are model outputs, not calibrated probabilities and not
proof of accuracy.

## Reconstruction and output

Digital words are grouped into visual lines, split at large column gutters,
ordered with the same banded heuristic, and converted into semantic headings,
paragraphs, lists, and safe links using conservative rules. OCR lines become
escaped paragraphs; the code does not guess that they are headings, lists, or
tables. The source scan is
included as a visual reference, clearly labeled as an unreviewed reference
rather than alternative text.

The browser print path creates the companion PDF in an isolated temporary
profile. The output is reopened and must contain a nonempty semantic structure
tree, marked-content or object references, a parent tree, a structure-parent
link for every page, document language, title metadata, and no Figure tags that
lack alternative text. The manifest records role counts and all of these checks.
A successful structural check establishes internal connectivity, not correct
reading order or PDF/UA/WCAG conformance.

The output directory contains:

- `<name>.accessible.html`, a reflowable review companion;
- `<name>.accessible.pdf`, the tagged reconstruction;
- `<name>.remediation.json`, the inspection, routing, hashes, stage results,
  OCR evidence, limitations, and output verification.

## Setup and command

Use Python 3.11 or 3.12. On Windows, from the repository root:

```powershell
./pipeline/setup.ps1 -Python python
./.venv/Scripts/python.exe -m pipeline.cli "input.pdf" --output-dir output/remediation
```

The setup script creates `.venv`, installs the locked packages, verifies the
bundled OCR model, and initializes the engine. Chrome, Chromium, or Edge is
required for PDF export. Set `PDF_PIPELINE_CHROME` to its executable when it is
not in a standard location or on `PATH`.

For development validation:

```powershell
./.venv/Scripts/python.exe -m unittest tests.test_pdf_ocr tests.test_pdf_remediation_pipeline -v
```

## Boundaries

This branch produces an OCR-assisted remediation draft. Human review is still
required for OCR text, reading order, inferred headings, lists, tables, figures,
mathematical notation, and non-text content. Current OCR is English only. It
does not provide general deskew, full-page 90/270-degree orientation handling,
handwriting recognition, layout understanding beyond the conservative
two-column heuristic, or automatic recovery of chart series and trends. Image
classification and drafts are triage aids; explicit approval remains mandatory.
Large documents can still exceed local time or memory limits.

The local pathway was validated with 30 automated tests, a real synthetic
image-only PDF through OCR and reconstruction, and an actual browser-generated
PDF through the expanded tag-tree audit. A digital-image fixture was also run
through extraction, review-file generation, approval, HTML insertion, PDF export,
and verification of a single Figure tag with approved alternative text.
