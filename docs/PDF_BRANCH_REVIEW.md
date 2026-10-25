# Planned PDF pipeline branch: integration review

## Purpose

This branch adds the decision-making and remediation core for PDFs to the
existing `automated-accessibility-enhancer` repository. It preserves the current
PowerPoint pipeline and website. The new PDF capability is command-line based;
no new website interface is included.

## What the branch changes

The new `pipeline` package provides a complete, review-oriented path from an
input PDF to three artifacts: semantic HTML, a tagged companion PDF, and a JSON
evidence manifest. Each page is classified as digital, scanned, hybrid, or
uncertain from explicit thresholds. Scanned and hybrid pages receive local OCR.
Routes that cannot be remediated safely stop and explain why. The page analysis
now identifies multi-column text, tables, embedded images, and dense graphics.
A conservative two-column ordering path is available. Digital figures now have
an extraction, evidence, classification, approval, and accessible-equivalent
pathway. Tables, chart values, and image semantics that cannot be supported by
the collected evidence remain explicit review gates.

The branch also adds focused tests, a pinned local setup, the verified English
OCR model and license, exact pathway documentation, and small README and
dependency updates. Generated files, local environments, temporary validation
artifacts, and the separate experimental website are excluded.

## Why the decision logic matters

A PDF can contain visible text without usable embedded text, or both a scan and
an incomplete text layer. Applying one treatment to every page can lose content
or duplicate it. The four-route decision keeps this distinction explicit:

- digital pages use their embedded text;
- scanned pages use OCR;
- hybrid pages use OCR and require comparison with the embedded layer;
- uncertain pages pause for a person to decide.

The same cautious behavior applies to encrypted files, signed files, and images
whose purpose and alternative text cannot yet be determined.

## Readiness assessment

The code is ready for a focused prototype branch and review. The 30-test PDF
suite passes, Python compilation succeeds, a real image-only fixture completes
OCR and reconstruction, and an actual browser-generated PDF passes the expanded
tag audit. That audit rejects empty tag trees and checks roles, content
references, parent-tree and page connections, language, title, and Figure alt
text presence.
The integration does not alter the PowerPoint execution path.

This is not yet a general-purpose or conformance-certified PDF remediator.
Production adoption should wait for broader document testing, cross-platform CI,
resource limits, broader image evaluation, automatic chart-data recovery, richer
layout understanding, and a defined human review workflow. PDF/UA or WCAG
claims require independent validation.

## Recommended review order

1. Review `pipeline/inspect_pdf.py` and `pipeline/models.py` for the page routes
   and thresholds.
2. Review `pipeline/reading_order.py`, `pipeline/ocr.py`, and
   `docs/OCR_PATHWAY.md` for ordering evidence and human review behavior.
3. Review `pipeline/images.py` and its review JSON before approving figure text.
4. Review `pipeline/rebuild_html.py`, `pipeline/chrome_pdf.py`, and
   `pipeline/cli.py` for reconstruction, export, verification, and manifests.
5. Run the focused tests, then process representative institutional PDFs before
   considering any production workflow.

## Commit boundary

`docs/PDF_BRANCH_MANIFEST.md` is the authoritative list of files included in
the prepared ZIP. The ZIP uses repository-relative paths and contains no Git
metadata, branch changes, commits, web build output, or local test artifacts.
