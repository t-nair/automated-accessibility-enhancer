# Non-scoring PDF remediation prototype

This Python module improves text-based PDFs by rebuilding their extracted text as
semantic HTML and exporting that HTML through Chrome's tagged-PDF print path. It
does not calculate an accessibility score or claim conformance.

## Run

Use Python 3.11 or 3.12. On Windows run `./pipeline/setup.ps1` from the project
directory once to install the local environment and pinned English OCR model.
See [the OCR pathway guide](../docs/OCR_PATHWAY.md) for exact processing rules.

```powershell
./.venv/Scripts/python.exe -m pipeline.cli "input.pdf" --output-dir output/remediation
```

The command writes an accessible HTML companion, a tagged PDF reconstruction,
and a JSON remediation manifest. The source PDF is never overwritten.

## Safety boundary

Every page is inspected. Digital text pages can be reconstructed. Scanned and
hybrid pages use local OCR and produce review-required drafts. Uncertain,
encrypted, signed, or digital image-bearing documents stop for additional processing.
Layout ML, AI alt-text drafting, and exact-layout tag repair remain future work.

`providers.py` defines OCR and alt-text interfaces; `ocr.py` implements real
offline OCR. Alternative-text generation is not implemented. Scans are retained
as visual references, not claimed to have reviewed image descriptions.

The generated document may reflow and change pagination. Human review is still
required for reading order, heading choices, lists, tables, and non-text content.
