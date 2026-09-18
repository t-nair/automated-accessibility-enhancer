# PDF branch file manifest

The prepared branch ZIP contains exactly these repository-relative files:

- `.gitignore`
- `README.md`
- `requirements.txt`
- `docs/OCR_PATHWAY.md`
- `docs/PDF_BRANCH_REVIEW.md`
- `docs/PDF_BRANCH_MANIFEST.md`
- `pipeline/__init__.py`
- `pipeline/chrome_pdf.py`
- `pipeline/cli.py`
- `pipeline/inspect_pdf.py`
- `pipeline/models.py`
- `pipeline/ocr.py`
- `pipeline/providers.py`
- `pipeline/rebuild_html.py`
- `pipeline/README.md`
- `pipeline/requirements.txt`
- `pipeline/requirements-lock.txt`
- `pipeline/setup.ps1`
- `pipeline/setup_models.py`
- `pipeline/ocr_models/en_PP-OCRv4_rec_mobile.onnx`
- `pipeline/ocr_models/LICENSE`
- `pipeline/ocr_models/README.md`
- `tests/test_pdf_ocr.py`
- `tests/test_pdf_remediation_pipeline.py`

The ZIP intentionally excludes the experimental PDF website, generated HTML or
PDF output, temporary test files, build directories, dependency folders,
virtual environments, Python caches, and Git metadata. It also excludes every
unchanged PowerPoint file because those files already exist on the target branch.
