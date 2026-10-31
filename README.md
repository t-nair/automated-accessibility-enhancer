# automated-accessibility-enhancer

## About
`automated-accessibility-enhancer` takes PowerPoint files and fixes common accessibility problems automatically. It moves slide titles to the front of the reading order so screen readers announce them first, writes descriptions for images that are missing them, and produces a report showing what it found on every slide.

You can use it two ways: through a small web page where you upload files, or by running the script directly on a folder of presentations.

This project was developed as part of the **University of Washington Advancing Accessibility for Engineering Education** initiative, under the guidance of **Prof. Anant** and **Dr. Lalitha Subramanian**. The work has been presented at **EDUCAUSE 2025 Online** as part of a broader effort to scale accessibility remediation in instructional materials.

Feedback on the design, architecture, or implementation is very welcome. You can reach me via [email](mailto:tanyan07@uw.edu) or [LinkedIn](https://www.linkedin.com/in/tanya-nair-617473287/).

---

## Workflow
![Workflow overview](https://github.com/user-attachments/assets/46f6381d-969f-4e2f-aa14-19f6b2b116c3)

What happens to a file after it is uploaded:

1. The file is saved and added to a queue
2. If it is an older `.ppt`, LibreOffice converts it to `.pptx` first. A `.pdf` goes to the PDF pipeline instead (see "Running the PDF remediation pipeline")
3. Each slide is checked, and its title is moved to the front of the reading order. A real title placeholder or a shape whose name contains "title" (case-insensitive) is used when available; otherwise, `title_detection.py` identifies the likely title based on size, position and wording (see `title_research/README.md`). A subtitle placeholder goes right after the title, and the date, footer and slide number go last; everything else keeps its order. Running the process again makes no further changes. A shape that sits above another in the same column but is read after it is listed in the report for a person to check.
4. Every picture without usable alt text is described by a vision model, and the description is written into the file. A filename (`photo.png`) or a generic PowerPoint name (`Picture 3`) counts as missing alt text. Descriptions written by a person are never overwritten.
5. The language is checked. A deck with no default language, an invalid language code or more than one language is listed in the report. If a language was chosen on the upload form, it is then set across the whole deck
6. An updated `.pptx` and a report are written to the output folder. Under the score, the report lists every change it made (titles moved, descriptions written, language set) by slide, with the value before and after

See [docs/wcag-coverage.md](docs/wcag-coverage.md) for which WCAG 2.1 criteria are fixed, reported only, or not covered yet.

---

## What you need

| Requirement | Needed for | Notes |
|---|---|---|
| Python 3.12 | Everything | |
| PyTorch | Image descriptions | About 2.5 GB |
| An NVIDIA GPU | Speed | Optional. Without one it still works, just slower |
| LibreOffice | Old `.ppt` files | Optional. `.pptx` files work without it |

The first time an image is described, the model files download automatically (about 4.4 GB). After that they are cached and reused.

---

## Setup

Clone the project, make a virtual environment and install the Python packages:

```
python -m venv .venv
.venv\Scripts\Activate.ps1      # macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
```

On Windows, if `python` is not recognized, reinstall Python with "Add python.exe to PATH" ticked, or use `py` instead.

If you have an NVIDIA GPU and want to use it, install the CUDA build of PyTorch instead of the default one. You can check whether it worked with:

```
python -c "import torch; print(torch.cuda.is_available())"
```

If that prints `True`, descriptions will be generated on the GPU, which is roughly five times faster.

To support older `.ppt` files, install LibreOffice:

```
winget install --id TheDocumentFoundation.LibreOffice
```

On Linux this is usually `sudo apt install libreoffice-impress`. If LibreOffice is not installed, `.pptx` files still work normally and `.ppt` uploads are turned away with a message explaining what to do.

---

## Running the web page

```
python app.py
```

This starts the Waitress server on http://localhost:5000; open that in a browser. To put it on a real server, behind HTTPS, see [docs/deployment.md](docs/deployment.md).

From there you can drag and drop one or more files, watch the progress bar while they are processed, read the report for each file, and download the fixed version. If a file fails, the page explains why and gives steps to fix it.

To use a different port, set `PORT` first. In PowerShell:

```
$env:PORT=5001
python app.py
```

---

## Running the script on a folder

If you would rather not use the web page, put your files in `pptx_input` and run:

```
python z_reorder.py
```

The results are written to `pptx_output`. Your original files stay in `pptx_input`, untouched. To move each original into `pptx_output` once it is done, add `--move-originals`; a file already in `pptx_output` with the same name is never replaced.

## Running the PDF remediation pipeline

PDFs can be uploaded on the web page like any other file, or run from the
command line. Both use the same `pipeline` package. It inspects every page and
routes it as digital, scanned, hybrid, or uncertain. It detects likely columns,
tables, images, and dense graphics. Digital text can be rebuilt semantically,
including a conservative two-column ordering path. Scanned and hybrid pages use
local English OCR and are always marked for human review.

Uncertain pages, signed or encrypted files, and unsupported complex content
stop for review rather than being silently rewritten.

On the web page, the download is a tagged PDF rebuilt from the original, titled
with the uploaded file name (unless the PDF already has a title) and set to the
language chosen on the form. The report lists how each page was handled. PDFs
with images stop until their descriptions are approved, which only the command
line can do for now (see `pipeline/README.md`). Exporting the tagged PDF needs
Chrome, Chromium or Edge on the server; set `PDF_PIPELINE_CHROME` if it is not
found. OCR (`rapidocr-onnxruntime`) needs Python 3.11 or 3.12.

On Windows:

```powershell
./pipeline/setup.ps1 -Python python
./.venv/Scripts/python.exe -m pipeline.cli "input.pdf" --output-dir output/remediation
```

---

## Running the tests

```
python -m pytest tests/ -q
```

They cover the helper functions (title detection, reading order, placeholder alt text), the whole pipeline running against the sample files in `Error Test PPTX`, processing a whole folder in one run, every page and route on the website, and the PDF pipeline (`test_pdf_*.py`).

The tests replace the image description model with a stand-in, so they finish in about a second instead of several minutes. Most of the time you see when running them is Python loading PyTorch, not the tests themselves. The tests use temporary folders, so running them never touches real uploads or the submissions file.

GitHub Actions runs these tests on every pull request (`.github/workflows/ci.yml`). The tests stub out the model, so CI never downloads it, but it still installs PyTorch and Transformers because `z_reorder.py` imports them. Not covered yet: a test that runs the real model, and any deployment step. Add a deploy workflow once there is a hosting target.

The sample decks in `Error Test PPTX` are not checked in. `tests/conftest.py` rebuilds them with `generate_test_pptx.py` at the start of every test run. To build them by hand:

```
python generate_test_pptx.py
```

---

## Project structure

```
app.py                  the website: uploading, status, reports, downloads
z_reorder.py            the pipeline: converting, fixing titles, writing descriptions
title_detection.py      finds the title of a slide that has no real title placeholder
title_model.joblib      the small trained model title_detection.py uses
title_research/         how the title model was chosen: data, experiments, results
generate_test_pptx.py   makes the sample files used by the tests
pipeline/               the PDF pipeline: page routing, OCR, image review, tagged PDF export

templates/
    index.html          upload page
    status.html         list of submissions
    details.html        one submission, its report and download

static/
    style.css
    script.js

tests/
    test_z_reorder.py   pipeline tests
    test_app.py         website tests

Error Test PPTX/        sample files for the tests, generated (not tracked)
uploads/                files as they are uploaded
processed/              finished files and reports
data/submissions.db     the list of submissions (SQLite)
data/secret_key.txt     used to sign the visitor cookie
pipeline.log            what happened, including every failure

pptx_input/             folder the script reads (you create it)
pptx_output/            where the script writes fixed files and reports
```

The database is made automatically the first time the app runs.

---

## Things to know

- **Descriptions are a starting point, not a final answer.** The model sometimes gets things wrong, especially names and dates. In testing it read a coffin labelled "Nagy Imre" as "Nagy János", and once invented a label that was not in the picture at all. The report page is there so descriptions can be checked before a file is handed to students.
- **Small text is reported, not changed.** The report lists text below 9 pt by slide, including sizes inherited from the layout and master. Raising sizes can break a layout, so a person decides. To use a different cutoff, set `MIN_FONT_PT` before starting, for example `$env:MIN_FONT_PT=12`.
- **Visitors are kept apart, but there is no login.** Each browser is given a random id in a signed cookie, and you only see files submitted with that same cookie. Nobody can open or download someone else's file, even with the link. But this is not real security: there are no passwords, clearing your cookies loses your history, and anyone sharing a browser shares a history. Real use needs proper accounts or university sign-in.
- **Progress is only kept in memory.** If the server restarts while a file is being processed, that file stays stuck showing "processing".
- **One file is processed at a time.** This is on purpose, because running several descriptions at once uses more GPU memory than most laptops have.
- **It must run as a single process.** The queue, the progress and the model live in the one process, so do not add Gunicorn workers or more instances. See [docs/deployment.md](docs/deployment.md).
- **"Permission denied" or "in use" errors** usually mean the deck is open in PowerPoint, or a OneDrive/Dropbox sync is holding it. Close it or pause syncing, then retry.

---

## Possible Improvements
- **Approving PDF image descriptions on the web page**, which today needs the command line (for PDFs exported from fixed decks, see [docs/pdf-export.md](docs/pdf-export.md))
- **Real accounts**, or university sign-in, instead of the cookie the site uses now
- **Better descriptions**, either from a larger model or a hosted one
- **Configurable rules** for heading detection and structure inference
- **Integration into automation platforms** (e.g. self-hosted workflow engines)

---

## Status
The pipeline and the website both work end to end, including older `.ppt` files, several files at once, and automatic image descriptions. Work is ongoing on documentation, accessibility auditing, and getting it ready to run somewhere other than a laptop.

Contributions, feedback, and design discussions are encouraged.
