# Automated Accessibility Enhancer

[![Tests](https://github.com/t-nair/automated-accessibility-enhancer/actions/workflows/ci.yml/badge.svg)](https://github.com/t-nair/automated-accessibility-enhancer/actions/workflows/ci.yml)

Upload a PowerPoint or a PDF, get back a more accessible file and a report that says what was wrong with it.

The Accessibility Enhancer fixes what it safely can (slide reading order, missing image descriptions, document language, PDF tags) and lists everything else a person needs to look at, scored by severity and mapped to WCAG 2.1. Everything runs on your own machine or server: no course file or image is sent to an outside service.

It was built as part of the **University of Washington Advancing Accessibility for Engineering Education** initiative, under the guidance of **Prof. Anant** and **Dr. Lalitha Subramanian**, and presented at **EDUCAUSE 2025 Online** as part of a broader effort to scale accessibility remediation in instructional materials.

---

## Contents

- [What it does](#what-it-does)
- [The web app](#the-web-app)
- [Quick start](#quick-start)
- [Command line](#command-line)
- [How a file is processed](#how-a-file-is-processed)
- [Deploying](#deploying)
- [Tests](#tests)
- [Project layout](#project-layout)
- [Things to know](#things-to-know)
- [Roadmap](#roadmap)
- [Contributing and contact](#contributing-and-contact)

---

## What it does

### PowerPoint (`.pptx`, and `.ppt` through LibreOffice)

| Check | WCAG 2.1 | What happens |
|---|---|---|
| Slide titles read first | 1.3.2 | **Fixed.** The title is moved to the front of the reading order so screen readers announce it first. Titles typed into a text box are found by a small trained model ([title_research/](title_research/README.md)). |
| Images with no description | 1.1.1 | **Fixed.** A local vision model writes a description into the file. A filename (`photo.png`) or a generic name (`Picture 3`) counts as missing. Descriptions a person wrote are never overwritten. |
| Presentation language | 3.1.1, 3.1.2 | **Fixed** when you pick a language on upload: set across every slide, master and layout. A missing, invalid or mixed language is always reported. |
| Slides with no title | 2.4.2 | Reported, by slide. |
| Low text contrast | 1.4.3 | Reported, with each ratio. Colours are resolved through the theme, the shape fill and the slide, layout and master backgrounds. |
| Small text | 1.4.4 | Reported, including sizes inherited from the layout and master (below 9 pt by default). |
| Vague link text | 2.4.4 | Reported: "click here", "read more", bare URLs, and linked pictures with no useful alt text. |

"Reported" means the tool does not guess: a new title, colour, font size or link wording is a design decision, so it is listed for a person to make. [docs/wcag-coverage.md](docs/wcag-coverage.md) covers all 50 Level A and AA criteria and what is, and is not, handled.

### PDF

Every page is inspected and routed as digital, scanned, hybrid or uncertain. Digital text is rebuilt as a **tagged PDF** with a title and language, including a conservative two-column reading order. Scanned and hybrid pages are read with local OCR and always marked for a person to check. Uncertain pages, signed or encrypted files and complex content stop for review rather than being silently rewritten. See [pipeline/README.md](pipeline/README.md) and [docs/OCR_PATHWAY.md](docs/OCR_PATHWAY.md).

### The report

Every file gets a report that opens with a score out of 100 for the file **as it was uploaded**, and the number of problems at each severity. Each problem is marked critical, serious, moderate or minor, with its WCAG criterion. An excerpt:

```text
Accessibility score: 77/100 (as uploaded)
Issues by severity: 1 critical, 2 serious, 1 moderate, 1 minor

 Shape: Picture 3
 - Alt Text: A bar chart comparing LED and incandescent efficiency.
 - [critical] The image had no real description, so this one was generated. Check it (WCAG 1.1.1).

 [moderate] Title was not read first, now fixed (WCAG 1.3.2): slides 2
 [serious] Slides with no title (WCAG 2.4.2): 4
 [minor] Text below 9 pt (WCAG 1.4.4):
 - Slide 3, Shape: Footnote - 8 pt: 'Source: DOE 2024'
 [serious] Text with too little contrast (WCAG 1.4.3):
 - Slide 5, Shape: Subtitle - 2.31:1, needs 4.5:1: 'Week 3 overview'
```

A critical problem costs 10 points, serious 5, moderate 2 and minor 1. Problems the tool fixed still count, so the score describes the original.

---

## The web app

The main way to use the tool. Drag in one or more `.pptx`, `.ppt` or `.pdf` files (up to 50 MB each), optionally pick the language, and the app:

- queues them and shows a live progress bar for each one
- lists every submission and its status on **Check Status**
- shows the scored report on each file's page, with a download of the fixed file
- explains any failure in plain words with steps to fix it, and offers **Try again**
- lets you delete a submission and its files

Each browser only sees its own submissions, every form is protected by a CSRF token, and if the server restarts, unfinished files are picked up again from their original uploads.

The app itself has been audited with axe-core on every page and state: **no WCAG 2.1 Level A or AA violations** ([docs/accessibility-audit.md](docs/accessibility-audit.md)).

---

## Quick start

You need **Python 3.11 or 3.12**. Optional extras:

| Optional | For |
|---|---|
| An NVIDIA GPU | Image descriptions about five times faster. Without one it works, just slower |
| LibreOffice | Old `.ppt` files. `.pptx` works without it |
| Chrome, Chromium or Edge | PDF uploads (set `PDF_PIPELINE_CHROME` if it is not found) |

```bash
git clone https://github.com/t-nair/automated-accessibility-enhancer.git
cd automated-accessibility-enhancer
python -m venv .venv
.venv\Scripts\Activate.ps1          # macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open http://localhost:5000. To use another port, set `PORT` first (`$env:PORT=5001` in PowerShell).

The first image description downloads the model (`Qwen/Qwen2-VL-2B-Instruct`, about 4.4 GB), which is then cached. PyTorch itself is about 2.5 GB.

**GPU:** install the CUDA build of PyTorch from pytorch.org before `requirements.txt`, then check:

```bash
python -c "import torch; print(torch.cuda.is_available())"
```

**LibreOffice:** `winget install --id TheDocumentFoundation.LibreOffice` on Windows, `sudo apt install libreoffice-impress` on Linux. Without it, `.ppt` uploads are turned away with a message explaining what to do.

On Windows, if `python` is not recognized, reinstall Python with "Add python.exe to PATH" ticked, or use `py`.

---

## Command line

### A folder of decks

```bash
python z_reorder.py                       # reads pptx_input/, writes pptx_output/
python z_reorder.py my_decks fixed_decks  # or name the folders
```

Each deck gets an `_updated.pptx` and an `_alt_text` report. Originals are left where they are; add `--move-originals` to move each one into the output folder once it is done (an existing file there is never replaced).

### A PDF

```bash
python -m pipeline.cli "input.pdf" --output-dir output/remediation
```

This writes a tagged PDF, an HTML companion and a JSON manifest. PDFs with images stop until each description is approved in a review file, which for now only the command line can do ([pipeline/README.md](pipeline/README.md)). On Windows, `./pipeline/setup.ps1 -Python python` sets up the environment in one step.

---

## How a file is processed

![Workflow overview](https://github.com/user-attachments/assets/46f6381d-969f-4e2f-aa14-19f6b2b116c3)

1. The upload is saved, recorded in SQLite and put on the queue. The original is kept untouched, so retries and restarts can start again from it.
2. A `.ppt` is converted to `.pptx` with LibreOffice. A `.pdf` goes to the PDF pipeline instead.
3. Each slide's title is moved to the front of the reading order. A real title placeholder, or a shape named "title", is used when there is one; otherwise [title_detection.py](title_detection.py) picks the likely title from its size, position and wording. Running it again changes nothing.
4. Every picture without usable alt text is described by the vision model, using the slide's text as context.
5. Language, missing titles, contrast, font size and link text are checked, and the language is set if one was chosen.
6. The fixed file and the scored report are written to `processed/`.

---

## Deploying

`python app.py` serves the app with [Waitress](https://docs.pylonsproject.org/projects/waitress/) on Windows or Linux. It must run as **one process**: the queue, the progress and the model live in it. [docs/deployment.md](docs/deployment.md) covers the server it needs, environment variables, where data is kept, a systemd unit, nginx with HTTPS, and the decisions still open (host, sign-in, upload limits, retention).

---

## Tests

```bash
python -m pytest tests/ -q
```

About 200 tests cover title detection, every report check, whole-folder runs, every page, route and error of the web app, CSRF, restart recovery, and the PDF and OCR pipeline. The vision model is replaced with a stand-in, so they run in seconds and never download it. The sample decks in `Error Test PPTX/` are rebuilt by [generate_test_pptx.py](generate_test_pptx.py) at the start of each run.

GitHub Actions runs them on every pull request and push to `main` ([.github/workflows/ci.yml](.github/workflows/ci.yml)).

---

## Project layout

```
app.py                  the web app: uploads, queue, status, reports, downloads
z_reorder.py            the PowerPoint pipeline: conversion, titles, descriptions, every check, the report
title_detection.py      finds the title of a slide with no title placeholder
title_model.joblib      the small trained model it uses
title_research/         how that model was chosen: data, experiments, results
pipeline/               the PDF pipeline: page routing, OCR, image review, tagged PDF export
templates/, static/     the web pages, stylesheet and script
tests/                  the test suite
generate_test_pptx.py   builds the sample decks the tests use
docs/                   WCAG coverage, deployment, accessibility audit, research notes

uploads/                original uploads (created at run time)
processed/              fixed files and reports
data/                   submissions.db and the cookie key
pipeline.log            what happened to every file, including failures
```

---

## Things to know

- **Descriptions are a starting point.** The model sometimes gets names and dates wrong, and has invented a label that was not in the picture. Every generated description is flagged in the report so it can be checked before a file reaches students.
- **There is no login yet.** Each browser gets a random id in a signed cookie, and nobody can open someone else's files, even with the link. But clearing cookies loses your history, and a shared browser shares it. Real use needs university sign-in.
- **One file at a time.** On purpose: several descriptions at once need more GPU memory than most laptops have.
- **"Permission denied" or "in use" errors** usually mean the deck is open in PowerPoint or a OneDrive or Dropbox sync is holding it. Close it or pause syncing and retry.
- **Small-text cutoff:** set `MIN_FONT_PT` (for example `$env:MIN_FONT_PT=12`) before starting.

---

## Roadmap

- Approving PDF image descriptions on the web page, not only the command line
- University sign-in instead of the cookie ([#77](https://github.com/t-nair/automated-accessibility-enhancer/issues/77))
- A host for a shared deployment ([#102](https://github.com/t-nair/automated-accessibility-enhancer/issues/102)), upload limits ([#74](https://github.com/t-nair/automated-accessibility-enhancer/issues/74)) and file retention ([#78](https://github.com/t-nair/automated-accessibility-enhancer/issues/78))
- Charts, SmartArt, grouped pictures and decorative images ([docs/wcag-coverage.md](docs/wcag-coverage.md))
- HTML pages ([#41](https://github.com/t-nair/automated-accessibility-enhancer/issues/41))
- Better descriptions, from a larger or hosted model

Open work is tracked in [issues](https://github.com/t-nair/automated-accessibility-enhancer/issues).

---

## Contributing and contact

Contributions, bug reports and design discussion are welcome; see [CONTRIBUTING.md](CONTRIBUTING.md). You can reach the maintainer, Tanya Nair, by [email](mailto:tanyan07@uw.edu) or on [LinkedIn](https://www.linkedin.com/in/tanya-nair-617473287/).
