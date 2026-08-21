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
2. If it is an older `.ppt`, LibreOffice converts it to `.pptx` first
3. Each slide is checked, and any shape named "Title" is moved to the front of the reading order
4. Every picture without alt text is described by a vision model, and the description is written into the file
5. An updated `.pptx` and a report are written to the output folder

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

Clone the project and install the Python packages:

```
pip install -r requirements.txt
```

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

Then open http://localhost:5000 in a browser.

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

The results are written to `pptx_output`.

---

## Running the tests

```
python -m pytest tests/ -q
```

There are 74 tests. They cover the helper functions, the whole pipeline running against the sample files in `Error Test PPTX`, and every page and route on the website.

The tests replace the image description model with a stand-in, so they finish in about a second instead of several minutes. Most of the time you see when running them is Python loading PyTorch, not the tests themselves. The tests use temporary folders, so running them never touches real uploads or the submissions file.

To make more sample files to test with:

```
python generate_test_pptx.py
```

---

## Project structure

```
app.py                  the website: uploading, status, reports, downloads
z_reorder.py            the pipeline: converting, fixing titles, writing descriptions
generate_test_pptx.py   makes the sample files used by the tests

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

Error Test PPTX/        sample files, both broken and valid
uploads/                files as they are uploaded
processed/              finished files and reports
data/submissions.db     the list of submissions (SQLite)
data/secret_key.txt     used to sign the visitor cookie
pipeline.log            what happened, including every failure
```

The database is made automatically the first time the app runs. If an older
`data/submissions.json` is sitting there, its contents are copied in once and the
JSON file is left alone.

---

## Things to know

- **Descriptions are a starting point, not a final answer.** The model sometimes gets things wrong, especially names and dates. In testing it read a coffin labelled "Nagy Imre" as "Nagy János", and once invented a label that was not in the picture at all. The report page is there so descriptions can be checked before a file is handed to students.
- **Visitors are kept apart, but there is no login.** Each browser is given a random id in a signed cookie, and you only see files submitted with that same cookie. Nobody can open or download someone else's file, even with the link. But this is not real security: there are no passwords, clearing your cookies loses your history, and anyone sharing a browser shares a history. Real use needs proper accounts or university sign-in.
- **Progress is only kept in memory.** If the server restarts while a file is being processed, that file stays stuck showing "processing".
- **One file is processed at a time.** This is on purpose, because running several descriptions at once uses more GPU memory than most laptops have.
- **The development server is not meant for real deployment.** It is fine for testing and demos.

---

## Possible Improvements
- **PDF support**, which is the most common format after PowerPoint
- **Real accounts**, or university sign-in, instead of the cookie the site uses now
- **Better descriptions**, either from a larger model or a hosted one
- **Configurable rules** for heading detection and structure inference
- **Integration into automation platforms** (e.g. self-hosted workflow engines)

---

## Status
The pipeline and the website both work end to end, including older `.ppt` files, several files at once, and automatic image descriptions. Work is ongoing on documentation, accessibility auditing, and getting it ready to run somewhere other than a laptop.

Contributions, feedback, and design discussions are encouraged.
