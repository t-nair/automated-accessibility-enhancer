# Setup Guide

Step-by-step instructions to get `automated-accessibility-enhancer` running on your own machine.

Follow the sections in order. Steps 1–5 are the setup; Step 6 is your first real run.

---

## Table of Contents

1. [What you need before you start](#step-1-what-you-need-before-you-start)
2. [Install Python](#step-2-install-python)
3. [Get the code](#step-3-get-the-code)
4. [Create a virtual environment](#step-4-create-a-virtual-environment)
5. [Install dependencies](#step-5-install-dependencies)
6. [Run the script](#step-6-run-the-script)
7. [What you get back](#step-7-what-you-get-back)
8. [Command reference](#command-reference)
9. [Troubleshooting](#troubleshooting)
10. [Known limitations](#known-limitations)
11. [Uninstall / clean up](#uninstall--clean-up)

---

## Step 1: What you need before you start

| Requirement | Details |
|---|---|
| **Operating system** | Windows 10/11, macOS, or Linux. All three work — commands are given for each. |
| **Python** | 3.9 or newer. Verified working on 3.13.7. |
| **Git** | Optional — you can download the code as a ZIP instead. |
| **Disk space** | ~50 MB. Descriptions come from Claude on Amazon Bedrock, so there is no model to download. |
| **AWS account** | Only needed for image descriptions, with Bedrock model access enabled. Without it, reading order is still fixed and pictures get placeholder alt text. |
| **Test files** | A few `.pptx` files. The script does not modify them in place, but working on copies is still good practice. |

---

## Step 2: Install Python

Skip this if you already have Python 3.9+.

**Check what you have first.** Open a terminal (Windows: PowerShell; macOS/Linux: Terminal) and run:

```bash
python --version
```

If that prints `Python 3.9.x` or higher, go to Step 3. If it errors, or prints Python 2.x, try `python3 --version` before installing anything.

**Windows**

1. Go to <https://www.python.org/downloads/windows/> and download the latest 64-bit installer.
2. Run the installer.
3. **Tick "Add python.exe to PATH"** on the first screen. This is the single most common setup mistake — if you miss it, `python` will not be found in your terminal.
4. Click *Install Now*.
5. Close and reopen PowerShell, then re-run `python --version` to confirm.

**macOS**

```bash
brew install python
```

If you don't have Homebrew, install it from <https://brew.sh>, or download the macOS installer from <https://www.python.org/downloads/macos/>.

**Linux (Debian/Ubuntu)**

```bash
sudo apt update && sudo apt install python3 python3-pip python3-venv
```

> On macOS and most Linux distributions the command is `python3`, not `python`. Wherever this guide says `python`, use `python3` instead.

---

## Step 3: Get the code

**Option A — with Git (recommended, makes updates easy)**

```bash
git clone https://github.com/t-nair/automated-accessibility-enhancer.git
```

```bash
cd automated-accessibility-enhancer
```

**Option B — without Git**

1. Open <https://github.com/t-nair/automated-accessibility-enhancer>.
2. Click the green **Code** button → **Download ZIP**.
3. Extract the ZIP somewhere you can find again (e.g. `Documents\accessibility-enhancer`).
4. Open a terminal and `cd` into the extracted folder.

**Confirm you're in the right place.** This should list `z_reorder.py`, `requirements.txt`, `README.md`, and `CONTRIBUTING.md`:

```bash
ls
```

On Windows PowerShell, `ls` works too. If you see those files, you're in the right directory.

---

## Step 4: Create a virtual environment

A virtual environment keeps this project's packages separate from the rest of your system, so nothing you install here can break another Python project.

**Create it** (run once, from inside the project folder):

```bash
python -m venv .venv
```

**Activate it** (run every time you open a new terminal to work on this project):

*Windows — PowerShell:*

```powershell
.\.venv\Scripts\Activate.ps1
```

*Windows — Command Prompt (cmd.exe):*

```
.\.venv\Scripts\activate.bat
```

*macOS / Linux:*

```bash
source .venv/bin/activate
```

You'll know it worked when your prompt is prefixed with `(.venv)`.

> **PowerShell blocks the activation script?** If you see *"running scripts is disabled on this system"*, run this once, then activate again:
>
> ```powershell
> Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
> ```

---

## Step 5: Install dependencies

Make sure `(.venv)` is showing in your prompt before running these.

There are three lists, so you install only what you need:

| File | What it gets you |
|---|---|
| `requirements.txt` | The pipeline on its own: `python-pptx`, `Pillow`, and the Anthropic SDK. |
| `requirements-web.txt` | The above, plus Flask for the website. |
| `requirements-dev.txt` | The above, plus pytest. |

To work on the code, install the last one:

```bash
pip install -r requirements-dev.txt
```

All three together are about 50 MB and take seconds.

**To get image descriptions**, you also need AWS credentials the SDK can find
(`aws configure`, environment variables, or an SSO profile) and Bedrock model
access turned on for the Claude model you want, under **Bedrock → Model access**
in the AWS console. `BEDROCK_MODEL_ID` and `BEDROCK_REGION` override the
defaults; see the README for what they accept.

**Verify the install:**

```bash
python -c "import pptx; print('python-pptx', pptx.__version__)"
```

Expected output: `python-pptx 1.0.2` (or newer).

---

## Step 6: Run the script

The script takes two arguments: the folder to read decks from, and the folder to write results to. There is nothing to edit inside the file.

**6a. Create an input folder and put some decks in it.**

*Windows:*

```powershell
mkdir "$HOME\Documents\pptx_input"
```

*macOS / Linux:*

```bash
mkdir -p ~/Documents/pptx_input
```

Copy a few `.pptx` files into it. You do not need to create the output folder — the script creates it for you.

**6b. Run it.**

*Windows:*

```powershell
python z_reorder.py "$HOME\Documents\pptx_input" "$HOME\Documents\pptx_output"
```

*macOS / Linux:*

```bash
python z_reorder.py ~/Documents/pptx_input ~/Documents/pptx_output
```

**6c. Read the output.** The script reports one line per deck:

```
OK      lecture01.pptx: 1 title(s) reordered, 3 shape(s) need alt text
OK      lecture02.pptx: 2 title(s) reordered, 0 shape(s) need alt text
FAILED  broken.pptx: PackageNotFoundError: Package not found at ...

Processed 2 of 3 file(s). Output written to /home/you/Documents/pptx_output
```

A file that fails is reported and skipped; the rest of the batch still runs. The exit code is `0` if everything succeeded and `1` if any file failed, so the script can be used in a larger pipeline.

**Your original files are left untouched** in the input folder. If you want them moved into the output folder once processed, add `--move-originals`.

---

## Step 7: What you get back

For an input file named `deck.pptx`, the output folder will contain:

| File | What it is |
|---|---|
| `deck_updated.pptx` | The enhanced deck. Title shapes have been moved to the front of each slide's z-order, so screen readers announce them first. All visual design and content is preserved. |
| `deck_alt_text.txt` | A per-slide report of every shape and its alt-text status. |

The report looks like this:

```
Alt-text report for deck.pptx
============================================================

Slide 1
  Title 1
    OK - text is read directly: Introduction to Statics
  Picture 2
    NEEDS ALT TEXT - alt text is a filename, not a description: 'diagram.png'
  Picture 3
    NEEDS ALT TEXT - image with no alt text

============================================================
Shapes needing alt text: 2
```

Three things to know about the report:

- It is a **read-only audit**. The script does not write alt text into the deck — that is the planned image alt-text feature, not current behavior. Use the report as a worklist and add the text yourself in PowerPoint.
- **`NEEDS ALT TEXT` includes images whose alt text is just a filename.** PowerPoint fills this in automatically when you insert a picture, so a deck can look like it has alt text everywhere while being useless to a screen reader.
- **Shapes with text are marked OK** rather than flagged. A screen reader reads their text content directly, so they do not need alt text.

**Sanity-check the reordering.** Open `deck_updated.pptx` in PowerPoint and use **Home → Arrange → Selection Pane**. PowerPoint lists shapes in *reverse* z-order there, so the title should now appear at the **bottom** of that list — meaning it is first in the reading order.

---

## Command reference

```
python z_reorder.py INPUT_DIR OUTPUT_DIR [--move-originals]
```

| Argument | Meaning |
|---|---|
| `INPUT_DIR` | Folder to scan for `.pptx` files. Not searched recursively. |
| `OUTPUT_DIR` | Folder to write results to. Created automatically if it does not exist. |
| `--move-originals` | After processing, move each original file out of the input folder and into the output folder. **Off by default** — without it, your input folder is left exactly as it was. |
| `-h`, `--help` | Show usage and exit. |

Files skipped automatically: anything not ending in `.pptx`, and PowerPoint's `~$` lock files for decks you currently have open.

---

## Troubleshooting

**`ModuleNotFoundError: No module named 'pptx'`**
Either the virtual environment isn't active (no `(.venv)` in your prompt — redo [Step 4](#step-4-create-a-virtual-environment)), or the install didn't happen (redo [Step 5](#step-5-install-dependencies)). Note the package is installed as `python-pptx` but imported as `pptx` — that's normal.

**`error: the following arguments are required: input_dir, output_dir`**
You ran the script with no arguments. It needs both folders — see [Step 6](#step-6-run-the-script).

**`Error: Input directory does not exist: ...`**
Check the path in the message against your file explorer. On Windows, wrap paths containing spaces in quotes.

**`FAILED <file>: PackageNotFoundError: Package not found at ...`**
That file isn't a valid `.pptx` — it may be a renamed `.ppt`, a corrupted download, or a placeholder file. Open it in PowerPoint and re-save it as `.pptx`.

**`PermissionError: [Errno 13] Permission denied`**
The deck is open in PowerPoint. Close it and re-run. This can also happen if the folder is syncing via OneDrive/Dropbox — pause syncing or work outside the synced directory.

**`No .pptx files found in ...`**
The folder has no `.pptx` files at the top level. The script does not search subfolders, and it skips `.ppt` and `.pptm`. Convert older `.ppt` files via PowerPoint's *Save As* first.

**`python` is not recognized (Windows)**
Python isn't on your PATH. Either reinstall with the *"Add python.exe to PATH"* box ticked, or use the launcher `py` instead of `python` in every command.

**A deck reports `0 title(s) reordered`**
That's usually correct: it means the titles were already first in the reading order, so nothing needed moving. Running the script twice on the same deck reports `0` the second time. If you believe a title was missed, see *Known limitations* below.

---

## Known limitations

- **Titles inside groups are not detected.** The script looks at top-level shapes only. Moving a shape out of a group would change the slide's layout, so grouped titles are deliberately left alone.
- **Title detection is placeholder-type first, name second.** A shape is treated as a title if it is a real PowerPoint title placeholder, or if its name contains "title" (case-insensitive). A hand-drawn text box acting as a title but named something else — "Header", say — won't be picked up. Rename it in the Selection Pane and re-run.
- **Alt text is reported, not generated.** See [Step 7](#step-7-what-you-get-back).
- **Only `.pptx` is supported.** Not `.ppt`, `.pptm`, or PDF.

---

## Uninstall / clean up

Deactivate the virtual environment:

```bash
deactivate
```

To remove everything, delete the project folder. The `.venv` directory lives inside it, so nothing is left behind elsewhere on your system.

---

## Getting help

Open an issue on the repository, or contact the maintainer — see **Contact** in [CONTRIBUTING.md](CONTRIBUTING.md).
