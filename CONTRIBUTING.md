# Contributing to automated-accessibility-enhancer

Thanks for your interest in contributing! 🎉  
This project focuses on **automated accessibility remediation for PowerPoint files**, with an emphasis on **reading order, heading structure, and screen-reader usability**, while preserving original slide design.

Contributions of all kinds—code, documentation, testing, and design feedback—are welcome.

---

## Code of Conduct

Please be respectful, constructive, and inclusive in all interactions.  
If you experience or observe unacceptable behavior, contact the maintainer (see **Contact**).

---

## Ways to Contribute

- Bug reports with minimal reproduction decks  
- New accessibility heuristics (reading order, heading detection, structure inference)  
- Test decks and regression tests  
- Documentation improvements  
- Integrations (CLI tools, automation workflows)

---

## Development Setup

```bash
git clone https://github.com/<your-username>/automated-accessibility-enhancer.git
cd automated-accessibility-enhancer
python -m venv .venv
source .venv/bin/activate          # Windows: .\.venv\Scripts\Activate.ps1
pip install -r requirements-dev.txt
```

Then put decks in `pptx_input/` and run:

```bash
python z_reorder.py
```

See the [README](README.md) for the web app, and [docs/](docs/README.md) for everything else.

---

## Testing Expectations

```bash
pip install -r requirements-dev.txt
python -m pytest tests/ -q
```

About 200 tests cover title detection, every report check, whole-folder runs, every page, route and error of the web app, CSRF, restart recovery, and the PDF and OCR pipeline. Bedrock is replaced with a stand-in, and S3 and DynamoDB with moto, so they run in seconds without AWS credentials. The sample decks in `Error Test PPTX/` are rebuilt by [generate_test_pptx.py](generate_test_pptx.py) at the start of each run.

[tests/test_axe.py](tests/test_axe.py) loads every page and state from [docs/accessibility-audit.md](docs/accessibility-audit.md) in headless Chromium and fails on any axe-core violation. It needs Chromium once (`python -m playwright install chromium`) and downloads axe-core from cdnjs; without Playwright it is skipped locally, never in CI.

GitHub Actions runs them on every pull request and push to `main` ([.github/workflows/ci.yml](.github/workflows/ci.yml)).

Every PR must explain how it was tested.

- Include a short **How I tested this** section in the PR
- Use only synthetic PowerPoint decks
- Tests run automatically on every PR and push to `main` (`.github/workflows/ci.yml`)

---

## Pull Request Process

1. Create a feature branch from `main`
2. Make small, focused commits
3. Open a PR describing:
   - What changed and why
   - How you tested it
4. Address review feedback constructively

---

## Contact

Maintainer: **Tanya Nair**  
Email: tanyan07@uw.edu  
LinkedIn: https://www.linkedin.com/in/tanya-nair-617473287/
