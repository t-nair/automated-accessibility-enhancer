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

Then run it against a directory of decks:

```bash
python z_reorder.py INPUT_DIR OUTPUT_DIR
```

See [SETUP.md](SETUP.md) for a detailed walkthrough.

---

## Testing Expectations

Every PR must explain how it was tested.

- Include a short **How I tested this** section in the PR
- Use only synthetic PowerPoint decks

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
