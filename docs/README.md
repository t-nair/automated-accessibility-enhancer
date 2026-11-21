# Documentation

## Using the tool

- [What it checks and fixes](what-it-checks.md): every check for slides and PDFs, the report, and the web app
- [WCAG 2.1 coverage](wcag-coverage.md): all 50 Level A and AA criteria, and which are fixed, reported or not covered
- [Command line](command-line.md): a folder of decks, or a single PDF
- [PDF export of fixed decks](pdf-export.md): getting a tagged PDF out of a fixed `.pptx`
- [Prezi](prezi.md): export to `.pptx` first

## Running it

- [Deploying](deployment.md): your own server, how production runs on AWS, and open decisions
- [AWS infrastructure](../infra/README.md): Terraform, CI deploys, sign-in, backups and restores, alarms
- [Contributing](../CONTRIBUTING.md): setup, tests and pull requests
- [Cost report](cost-report.docx): monthly AWS cost and credit runway by scenario and planned feature (Word)
- [Cost estimator](cost_estimator.ipynb): the notebook behind the report; change the parameters and rerun

## How it works

- [How it works](how-it-works.md): the architecture diagram, how a file is processed, and the project layout
- [PDF pipeline](../pipeline/README.md): running it, its safety boundary, and how ready it is
- [PDF routing and OCR](ocr-pathway.md): the exact rules for each kind of page
- [OCR model](../pipeline/ocr_models/README.md): which model, its license and checksum

## Decisions and research

- [Accessibility audit of the web app](accessibility-audit.md): axe-core on every page and state
- [Title detection research](../title_research/README.md): how the model that finds slide titles was chosen
- Decisions on hosting, sign-in and backups are recorded in [deployment.md](deployment.md#production) and in the closed issues they link
