# Automated Accessibility Enhancer: Project Plan

Last updated: 2026-10-06 · Owner: Tanya Nair

This plan is aligned with the repository's issues and milestones. Items with no matching issue are marked **(no issue yet)** so they can be filed or dropped.

## 1. Summary

The Automated Accessibility Enhancer is an open-source Python pipeline that remediates accessibility problems in faculty course files. It is a scoped tool that UW can run internally. It handles PowerPoint (`.pptx`, and `.ppt` via headless LibreOffice) and is currently hosted only temporarily through Cloudflare.

Early results: on six faculty PPTX files that Anthology Ally flagged for missing headings, every Ally score rose. The best case went from 51% to 99%, with flagged violations dropping from 48 to 3.

Remaining weaknesses:

- Alt-text accuracy on names, dates and other text-heavy content.
- A 4.4 GB model footprint.
- No proper deployment or login.

## 2. Vision and context

The goal is an internally hosted tool that detects **and fixes** accessibility problems in faculty files before students encounter them. Most existing tools only flag problems.

Commercial tools such as OCR-based remediation suites convert files to tagged PDF, ePub, Braille and audio, and fix missing alt text, titles and contrast inline in Canvas. Some cap a file's score by severity: severe issues hold it at 0-30%, major at 30-70%, minor at 70-100%.

| Tool | Role | Fixes or flags |
|---|---|---|
| PowerPoint built-in checker | Native baseline in Microsoft 365 | Flags |
| Anthology Ally | Higher-ed standard in Canvas, Blackboard, Moodle; scores files | Flags, some auto-remediation in Canvas |
| Grackle | Google Workspace docs and slides | Flags |
| Adobe Acrobat Accessibility Checker | Commercial checker; guides manual remediation | Flags and guides |
| PAC | Verifies structural tags, useful after PPTX to PDF | Flags |
| PAVE | PDF checker against WCAG 2.2 | Flags only |
| WAVE | Website evaluation | Flags |
| Pope Tech | WAVE-powered Canvas scans against WCAG 2.2 A/AA; cannot score or reorder files inside external PDFs | Flags |

## 3. Current state

The tool accepts batches of PowerPoint files through a Flask web interface and returns a slide-by-slide accessibility report. It is a prototype.

Recently shipped:

- Improved page layout and functionality.
- Drag and drop for several files, with a one-at-a-time queue, progress bar and delete button.
- Alt text from `qwen2-vl-2b-instruct`, with slide text fed in as context.
- `.ppt` support via headless LibreOffice.
- Per-visitor private history via a signed cookie (work in progress).
- Report grouped by slide.
- Test suite started, and GitHub Actions CI runs on every PR.

What it does well: reading order and titles, and sorting text into headings and body. What it does not do yet: alt text trustworthy enough to ship unreviewed.

Main issues:

- Mistakes on names, dates and text-based content. Adding slide text fixed some and introduced others, mainly descriptions unrelated to the image.
- No login. UW affiliates would need NetID.
- Not properly deployed.
- About 4.4 GB of model files.
- Axe DevTools results for the web interface are still to be added.

## 4. Workstreams

### 4.1 Slides (core product)

| Milestone | Issues |
|---|---|
| Slides: core checks | #7 title detection, #23 missing titles, #24 language, #25 contrast, #26 font size, #27 table headers, #28 link text |
| Slides: alt text and media | #29 charts/SmartArt/shapes, #30 grouped pictures, #31 decorative images, #32 master and layouts, #33 flag generated descriptions for review, #34 video/audio captions, #35 animations and transitions, #36 unnamed sections |

Open decision on alt text: a local model needs a VM and a heavier deployment, while an API can run on a lighter host such as Render but costs per use. Usage and budget will be explored before choosing. **(no issue yet)** #33 is the closest existing item: it marks generated descriptions as needing human review, which is the right mitigation whichever option wins.

Reducing the 4.4 GB footprint also has **no issue yet**. It depends on the alt-text decision.

### 4.2 Web interface and deployment

| Milestone | Scope |
|---|---|
| Faculty web intake | #17 production frontend (WCAG 2.1 AA, responsive, no-JS resilience) |
| Documentation and accessibility audit | Architecture writeup, test coverage doc, known limitations, axe + manual screen reader audit. This is where the Axe DevTools results above get produced. |
| Accessibility conformance | Contrast palette, responsive layout, in-place polling, inline validation, focus management, base template. Done when axe reports zero violations in CI, 320px reflow is clean and focus indicators are at least 3:1. |
| Observability and fault tolerance | Structured logging with correlation IDs, retry with backoff, axe + pytest in CI, regression tests on the Error Test PPTX corpus |
| Orchestration and notifications | Email on completion or failure, storage lifecycle and retention, the n8n boundary question |
| Deployment and access control | CSRF, `MAX_CONTENT_LENGTH`, WSGI docs, durable queue, UW NetID/SSO replacing cookie identity, retention policy in the UI |
| Submission dashboard and reporting | History and score improvement over time |

Cloud hosting is in progress and waiting on a credits update. NetID integration lives in the Deployment and access control milestone and in #17 gap 10.

### 4.3 Reporting

| Issue | Work |
|---|---|
| #48 | Severity and score per file |
| #49 | Before and after for each change |
| #50 | WCAG coverage table in the docs |

Severity-capped scoring (see section 2) is a useful model for #48.

### 4.4 Beyond slides

Milestone: Beyond slides: web pages and other formats.

**HTML course pages are the next new format in the repo.** #41 accepts `.html` and produces a read-only report. Checks build on it: #42 alt text, #43 headings, #44 links, #45 `lang`, #46 form labels and table headers, #47 research on checks that need a rendered page.

**PDF** is currently tracked only as #40, a check that PDF exports of fixed decks keep tags and titles. Full PDF remediation (the second-format proposal below) has **no issue yet**.

Other format issues: #37 Prezi (research), #38 Google Slides and Keynote export path, #39 `.ppt` conversion tests.

#### PDF remediation proposal

PDF was proposed as the second format. In 2020-2021 Anthology Ally figures, PDFs, Word files and PowerPoints appear in a 15:9:5 ratio, so PDFs affect course scores the most.

- Tagging and alt text can be automated. Reading order cannot yet be generated reliably, and that is the main risk.
- Remediating in place is favored over converting, since most fixes can be made inside the PDF.
- Scanned PDFs with no usable text are the exception: invisible OCR text can be layered on, but labeling becomes increasingly inferred, so conversion may work better.
- Linking an HTML version alongside a PDF is often more accessible for screen readers.
- The proposal says this adds a stage to an n8n flow. The repo lists the n8n boundary as an open question (Orchestration and notifications), so that part is undecided.

**To reconcile:** decide whether PDF remediation or the HTML pathway comes first. The repo is organized around HTML. If PDF remediation stays the plan, file a PDF issue and place it in the Beyond slides milestone.

### 4.5 Competitive benchmarking **(no issue yet)**

The benchmark compares the pipeline with existing checkers on PPTX remediation only. Nothing in the repo tracks it, so the results should be filed or linked from the docs (#50 is the nearest home).

Data points:

- Alt-text and title detection
- Reading order
- Color contrast (4.5:1)
- Table headers
- Detection versus correction, and fix quality
- Preservation of look and function
- Speed, daily usability, actionable feedback
- Accuracy and WCAG 2.1 alignment

Scoring follows the four WCAG 2.1 principles (perceivable, operable, understandable, robust), plus file-format compatibility, time, automation (fixes versus only flags) and possible cost.

**Test set.** The pipeline mainly tags reading order and titles and sorts text into headings, so it was tested on files Ally flagged for matching issues: wrong starting heading level, tables without headers, untagged documents, no headings. That gave 84 files at varying Ally scores. The only PPTX files among them carried "no headings". Each was run through the pipeline and put back in its Canvas course to see whether Ally still flagged it.

| File | Ally before | Ally after | Violations |
|---|---|---|---|
| UW - Ultimate Guide in Slides - December 2021 | 55% | 68% | 69 to 4 |
| Lesson 6: Content Marketing (updated Sept 2024) | 49% | 61% | 55 to 20 |
| accessible_Lesson 6: Content Marketing | 49% | 61% | 55 to 20 |
| Lesson 4 - A Beginners Guide to SEO - Sep 2024 | 51% | 99% | 48 to 3 |
| Lesson 10: Mini Media Plan Presentation Template | 48% | 68% | 2 to 1 |
| UW - Ultimate Guide in Slides - June 2020 | 28% | 48% | 76 to unknown |

The SEO lesson is the strongest result: only contrast issues remained, which the pipeline does not fix (see #25). A seventh file, "UW - Ultimate Guide in Slides - June 2020 (1)", is excluded because it had no Ally score.

Note: the two Lesson 6 rows have identical numbers and one is an `accessible_` copy of the other, so they may not be independent results.

## 5. Roadmap

1. **Decide the alt-text approach** (local model vs API), after usage and budget exploration. No issue yet.
2. **Deploy properly.** Waiting on cloud credits. Covered by the Deployment and access control milestone.
3. **Add real login** with UW NetID. Deployment and access control milestone.
4. **Next format.** HTML (#41-47) is what the repo plans. PDF remediation is proposed but not yet filed (#40 covers exports only).
5. **Test with the faculty files** and keep growing the suite. See the Observability and fault tolerance milestone, #39 and #12.
6. **Reduce the 4.4 GB of model files.** No issue yet.

## 6. Repository practices

- `main` is protected by a ruleset: no force-pushes, and every merge goes through a pull request with one approval. Contributors can approve each other's PRs. The rule mainly keeps bot-generated code out.
- CI runs on every pull request and reports pass or fail before merge.
- Open issues are grouped by milestone. An outside bot has already appeared, a sign the repo is getting traction.
- Older housekeeping issues (#1-#14: CLI args, `python-pptx` internals, overwrite safety, logging, dry run, packaging) have no milestone and still need triage.

## Appendix: Accessibility reference

Standards: WCAG 2.1 and Section 508. Training: UW Accessible Technology Services, the UW Bothell digital accessibility training page, Deque trainings.

### Core document requirements

| Area | Requirement |
|---|---|
| Headings | Real heading styles, start at level 1, no skipped levels, meaningful and concise |
| Alt text | Accurate and under 150 characters; for long descriptions use alt text as a summary and point elsewhere; mark decorative images as decorative |
| Color and contrast | Text 4.5:1; large text and non-text controls 3:1; never convey meaning by color alone |
| Tables | Real tables from the insert tool; simple, one header row and one header column; no merged or split cells; no blank cells |
| Links | Semantic and meaningful; avoid vague text such as "link to" |
| Fonts and layout | Readable fonts, 12 pt in documents and 24 pt in presentations, left-justified, brief paragraphs |
| Metadata | Document title and language set |

### Format-specific notes

- **PowerPoint:** Reading order follows the order objects were added (check the Selection Pane). Every slide needs a meaningful title. Avoid transitions and keep animations simple. The built-in equation tool is not accessible; insert equations as images with alt text. Tables need alt text added manually.
- **Word:** Use built-in heading styles. Built-in equations are generally accessible. Use the Review tab checker.
- **Excel:** Keep A1 non-blank, name each worksheet uniquely, avoid images, merged cells, frozen panes and hidden rows, and mark blank cells.
- **PDF:** Tags give assistive technology the structure. Order: optimize the source, convert to tagged PDF, add metadata, touch up tags, fix reading and tab order, then check. Scanned files need OCR first. Checkers such as PAC do not replace testing with NVDA (free) or JAWS (paid).

### Mobile and interaction

- Touch targets at least 44 x 44 pt (iOS) or 48 dp (Android).
- Text resizes to 200% without clipping.
- Custom controls (carousels, accordions) expose name, role, value and state.

### Why it matters

One in four adults has a disability, and many are never registered for accommodations. Digital accessibility lawsuits rose sharply after 2018, with about 4,000 in 2022 alone. Accessible design also helps people who do not need it, as an elevator does.
