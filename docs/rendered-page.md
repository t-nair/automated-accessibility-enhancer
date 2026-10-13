# Checks that need a rendered page

Decision: **do not add a headless browser yet.** Revisit once HTML support ([#41](https://github.com/t-nair/automated-accessibility-enhancer/issues/41)) has shipped its static checks. When it is time, use **axe-core driven by Playwright**, as an optional "report only" step.

This is a research note for [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47). The comparison table is sourced from the linked pages (checked 2026-10-10); the "Spike results" section is measured; the cost bullets are partly general knowledge.

## Why static HTML is not enough

| Criterion | Why it needs a rendered page |
|---|---|
| [1.4.3](https://www.w3.org/WAI/WCAG21/Understanding/contrast-minimum.html) Contrast (Minimum) | Colours come from the cascade, inheritance, opacity and backgrounds (images, gradients). A static parser sees only inline styles. |
| [2.1.1](https://www.w3.org/WAI/WCAG21/Understanding/keyboard.html) Keyboard | Depends on scripted event handlers and the real focus order. |
| [2.4.1](https://www.w3.org/WAI/WCAG21/Understanding/bypass-blocks.html) Bypass Blocks | A skip link must exist, be reachable first, and move focus. Landmarks and headings can partly be checked statically. |

The same reason applies to the other rows in `wcag-coverage.md` marked "Needs a rendered page" (1.3.4, 1.4.4, 1.4.10, 1.4.12, 1.4.13, 2.1.2).

## Options

| Option | What it gives | Cost |
|---|---|---|
| **Static only** (BeautifulSoup/lxml, no browser) | Alt text, labels, headings, landmarks, `lang`, table headers, explicit skip-link presence. Fast, no new runtime. | Cannot judge contrast, keyboard or focus. |
| **axe-core in a headless browser** | Runs on the rendered DOM, so computed-style contrast (1.4.3) and bypass/landmark rules work. Reports "violations" and "incomplete" (a person must review), which maps onto our Fixed / Reported only split. Its stated goal is zero false positives. MPL-2.0. | A browser binary plus a JS library. Python wrappers: [axe-playwright-python](https://pypi.org/project/axe-playwright-python/) (0.1.8, July 2026, one maintainer) and [axe-selenium-python](https://pypi.org/project/axe-selenium-python/) (3.0.0, April 2026, bundles axe-core 4.10.2, needs Python 3.12+). Or inject `axe.min.js` ourselves, as the spike did. |
| **Playwright/Selenium scripts of our own** | Can Tab through the page and detect traps or an unreachable control (2.1.1, 2.1.2). | We write and maintain the rules ourselves. |
| **[pa11y](https://github.com/pa11y/pa11y)** | A wrapper that drives headless Chrome (Puppeteer) and runs HTML_CodeSniffer (default) and/or axe-core. Exit codes suit CI. | Node 22.13+ toolchain, LGPL-3.0, and results from two engines can duplicate. Adds no coverage over axe-core itself if the axe runner is used. |
| **[Lighthouse](https://developer.chrome.com/docs/lighthouse/accessibility/scoring)** | Its accessibility category is built on axe-core and gives a weighted score. Lists focus traps, tab order and custom control roles as manual checks. | Node tool; a score is less useful to us than per-element findings. Only a subset of audits affects the score. |

## How much automation can see

- Deque's study of 2,000+ audits, 13,000+ pages and about 300,000 issues found automated testing fully covered [57% of issues by count](https://www.deque.com/blog/automated-testing-study-identifies-57-percent-of-digital-accessibility-issues/). That is a vendor figure, measured on issues found in first-time audits, not on share of WCAG criteria.
- The [axe-core README](https://github.com/dequelabs/axe-core) uses the same 57% and says "incomplete" results need manual review. Other sources quote 30 to 50%; treat the share as "roughly a third to a half".
- The [DWP accessibility manual](https://accessibility-manual.dwp.gov.uk/best-practice/automated-testing-using-axe-core-and-pa11y) estimates axe-core plus pa11y at about 35%, a single-source figure.
- Lighthouse's [own scoring page](https://developer.chrome.com/docs/lighthouse/accessibility/scoring) lists focus traps and logical tab order as manual checks, which matches what the spike found.

## Recommendation

1. **Now:** static checks only, as planned in #41 and the related issues. They cover most of the cheap wins and need no new dependency.
2. **Later, if users ask for contrast or keyboard reports:** add Playwright (Python) plus axe-core as an *optional* extra, loaded only when the input is HTML and the user opts in. Not in `requirements.txt` by default.
3. **Report only.** Rendered checks tell us a page fails; fixing computed colours or focus behaviour safely is a design decision, so these results go in the report, not into automatic edits.
4. **Add a small Tab-walk check of our own** for 2.1.1 and 2.1.2 and for skip links. The spike showed axe-core misses traps and does not require a skip link.

## Costs to weigh

- **Size.** Playwright downloads a Chromium build (about 270 MB measured). The project already depends on torch and transformers, so this is not unusual here, but it matters for CI and for anyone who only wants slides.
- **Deployment.** The browser must be installable on the machine running the Flask app.
- **Speed.** A page load per file, not milliseconds. Fine for a batch job, awkward for an instant upload form.
- **Scope.** Slides are `.pptx`; none of this applies to them. It only pays off after HTML input exists.
- **Limits of automation.** Automated engines are widely reported to catch only a minority of WCAG issues. Expect "incomplete" results (for example, text over images) that still need a person.
- **Licence.** axe-core is open source under MPL-2.0. Check that this suits the project before bundling it.

## Spike results

Run on 2026-10-10 with axe-core 4.14.0 and Playwright 1.61 (Python, headless Chromium), on three small hand-written pages in [`rendered-page-spike/`](rendered-page-spike/). Run it with `python run.py` from that folder after `npm pack axe-core` and extracting it to `package/`. The pages are deliberately broken, so this shows what axe-core can and cannot see, not how common the problems are.

| Page | axe-core found | axe-core missed or could not decide |
|---|---|---|
| **A** static export: low-contrast paragraph, banner and link, image without alt, working skip link | 1.4.3: 5 nodes. 1.1.1: 1 node. | A `<div onclick>` that is not focusable (2.1.1): not flagged. |
| **B** LMS-like: no skip link, text over a gradient and a background image, focus outline removed, no `lang` | Missing `lang`. | Contrast over the gradient and image: 11 nodes in **incomplete**, not decided. No skip-link finding, because headings satisfy axe's bypass rule. Removed focus outline (2.4.7): not flagged. |
| **C** trap: input that swallows Tab, `<span role="button">` with no tabindex, no skip link | Nothing at all. | The **keyboard trap** (2.1.2), the unfocusable fake button (2.1.1), and the missing skip link. A scripted Tab walk of the page did show focus stuck on the input. |

What this means:

- **1.4.3 contrast: worth it.** axe-core decides plain text and background colours well. Text over gradients and images lands in "incomplete" and still needs a person.
- **2.4.1 bypass blocks: weak.** A page passes if it has headings or landmarks, even with no skip link. Whether that is acceptable is a judgement call; axe-core will not tell us a skip link is missing.
- **2.1.1 / 2.1.2 keyboard: not covered.** axe-core does not find traps or non-focusable click targets. A small Tab-walk script of our own (about 10 lines in the spike) found the trap, so that is the part we would write ourselves.

Measured cost:

| Item | Size or time |
|---|---|
| axe-core (`axe.min.js`) | 0.65 MB (3.4 MB unpacked package) |
| Playwright Python package | 106 MB on disk (includes its Node driver) |
| Chromium headless shell | 272 MB on disk (full Chromium: 428 MB) |
| Browser launch | 0.55 s once |
| Per page (load, inject, run axe) | 0.2 to 0.4 s for these tiny local pages |

Per-page times are for small local files with no network. Real pages with images, fonts and scripts will be slower; we did not measure that.

## Not verified yet

- Real pages: all three test pages are hand-made. Run the same script on a few real course pages to check noise and timing.
- Whether `axe-core` run through Playwright works on the deployment machine, which was not tried.

## Revisit when

- HTML input and its static checks have shipped.
- Users report contrast or keyboard problems that static checks cannot see.
- Real course pages show axe-core's contrast results are mostly decided rather than "incomplete".
