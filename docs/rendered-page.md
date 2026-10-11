# Checks that need a rendered page

Decision: **do not add a headless browser yet.** Revisit once HTML support ([#41](https://github.com/t-nair/automated-accessibility-enhancer/issues/41)) has shipped its static checks. When it is time, use **axe-core driven by Playwright**, as an optional "report only" step.

This is a research note for [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47). It is based on the tools' documentation and general knowledge, not on a hands-on trial (see "Not verified yet").

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
| **axe-core in a headless browser** | Computed-style contrast (1.4.3), bypass/landmark rules (2.4.1), many keyboard-related rules (focusable scroll regions, `tabindex`, etc.). Returns "violations" and "incomplete" (needs a human), which maps onto our Fixed / Reported only split. | A browser binary plus a JS library. |
| **Playwright/Selenium scripts of our own** | Can Tab through the page and detect traps or an unreachable control (2.1.1, 2.1.2). | We would write and maintain the rules ourselves. |
| **pa11y / Lighthouse** | Wrappers around axe-core or similar engines. | Node toolchain on top, for no extra coverage. |

## Recommendation

1. **Now:** static checks only, as planned in #41 and the related issues. They cover most of the cheap wins and need no new dependency.
2. **Later, if users ask for contrast or keyboard reports:** add Playwright (Python) plus axe-core as an *optional* extra, loaded only when the input is HTML and the user opts in. Not in `requirements.txt` by default.
3. **Report only.** Rendered checks tell us a page fails; fixing computed colours or focus behaviour safely is a design decision, so these results go in the report, not into automatic edits.
4. **Skip our own keyboard crawler** unless axe-core's coverage of 2.1.1 and 2.1.2 turns out too thin.

## Costs to weigh

- **Size.** Playwright downloads a Chromium build (hundreds of MB). The project already depends on torch and transformers, so this is not unusual here, but it matters for CI and for anyone who only wants slides.
- **Deployment.** The browser must be installable on the machine running the Flask app.
- **Speed.** A page load per file, not milliseconds. Fine for a batch job, awkward for an instant upload form.
- **Scope.** Slides are `.pptx`; none of this applies to them. It only pays off after HTML input exists.
- **Limits of automation.** Automated engines are widely reported to catch only a minority of WCAG issues. Expect "incomplete" results (for example, text over images) that still need a person.
- **Licence.** axe-core is open source under MPL-2.0. Check that this suits the project before bundling it.

## Not verified yet

- We have not run axe-core on any page or measured install size and run time here.
- We have not checked how well it handles the page shapes this tool is likely to receive (static exports, LMS pages).
- We have not confirmed which of 2.1.1 and 2.4.1 axe-core can fully decide versus only flag as "incomplete".

Spike to settle these: run axe-core through Playwright on three representative pages, record violations, "incomplete" items, install size and seconds per page, and write the numbers here.

## Revisit when

- HTML input and its static checks have shipped.
- Users report contrast or keyboard problems that static checks cannot see.
- The spike shows axe-core decides most of 1.4.3 and 2.4.1 with few "incomplete" results.
