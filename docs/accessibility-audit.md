# Accessibility audit of the web app

An axe-core audit of every page of the web app, in every state it can be in.

**Result: no WCAG 2.1 Level A or AA violations on any page or state.** One Level AAA
contrast shortfall was found on the status badges and the error box. AAA is above the level
this project targets ([docs/wcag-coverage.md](wcag-coverage.md) covers A and AA), so it is
recorded below as a note, not as work to do.

Last run: 2026-10-29, against commit `291c27c`.

---

## How it was run

| | |
|---|---|
| Engine | axe-core 4.10.2, injected into the live page and run with `axe.run()` |
| Browser | Chromium 152 |
| Viewport | 1280 x 900 |
| Server | `app.py` served by Waitress on `127.0.0.1:5173` |

Two passes were made over every state:

1. **The default ruleset** — WCAG 2.0 and 2.1 Level A and AA, plus axe's best-practice
   rules. This is what `axe.run()` runs on its own, and what axe DevTools shows by default.
2. **Every rule axe has** (104 rules), which adds the Level AAA and experimental rules.
   This is where the AAA finding came from.

Each state was reached in the running app rather than from a saved copy of the HTML, so the
stylesheet, the JavaScript and the real markup were all in place. The four submission
statuses were written into the database directly, because a real run of the pipeline cannot
be held in `queued` or `processing` long enough to audit. "Reached by" below says how.

A note on viewport: an early run reported ten `color-contrast` "needs review" results on the
upload page. They were an artifact of a zero-width viewport, in which the browser lays
nothing out and axe cannot resolve a background colour. They do not reproduce at a real
size. Always set a viewport before running axe in a headless or hidden browser.

---

## Results

0 violations means axe reported no violations at all for that state. "Passes" is the number
of rules that ran and passed, and is only there to show the run was not a no-op.

### Upload page (`/`)

| State | Reached by | A/AA violations | Passes | All rules |
|---|---|---|---|---|
| Empty | `GET /` | 0 | 33 | 0 |
| Validation error | `POST /upload` with no file chosen, which flashes "Please choose at least one file before submitting." and redirects back | 0 | 37 | 0 |
| Flash message | `POST /upload` with a stale CSRF token, which re-renders the page with a flash at HTTP 400 | 0 | 37 | 0 |

The "validation error" and "flash message" states are the same markup: every message on this
page goes through the one `.flash-message` block. They are listed separately because the
issue asks for both, and because they are different responses — a 302 back to `/`, and a 400
rendered in place.

### Status page (`/status`)

| State | Reached by | A/AA violations | Passes | All rules |
|---|---|---|---|---|
| No submissions | `GET /status` with an empty table | 0 | 20 | 0 |
| Queued, processing, done and failed | one submission in each status, so all four badges and both delete-button states render together | 0 | 27 | 1 rule, 3 nodes (AAA, see below) |
| With a flash message | `GET /details/<unknown id>`, which flashes "That submission could not be found." and redirects here | 0 | 36 | 1 rule, 3 nodes (AAA) |

All four statuses were audited on one page rather than four, because the table is one row per
submission and a row in any status is the same markup with a different badge class. Together
they also cover the mixed case, which is what faculty actually see.

### Details page (`/details/<id>`)

| State | Reached by | A/AA violations | Passes | All rules |
|---|---|---|---|---|
| Queued | a submission in `queued` — progress box, no delete button | 0 | 26 | 0 |
| Processing | a submission in `processing` — progress box and delete button | 0 | 26 | 1 node (AAA) |
| Done, with report | a submission in `done` with a report file on disk — download link and `<pre>` report | 0 | 22 | 1 node (AAA) |
| Failed | a submission in `error`, with an error message and three recovery steps | 0 | 26 | 7 nodes (AAA) |

### Delete confirmation

Deleting asks `confirm("Delete this submission and its files? This cannot be undone.")`
([static/script.js:60](../static/script.js#L60)). That is the browser's own dialog, not page
content, so **there is nothing for axe to inspect** — no DOM, no styles of the app's, no
markup the app controls. The browser makes it accessible itself: it takes focus, it is
announced as a dialog, and Enter and Escape work. Nothing to report, and nothing to fix.

The state after confirming — `/status` with "Deleted finished-deck.pptx." — is the status
page's flash state above, and is clean at A and AA.

---

## The one finding

### 1.4.6 Contrast (Enhanced) — Level AAA

| | |
|---|---|
| axe rule | `color-contrast-enhanced` |
| WCAG criterion | [1.4.6 Contrast (Enhanced)](https://www.w3.org/WAI/WCAG21/Understanding/contrast-enhanced.html), Level **AAA** |
| axe impact | serious |
| Needs | 7:1 for normal text |
| Pages and states | status page with submissions (3 nodes); details page processing (1), done (1), failed (7) |

| Element | Colours | Ratio | Where |
|---|---|---|---|
| `.status-error` badge | `#a94442` on `#f2dede` | 4.53:1 | status table, details page |
| `.status-done` badge | `#3c763d` on `#dff0d8` | 4.56:1 | status table, details page |
| `.status-processing` badge | `#7a5b00` on `#fff3cd` | 5.70:1 | status table, details page |
| `.error-box` body text, its "Try this:" lead-in and the recovery steps | `#a94442` on `#f2dede` | 4.53:1 | details page, failed |

`.status-queued` (`#444` on `#e8eaed`) clears 7:1 and is not listed.

Every one of these passes 1.4.3 Contrast (Minimum) at AA, which needs 4.5:1. They fail only
the AAA threshold. This project targets Level A and AA, so **no issue was filed**. If the
target ever moves to AAA, the fix is to darken those four foreground colours in
`static/style.css`; the backgrounds can stay as they are.

---

## What this audit does not cover

axe-core finds roughly a third of WCAG problems. It cannot judge whether text makes sense,
whether a reading order is logical, or whether a keyboard path is usable. These pages still
need a manual pass for:

- keyboard order and visible focus through the upload form, the drop zone and the table
- the skip link, which is only visible on focus
- a screen reader over the live progress region and the flash messages
- reflow and resize (1.4.10, 1.4.4) at 320 CSS pixels and 200% text, which axe does not test
- the `<pre>` report block, which a screen reader reads line by line

---

## Re-running this

Each page can be audited with axe DevTools, or without the extension by pasting axe-core
into the browser console:

```js
const src = await (await fetch("https://cdnjs.cloudflare.com/ajax/libs/axe-core/4.10.2/axe.min.js")).text();
document.head.appendChild(Object.assign(document.createElement("script"), { textContent: src }));

const result = await axe.run();                        // A and AA, plus best practice
console.table(result.violations.map(v => ({ rule: v.id, impact: v.impact, nodes: v.nodes.length })));

const all = axe.getRules().map(r => r.ruleId);         // adds AAA and experimental
console.log((await axe.run(document, { runOnly: { type: "rule", values: all } })).violations);
```

The four submission statuses need rows in the database. Stop the app first, or the background
worker will move the row on again, then:

```
sqlite3 data/submissions.db "UPDATE submissions SET status = 'processing' WHERE id = '<id>'"
```
