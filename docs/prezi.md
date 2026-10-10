# Prezi presentations

Decision: **document only**. The tool does not read Prezi files. Authors export the Prezi to `.pptx` and run that file through the tool.

This is a research note for [#37](https://github.com/t-nair/automated-accessibility-enhancer/issues/37). It is based on Prezi's help pages and third-party accessibility guides, not on a hands-on export test (see "Not verified yet").

## What Prezi can export

| Option | Result | Usable here? |
|---|---|---|
| [Download PPTX](https://support.prezi.com/hc/en-us/articles/36014448143767-How-do-I-download-my-Prezi-presentation-as-PowerPoint-slides) | One slide per frame, opens in PowerPoint. Zoom and pan transitions are not kept. | Yes. This is the path. |
| Download PDF | One static page per frame. Dynamic content is lost. | Not yet. PDF remediation has no issue (see `PLAN.md`). |
| [Portable presentation](https://support.prezi.com/hc/en-us/articles/360003478454-Presenting-and-viewing-a-downloaded-presentation-portable-presentation) (EXE or ZIP) | Offline player, not editable. Needs a Plus plan or higher. | No. Nothing to remediate. |

## Accessibility of Prezi itself

- Prezi documents a "Visible to screen readers" setting that adds a label and description to objects. The help article we found covers Infographics, so it may not apply to every Prezi product.
- Third-party university guides say Prezi has weak screen reader and keyboard support and recommend a different tool. These guides are undated, and Prezi's own marketing says the opposite, so the sources conflict.
- Zooming and rotating between frames can disorient some learners. The tool cannot fix this, because it is a design choice in the Prezi, not a property of the file.

## Recommended workflow

1. In Prezi, use Share, then **Download PPTX**.
2. Upload the `.pptx` to the tool like any other deck.
3. Review the report. Fix what it lists as "reported only".
4. Give learners the remediated `.pptx` (or its PDF export) instead of, or next to, the Prezi link.

## Screen reader support in the portable player

**Untested.** We found no published testing of the portable EXE/ZIP player with NVDA or JAWS, and we have not tested it ourselves. Assume it is not accessible until someone checks.

What the sources say, none of it specific to the offline player:

- Prezi's [accessible design guidance](https://support.prezi.com/hc/en-us/articles/29413944153879-How-to-design-more-accessible-Prezi-presentations) recommends a PDF version for screen reader users.
- A [CUNY accessibility guide](https://guides.cuny.edu/accessibility/prezi) and a [2015 review](https://barrydahl.com/2015/01/08/accessibility-concerns-of-using-prezi-in-education/) say Prezi is not accessibility compliant, cannot be tabbed through, and only moves between frames with the arrow keys. These may predate current features.
- Prezi's [offline help](https://support.prezi.com/hc/en-us/articles/360003498953-Presenting-and-viewing-a-presentation-offline) covers exporting and presenting, and says nothing about screen readers.

Open questions: whether labels set with "Visible to screen readers" carry into the offline player, and whether the player exposes text and focus to a screen reader at all.

To test: export a portable Prezi (Plus plan or higher), open it with NVDA (free, Windows), and try reading text, moving between frames by keyboard and reaching labelled objects. Record the result here.

## Not verified yet

The PPTX export has not been tested with this tool. Open questions:

- Does each exported slide keep a title, or do frames come out as untitled slides? (The tool would flag these.)
- Is any alt text set in Prezi carried over to the pictures in the `.pptx`?
- Is the reading order sensible?

Export one real Prezi, run it through the tool, and record the answers here.

## Test files: more needed

The only Prezi test file so far is `Error Test PPTX/sim_prezi_export.pptx`. It is **simulated**, built by `generate_test_pptx.py`, and is a guess at the export shape (three slides, each one full-slide picture with no alt text and no title). It tests that the tool copes with that shape, not that Prezi's export looks like that.

We still need **real** exports, ideally several, covering:

- a Prezi with text frames and pictures, with and without alt text set in Prezi
- a Prezi with a path that zooms in and out of nested frames
- one with embedded video or a chart

Public Prezis cannot be downloaded without an account, and other people's slides should not be committed without permission. Use our own Prezis or ones shared with us, then replace or add to the simulated file.

## Revisit when

- The export test shows titles or alt text are lost in bulk.
- Prezi publishes an API or an accessibility export.
