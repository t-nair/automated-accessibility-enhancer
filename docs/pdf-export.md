# PDF export of fixed decks

Research note for [#40](https://github.com/t-nair/automated-accessibility-enhancer/issues/40). Question: does a PDF exported from a fixed `.pptx` keep tags, title and language?

## Answer

**Export from PowerPoint with "Document structure tags for accessibility" ticked** (File, Save As, PDF, Options; it is on by default). That route keeps tags, alt text, headings and reading order. Set the document title in PowerPoint first, because the tool does not, and the PDF has no title without it.

**LibreOffice: not tested.** It was not installed on the test machine. Do not assume it behaves the same. See "Not verified yet".

## What was tested

PowerPoint (Microsoft 365, Windows), PowerShell/Python COM automation, output inspected with `pypdf`. Decks were the repo's `issue_all_issues_combined.pptx` and `control_no_issues.pptx` after the tool's title move and alt text fixes, plus a deck built with `python-pptx` using a real title placeholder.

| Check | Save As PDF (default) | Export, tags ticked | Export, tags unticked |
|---|---|---|---|
| Tagged PDF (`/MarkInfo`, `/StructTreeRoot`) | Yes | Yes | **No** |
| Picture alt text as `/Figure` `/Alt` | Yes | Yes | No (no tags to hold it) |
| Real title placeholder tagged `/H1` | Yes | Yes | No |
| Title text box (not a placeholder) | Plain `/Textbox`, not a heading | same | No |
| Document title (`/Title`) | Only if set in File, Info | same | same |
| Document language (`/Lang`) | `en` | `en` | `en` |
| Per-run language | Kept as `/Lang` on spans | same | not checked |

## Findings

1. **Tags survive by default.** A plain Save As PDF is tagged. Unticking the tags option removes all of them, so check it before exporting.
2. **Alt text carries over.** The alt text the tool writes appears as `/Alt` on each `/Figure`.
3. **Only real title placeholders become headings.** The tool's title step fixes reading order, but a text box typed as a title is exported as a plain text box, not a heading. A screen reader user gets no heading navigation on that slide. Fix: use the slide layout's title placeholder.
4. **The PDF has no title unless the `.pptx` has one.** The tool does not set the document title, so exports have an empty `/Title`. The PDF still asks viewers to show the title (`DisplayDocTitle`), so they fall back to the file name or blank. Set File, Info, Title before exporting. This is a gap in the tool: WCAG 2.4.2 for the file itself.
5. **Document language is always `en`.** It did not change when text was set to French. Per-run language does carry onto spans, so mixed-language text is marked, but the document default is not taken from the deck.

## Guidance for users

1. Run the deck through the tool and fix what it reports.
2. Make sure each slide title uses the title placeholder.
3. Set File, Info, Title.
4. Save As PDF and leave "Document structure tags for accessibility" ticked. Do not use Print to PDF, which drops tags.
5. Open the PDF in Acrobat's accessibility checker, or PAC, and set the document language there if it is not English.

## Not verified yet

- **LibreOffice export** (File, Export as PDF, "Tagged PDF" option). Not run. To test: export the same fixed deck, then check `/MarkInfo`, `/StructTreeRoot`, `/Title`, `/Lang` and `/Figure` `/Alt` as above.
- **Print to PDF** and other virtual printers: advice in step 4 comes from Microsoft's documentation, not from a test here.
- **Mac PowerPoint** and PowerPoint for the web.
- **Actual screen reader reading** of the PDFs. Only the file structure was inspected.

## Possible follow-up

The tool could set the core document title (and language) from the first slide's title when they are empty. That would fix findings 4 and 5 at the source. Not part of this issue.
