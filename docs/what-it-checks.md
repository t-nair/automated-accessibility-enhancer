# What it checks and fixes

## PowerPoint (`.pptx`, and `.ppt` through LibreOffice)

| Check | WCAG 2.1 | What happens |
|---|---|---|
| Reading order | 1.3.2 | **Fixed** by role: the title first, then the subtitle, then the content in the author's order, then the date, footer and slide number. Titles typed into a text box are found by a small trained model ([title_research/](../title_research/README.md)). A shape that sits above another but is read after it is reported. |
| Images with no description | 1.1.1 | **Fixed.** Claude on Amazon Bedrock writes a description into the file. A filename (`photo.png`) or a generic name (`Picture 3`) counts as missing. Descriptions a person wrote are never overwritten. |
| Presentation language | 3.1.1, 3.1.2 | **Fixed** when you pick a language on upload: set across every slide, master and layout. A missing, invalid or mixed language is always reported. |
| Slides with no title | 2.4.2 | Reported, by slide. |
| Low text contrast | 1.4.3 | Reported, with each ratio. Colours are resolved through the theme, the shape fill and the slide, layout and master backgrounds. |
| Small text | 1.4.4 | Reported, including sizes inherited from the layout and master (below 9 pt by default). |
| Vague link text | 2.4.4 | Reported: "click here", "read more", bare URLs, and linked pictures with no useful alt text. |

"Reported" means the tool does not guess: a new title, colour, font size or link wording is a design decision, so it is listed for a person to make. [wcag-coverage.md](wcag-coverage.md) covers all 50 Level A and AA criteria and what is, and is not, handled.

## PDF

Every page is inspected and routed as digital, scanned, hybrid or uncertain. Digital text is rebuilt as a **tagged PDF** with a title and language, including a conservative two-column reading order. Scanned and hybrid pages are read with local OCR and always marked for a person to check. Uncertain pages, signed or encrypted files and complex content stop for review rather than being silently rewritten. See [pipeline/README.md](../pipeline/README.md) and [ocr-pathway.md](ocr-pathway.md).

## The report

Every file gets a report that opens with a score out of 100 for the file **as it was uploaded**, and the number of problems at each severity. Each problem is marked critical, serious, moderate or minor, with its WCAG criterion. An excerpt:

```text
Accessibility score: 77/100 (as uploaded)
Issues by severity: 1 critical, 2 serious, 1 moderate, 1 minor

 Changes made (2):
 - Slide 2, reading order - before: Content Placeholder 2, Title 1, after: Title 1, Content Placeholder 2
 - Slide 3, Shape: Picture 3 - alt text before: 'IMG_0412.png', after: 'A bar chart comparing LED and incandescent efficiency.'

 Shape: Picture 3
 - Alt Text: A bar chart comparing LED and incandescent efficiency.
 - [critical] The image had no real description, so this one was generated. Check it (WCAG 1.1.1).

 [moderate] Title, subtitle or footer was out of reading order, now fixed (WCAG 1.3.2): slides 2
 [serious] Slides with no title (WCAG 2.4.2): 4
 [minor] Text below 9 pt (WCAG 1.4.4):
 - Slide 3, Shape: Footnote - 8 pt: 'Source: DOE 2024'
 [serious] Text with too little contrast (WCAG 1.4.3):
 - Slide 5, Shape: Subtitle - 2.31:1, needs 4.5:1: 'Week 3 overview'
```

Under the score, **Changes made** lists every change written into the fixed file (reading order, descriptions, table header rows, language) by slide, with the value before and after, so a reviewer can check each one.

A critical problem costs 10 points, serious 5, moderate 2 and minor 1. Problems the tool fixed still count, so the score describes the original.

## The web app

The main way to use the tool. Drag in one or more `.pptx`, `.ppt` or `.pdf` files (up to 50 MB each), optionally pick the language, and the app:

- queues them and shows a live progress bar for each one
- lists every submission and its status on **Check Status**
- shows the scored report on each file's page, with a download of the fixed file
- explains any failure in plain words with steps to fix it, and offers **Try again**
- lets you delete a submission and its files

Each browser only sees its own submissions, every form is protected by a CSRF token, and if the server restarts, unfinished files are picked up again from their original uploads.

The app itself has been audited with axe-core on every page and state: **no WCAG 2.1 Level A or AA violations** ([accessibility-audit.md](accessibility-audit.md)).
