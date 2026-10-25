"""Create a semantic, reflowable HTML reconstruction for text-based PDFs."""

from __future__ import annotations

import html
import re
from pathlib import Path
from urllib.parse import urlparse

import pdfplumber

from .models import DocumentInspection, PageRoute
from .ocr import OcrContent


HEADING_LABELS = {
    "overview",
    "prerequisites",
    "final deliverables",
    "stretch goals",
    "what you will learn",
    "week-by-week plan (july 27 - august 28)",
    "accessibility 101 notes",
    "deque training",
    "basic pdf accessibility 2.0",
    "n8n flow",
}
URL_RE = re.compile(r"https?://[^\s<]+")


def _clean_line(line: str) -> str:
    return (
        line.replace("\ufb01", "fi")
        .replace("\ufb02", "fl")
        .replace("\u2013", "-")
        .replace("\u2014", "-")
        .strip()
    )


def _linkify(value: str) -> str:
    parts: list[str] = []
    cursor = 0
    for match in URL_RE.finditer(value):
        parts.append(html.escape(value[cursor : match.start()]))
        matched = match.group(0)
        url = matched.rstrip(".,)")
        suffix = matched[len(url) :]
        parsed = urlparse(url)
        if parsed.scheme not in {"http", "https"}:
            parts.append(html.escape(matched))
        else:
            parts.append(
                f'<a href="{html.escape(url, quote=True)}">{html.escape(url)}</a>'
                f"{html.escape(suffix)}"
            )
        cursor = match.end()
    parts.append(html.escape(value[cursor:]))
    return "".join(parts)


def _is_heading(line: str) -> bool:
    normalized = line.rstrip(":").strip().lower()
    if normalized in HEADING_LABELS:
        return True
    if re.match(r"^(project\s+\d+|week\s+\d+)(?:\s*:|\b)", normalized):
        return len(line) <= 120
    return line.endswith(":") and len(line) <= 70 and not line.startswith(("-", "•"))


def _logical_entries(lines: list[str]) -> list[tuple[str, str]]:
    entries: list[tuple[str, str]] = []
    pending_kind: str | None = None
    pending_text = ""
    pending_indent = 0
    blank_after_item = False

    def flush() -> None:
        nonlocal pending_kind, pending_text, pending_indent, blank_after_item
        if pending_kind and pending_text:
            entries.append((pending_kind, pending_text))
        pending_kind = None
        pending_text = ""
        pending_indent = 0
        blank_after_item = False

    for raw in lines:
        indent = len(raw) - len(raw.lstrip())
        line = _clean_line(raw)
        if not line or line in {".", ".,", ","}:
            if pending_kind in {"ol-item", "alpha-upper-item", "alpha-lower-item", "ul-item"}:
                blank_after_item = True
            else:
                flush()
            continue
        bullet = re.match(r"^(?:[-•]|\d+[.)]|[A-Za-z][.)])\s+(.*)$", line)
        if bullet:
            flush()
            marker = line.split(maxsplit=1)[0]
            if marker[0].isdigit():
                pending_kind = "ol-item"
            elif marker[0].isalpha():
                pending_kind = "alpha-upper-item" if marker[0].isupper() else "alpha-lower-item"
            else:
                pending_kind = "ul-item"
            pending_text = bullet.group(1).strip()
            pending_indent = indent
            blank_after_item = False
            continue
        if _is_heading(line):
            flush()
            level = "h2" if re.match(r"^(project\s+\d+|week\s+\d+)", line.lower()) else "h3"
            entries.append((level, line.rstrip(":")))
            continue
        if pending_kind in {"ol-item", "alpha-upper-item", "alpha-lower-item", "ul-item"} and indent > pending_indent:
            pending_text += " " + line
            blank_after_item = False
        elif pending_kind == "p" and not blank_after_item:
            pending_text += " " + line
        else:
            flush()
            pending_kind = "p"
            pending_text = line
            pending_indent = indent
    flush()
    return entries


def _semantic_blocks(lines: list[str]) -> list[str]:
    blocks: list[str] = []
    list_kind: str | None = None

    def close_list() -> None:
        nonlocal list_kind
        if list_kind:
            blocks.append("</ol>" if list_kind.startswith("ol") else "</ul>")
            list_kind = None

    for kind, content in _logical_entries(lines):
        if kind.endswith("-item"):
            wanted = {
                "ol-item": "ol",
                "alpha-upper-item": "ol-alpha-upper",
                "alpha-lower-item": "ol-alpha-lower",
                "ul-item": "ul",
            }[kind]
            if list_kind != wanted:
                close_list()
                list_kind = wanted
                blocks.append(
                    {
                        "ol": "<ol>",
                        "ol-alpha-upper": '<ol type="A">',
                        "ol-alpha-lower": '<ol type="a">',
                        "ul": "<ul>",
                    }[wanted]
                )
            blocks.append(f"<li>{_linkify(content)}</li>")
            continue
        close_list()
        if kind in {"h2", "h3"}:
            blocks.append(f"<{kind}>{_linkify(content)}</{kind}>")
        else:
            blocks.append(f"<p>{_linkify(content)}</p>")
    close_list()
    return blocks


def build_semantic_html(
    source_pdf: str | Path,
    inspection: DocumentInspection,
    output_html: str | Path,
    ocr_pages: dict[int, OcrContent] | None = None,
) -> Path:
    ocr_pages = ocr_pages or {}
    blocked = [page for page in inspection.pages if page.route != PageRoute.DIGITAL and page.page_number not in ocr_pages]
    if blocked:
        routes = ", ".join(f"page {page.page_number}: {page.route}" for page in blocked)
        raise RuntimeError(
            "Semantic rebuild stopped because OCR/review providers are unavailable for " + routes
        )
    image_pages = [page.page_number for page in inspection.pages if page.image_count and page.page_number not in ocr_pages]
    if image_pages:
        pages = ", ".join(str(number) for number in image_pages)
        raise RuntimeError(
            "Semantic rebuild stopped because image-purpose and alt-text review are unavailable for "
            f"page(s) {pages}."
        )

    sections: list[str] = []
    with pdfplumber.open(str(source_pdf)) as document:
        for page_number, page in enumerate(document.pages, start=1):
            if page_number in ocr_pages:
                recognized = ocr_pages[page_number]
                # Preserve detected line boundaries; paragraph/column reconstruction is not inferred.
                body = "\n".join(f'<p>{_linkify(line.text)}</p>' for line in recognized.page.lines)
                if not body:
                    body = '<p>No text was recognized on this page. Review the source image.</p>'
                sections.append(
                    f'<section class="source-page"><h2>OCR draft: source page {page_number}</h2>'
                    '<p>Unverified transcription. Check reading order, missing text, tables, and figures.</p>'
                    f'{body}<figure><img style="max-width:100%" src="data:image/png;base64,{recognized.image_data}" '
                    f'alt="Source scan of page {page_number}, provided for visual comparison. Non-text content requires review.">'
                    f'<figcaption>Source page {page_number}: visual reference, not a reviewed text alternative.</figcaption>'
                    '</figure></section>'
                )
                continue
            text = page.extract_text(layout=True) or page.extract_text() or ""
            body = "\n".join(_semantic_blocks(text.splitlines()))
            sections.append(
                f'<section class="source-page" aria-label="Content from source page {page_number}">{body}</section>'
            )

    language = html.escape(inspection.language or "en", quote=True)
    title = html.escape(inspection.title)
    warning = (
        "This is a reflowed accessibility reconstruction. Visual pagination may differ from the source PDF."
    )
    if ocr_pages:
        warning = "OCR DRAFT - HUMAN REVIEW REQUIRED. " + warning + " OCR may omit or misread text; figures and tables have not been remediated."
    document_html = f"""<!doctype html>
<html lang="{language}">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{title} - Accessible Reconstruction</title>
  <style>
    @page {{ size: Letter; margin: 0.72in 0.78in 0.75in; }}
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; color: #17202a; font: 11pt/1.5 Arial, sans-serif; }}
    h1 {{ color: #4b2e83; font-size: 23pt; line-height: 1.15; margin: 0 0 12pt; }}
    h2 {{ color: #4b2e83; font-size: 16pt; margin: 20pt 0 7pt; break-after: avoid; }}
    h3 {{ color: #1f5c78; font-size: 12.5pt; margin: 14pt 0 5pt; break-after: avoid; }}
    p {{ margin: 0 0 7pt; orphans: 3; widows: 3; }}
    ul, ol {{ margin: 4pt 0 9pt 21pt; padding: 0; }}
    li {{ margin: 0 0 4pt; }}
    a {{ color: #005ea8; text-decoration: underline; }}
    .notice {{ border-left: 5px solid #e8b923; background: #fff8d8; padding: 10pt 12pt; margin: 0 0 18pt; }}
    .source-page + .source-page {{ margin-top: 14pt; }}
  </style>
</head>
<body>
  <main>
    <h1>{title}</h1>
    <p class="notice"><strong>About this version:</strong> {html.escape(warning)}</p>
    {''.join(sections)}
  </main>
</body>
</html>
"""
    output = Path(output_html)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(document_html, encoding="utf-8")
    return output
