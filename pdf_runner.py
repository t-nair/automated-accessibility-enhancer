"""Runs one uploaded PDF through the remediation pipeline in `pipeline/`.

Shared by the local web server (app.py) and the AWS Lambda (lambda_handler.py).
It must not import Flask, because the Lambda image does not carry the website.
"""

import os
import json
import logging
from pathlib import Path

from z_reorder import severity_summary


# the PDF stages that count towards the progress bar, in the order they finish
PDF_STAGES = ("inspect", "image-remediation", "semantic-html", "tagged-pdf-export", "verify")


# the PDF version of process_one_file: gives back whether it worked, a message and steps to try
# output_folder gets the rebuilt PDF and its files, report_folder the <stem>_alt_text report
def process_one_pdf(saved_name, input_folder, output_folder, report_folder, original_filename, progress_callback, language=None):
    # imported here so a server without the PDF packages can still fix PowerPoint files
    try:
        from pipeline.cli import run as remediate_pdf
    except ImportError as e:
        logging.error(f"{saved_name} not processed: the PDF pipeline could not be loaded. Error: {e}")
        return False, "PDF support is not installed on this server.", [
            "Upload a PowerPoint file instead, or ask the accessibility team to install the PDF requirements.",
        ]

    finished_stages = set()

    def on_stage(name, status, message):
        logging.info(f"{saved_name}: PDF stage {name} {status}. {message}")

        if status == "complete" and name in PDF_STAGES:
            finished_stages.add(name)
            progress_callback(len(finished_stages), len(PDF_STAGES))

    stem = os.path.splitext(saved_name)[0]
    output_folder = Path(output_folder)

    try:
        manifest = remediate_pdf(
            Path(input_folder) / saved_name, output_folder, on_stage,
            fallback_title=os.path.splitext(original_filename)[0], language=language,
        )
    except Exception as e:
        logging.exception(f"PDF remediation of {saved_name} stopped. Error: {e}")
        return pdf_error_details(output_folder / (stem + ".remediation.json"), e)

    report_path = os.path.join(report_folder, stem + "_alt_text")

    with open(report_path, "w", encoding="utf-8") as f:
        f.write(pdf_report_text(original_filename, manifest))

    return True, None, None


# the pipeline stops on purpose for things a person has to decide, so explain those plainly
def pdf_error_details(manifest_path, error):
    image_issues = []

    if manifest_path.exists():
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        image_issues = manifest.get("imageRemediation", {}).get("issues", [])

    if image_issues:
        return False, f"This PDF has images whose descriptions need a person to approve them ({len(image_issues)} issue(s)).", [
            "Image descriptions are never published without approval, and approving them is not on this website yet.",
            "Ask the accessibility team to run the PDF command-line tool with an approved image review file.",
        ]

    # these messages come from the pipeline's own checks and say what stopped it.
    # inspect_pdf raises ValueError for an encrypted file or one that is not a PDF.
    if isinstance(error, (RuntimeError, ValueError)):
        return False, f"We could not rebuild this PDF: {error}", [
            "If the PDF is signed or password protected, upload an unsigned, unlocked copy.",
            "Otherwise contact the accessibility team, who can review it by hand.",
        ]

    return False, "Something went wrong while rebuilding this PDF.", [
        "Please try again. If it keeps failing, contact the accessibility team.",
    ]


def pdf_report_text(original_filename, manifest):
    inspection = manifest["inspection"]
    pages_to_review = [page for page in inspection["pages"] if page["needs_review"]]

    # a missing language is serious (WCAG 3.1.1), a page a person must check is moderate
    issues = ["moderate"] * len(pages_to_review)
    language = inspection["language"]

    if not language:
        issues.append("serious")

    lines = [
        severity_summary(issues),
        f"PDF remediation report for {original_filename}",
        f"Title: {inspection['title']}",
        f"Language: {language or '[serious] not set (en assumed)'}",
        "",
    ]

    if manifest.get("reviewRequired"):
        lines.append("A person needs to review this file. It has OCR text or a layout that was inferred.")
    else:
        lines.append("The PDF was rebuilt with tags, a title and a language.")

    lines.append("")

    for page in inspection["pages"]:
        line = f"Page {page['page_number']}: {page['route']}"

        if page["complex_content"]:
            line += ", contains " + ", ".join(page["complex_content"])

        if page["needs_review"]:
            line += " [moderate] (needs review)"

        lines.append(line)

    lines.append("")
    lines.append("Limits of this rebuild:")

    for limitation in manifest["limitations"]:
        lines.append("- " + limitation)

    return "\n".join(lines) + "\n"
