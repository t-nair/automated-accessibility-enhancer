"""AWS Lambda entry point for the accessibility pipeline.

A deck lands in the uploads bucket, S3 notifies this function, and the same
pipeline the website uses runs against it. The fixed deck and its report are
written to the processed bucket under the same key, alongside a small JSON file
saying whether it worked, so the frontend can show the result without needing
to talk to the function itself.

Nothing here is specific to Lambda beyond moving files in and out of S3. The
actual work is `z_reorder.process_one_file` for a deck and
`pdf_runner.process_one_pdf` for a PDF, exactly as `app.py` calls them.
"""

import os
import json
import time
import shutil
import logging
import urllib.parse

import boto3

from z_reorder import process_one_file
import pdf_runner
from pdf_runner import process_one_pdf

# z_reorder calls logging.basicConfig(filename=...) when it is imported, which
# is a no-op here: the Lambda runtime has already put a handler on the root
# logger, so no file is opened and every log line goes to CloudWatch instead.
# The level still has to be set, because Lambda defaults the root logger to
# WARNING and the pipeline says everything useful at INFO.
logging.getLogger().setLevel(os.environ.get("LOG_LEVEL", "INFO"))

PROCESSED_BUCKET = os.environ.get("PROCESSED_BUCKET", "")

# the table the website reads. The status file below is still written next to the
# results, but the site watches the table, because that is also where the progress
# bar gets its numbers from while a deck is still being worked on.
SUBMISSIONS_TABLE = os.environ.get("SUBMISSIONS_TABLE", "")

# the only writable place in a Lambda container
WORK_ROOT = "/tmp/work"

ALLOWED_EXTENSIONS = (".pptx", ".ppt", ".pdf")

# process_one_file also moves the original upload into the output folder, but
# that copy is already sitting in the uploads bucket, so it is not sent back.
# The .json is the report arranged by slide, which the details page shows as its
# summary and slide by slide view; without it the page falls back to plain text.
RESULT_SUFFIXES = ("_updated.pptx", "_alt_text", "_alt_text.json")

# a PDF's results go in a "<name>_pdf" folder next to its report, the same layout
# storage_local.py uses, so storage_aws.py can find them by name. The extracted
# images and the image review file stay behind: nothing on the site reads them.
PDF_RESULT_SUFFIXES = (".accessible.pdf", ".accessible.html", ".remediation.json")


def is_pdf(filename):
    return filename.lower().endswith(".pdf")


def pdf_folder_name(filename):
    return os.path.splitext(filename)[0] + "_pdf"

s3 = boto3.client("s3")

_table = None


def get_table():
    """The submissions table, or None when the function is running without one."""
    global _table

    if not SUBMISSIONS_TABLE:
        return None

    if _table is None:
        _table = boto3.resource("dynamodb").Table(SUBMISSIONS_TABLE)

    return _table


def submission_id_from_key(key):
    """The website saves a deck as "<id>_<original name>", so the id is the prefix.

    Returns None for anything that does not look like one of ours, such as a file
    put in the bucket by hand, so the pipeline still runs but no row is touched.
    """
    name = os.path.basename(key)

    if "_" not in name:
        return None

    return name.split("_", 1)[0]


def update_row(submission_id, changes):
    """Write to the submission's row, and never fail the deck over it.

    A caption that has been generated matters more than the progress bar being
    right, so a table problem is logged and the run carries on.
    """
    table = get_table()

    if table is None or submission_id is None:
        return

    assignments = []
    names = {}
    values = {}
    number = 0

    for field, value in changes.items():
        number += 1
        names[f"#f{number}"] = field
        assignments.append(f"#f{number} = :v{number}")
        values[f":v{number}"] = value

    try:
        table.update_item(
            Key={"id": submission_id},
            UpdateExpression="SET " + ", ".join(assignments),
            ExpressionAttributeNames=names,
            ExpressionAttributeValues=values,
        )
    except Exception as e:
        logging.warning(f"Could not update submission {submission_id} in {SUBMISSIONS_TABLE}. Error: {e}")


def read_language(submission_id):
    """The language chosen on the upload form, or None to leave the deck's own."""
    table = get_table()

    if table is None or submission_id is None:
        return None

    try:
        item = table.get_item(Key={"id": submission_id}).get("Item") or {}
    except Exception as e:
        logging.warning(f"Could not read submission {submission_id} from {SUBMISSIONS_TABLE}. Error: {e}")
        return None

    return item.get("language")


def status_key(key):
    return key + ".status.json"


def write_status(key, was_successful, message, steps, status=None):
    """Put a small JSON file next to the results saying how the deck went.

    The website reads this instead of guessing from which files exist, so a
    failure explains itself the same way it does in the browser today.
    """
    body = json.dumps({
        "source_key": key,
        "status": status or ("done" if was_successful else "error"),
        "error_message": message,
        "error_steps": steps,
    }, indent=2)

    s3.put_object(
        Bucket=PROCESSED_BUCKET,
        Key=status_key(key),
        Body=body.encode("utf-8"),
        ContentType="application/json",
    )


def process_one_object(bucket, key):
    """Run one deck through the pipeline. Returns the (ok, message, steps) triple."""
    filename = os.path.basename(key)

    if not filename.lower().endswith(ALLOWED_EXTENSIONS):
        logging.warning(f"Ignoring {key}: not a .pptx, .ppt or .pdf file.")
        return None

    input_folder = os.path.join(WORK_ROOT, "input")
    output_folder = os.path.join(WORK_ROOT, "output")

    # a warm container reuses /tmp, so start from an empty pair of folders every
    # time rather than picking up whatever the last invocation left behind
    shutil.rmtree(WORK_ROOT, ignore_errors=True)
    os.makedirs(input_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)

    submission_id = submission_id_from_key(key)

    # started_at is how the website tells a run that is still going from one that
    # died without saying so, since nothing else is watching this function
    update_row(submission_id, {
        "status": "processing",
        "started_at": int(time.time()),
        "captions_done": 0,
        "captions_total": 0,
    })

    def report_progress(captions_done, total_to_caption):
        update_row(submission_id, {
            "captions_done": captions_done,
            "captions_total": total_to_caption,
        })

    try:
        download_path = os.path.join(input_folder, filename)
        s3.download_file(bucket, key, download_path)
        logging.info(f"Downloaded s3://{bucket}/{key}")

        language = read_language(submission_id)

        status = None

        if is_pdf(filename):
            # the upload is saved as "<id>_<original name>", so drop the id for the title
            original_filename = filename.split("_", 1)[1] if submission_id else filename
            pdf_output = os.path.join(output_folder, pdf_folder_name(filename))
            was_successful, message, steps = process_one_pdf(
                filename, input_folder, pdf_output, output_folder,
                original_filename, report_progress, language,
                image_decisions=download_image_decisions(bucket, key, input_folder),
            )

            # stopped for a person to approve image descriptions, so the website
            # needs the review file and pictures to show them
            if not was_successful and pdf_runner.review_is_pending(pdf_output, filename):
                upload_review(key, pdf_output)
                status = pdf_runner.NEEDS_REVIEW
        else:
            was_successful, message, steps = process_one_file(
                filename, input_folder, output_folder, report_progress, language
            )

        if was_successful:
            upload_results(key, output_folder)

        write_status(key, was_successful, message, steps, status)
        write_outcome(submission_id, key, was_successful, message, steps, status)

        return was_successful, message, steps
    except Exception as e:
        logging.error(f"Unexpected problem while processing {key}. Error: {e}")
        if is_pdf(filename):
            second_step = "If it fails again, contact the accessibility team."
        else:
            second_step = "If it fails again, open the file in PowerPoint and save it as a new .pptx."

        write_outcome(submission_id, key, False, "Something went wrong while processing this file.", [
            "Try again from the submission page.",
            second_step,
        ])
        raise
    finally:
        shutil.rmtree(WORK_ROOT, ignore_errors=True)


def write_outcome(submission_id, key, was_successful, message, steps, status=None):
    """Put the finished state on the row the website reads."""
    changes = {"status": status or ("done" if was_successful else "error")}

    if was_successful:
        name_without_extension = os.path.splitext(os.path.basename(key))[0]
        changes["alt_text_filename"] = name_without_extension + "_alt_text"
    else:
        changes["error_message"] = message or "Something went wrong while processing this file."
        changes["error_steps"] = steps or ["Try again from the submission page."]

    update_row(submission_id, changes)


def download_image_decisions(bucket, key, input_folder):
    """A reviewer's decisions, saved by the website next to the PDF, or None.

    There are none on a first run. Without s3:ListBucket a missing object comes
    back as access denied rather than not found, so any failure means none.
    """
    name = pdf_runner.decisions_file_name(os.path.basename(key))
    prefix = os.path.dirname(key)
    path = os.path.join(input_folder, name)

    try:
        s3.download_file(bucket, name if not prefix else prefix + "/" + name, path)
    except Exception:
        return None

    logging.info(f"Using the reviewer's image decisions for {key}")
    return path


def upload_review(key, pdf_output):
    """Copy the image review file and its pictures to where storage_aws.py reads them."""
    filename = os.path.basename(key)
    prefix = os.path.dirname(key)
    folder = pdf_folder_name(filename)
    review = pdf_runner.read_review(pdf_output, filename)
    images = pdf_runner.images_folder_name(filename)

    to_upload = [pdf_runner.review_file_name(filename)]
    to_upload.extend(images + "/" + name for name in pdf_runner.review_image_names(review))

    for produced in to_upload:
        source = os.path.join(pdf_output, produced)

        if not os.path.exists(source):
            continue

        destination = folder + "/" + produced
        destination = destination if not prefix else prefix + "/" + destination
        s3.upload_file(source, PROCESSED_BUCKET, destination)

    logging.info(f"Uploaded the image review for {key} ({len(to_upload)} files)")


def upload_results(key, output_folder):
    """Copy the fixed deck and its report to the processed bucket.

    They keep the uploaded key's folder, so a deck at `abc123/lecture.pptx`
    produces `abc123/lecture_updated.pptx` and `abc123/lecture_alt_text`, and a
    PDF at `abc123/notes.pdf` produces `abc123/notes_alt_text` and
    `abc123/notes_pdf/notes.accessible.pdf`.
    """
    prefix = os.path.dirname(key)
    to_upload = [name for name in os.listdir(output_folder) if name.endswith(RESULT_SUFFIXES)]

    pdf_folder = pdf_folder_name(os.path.basename(key))

    if is_pdf(key) and os.path.isdir(os.path.join(output_folder, pdf_folder)):
        for name in os.listdir(os.path.join(output_folder, pdf_folder)):
            if name.endswith(PDF_RESULT_SUFFIXES):
                to_upload.append(pdf_folder + "/" + name)

    for produced in sorted(to_upload):
        destination = produced if not prefix else prefix + "/" + produced

        s3.upload_file(
            os.path.join(output_folder, produced),
            PROCESSED_BUCKET,
            destination,
        )
        logging.info(f"Uploaded s3://{PROCESSED_BUCKET}/{destination}")


def lambda_handler(event, context):
    if not PROCESSED_BUCKET:
        raise RuntimeError("PROCESSED_BUCKET is not set, so there is nowhere to put the results.")

    results = []

    for record in event.get("Records", []):
        bucket = record["s3"]["bucket"]["name"]

        # S3 sends the key form-encoded, so a deck with a space in its name
        # arrives as "my+lecture.pptx" and has to be turned back first
        key = urllib.parse.unquote_plus(record["s3"]["object"]["key"])

        outcome = process_one_object(bucket, key)

        if outcome is None:
            results.append({"key": key, "status": "skipped"})
            continue

        was_successful, message, _ = outcome
        results.append({
            "key": key,
            "status": "done" if was_successful else "error",
            "error_message": message,
        })

    return {"processed": results}
