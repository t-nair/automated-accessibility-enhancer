"""AWS Lambda entry point for the accessibility pipeline.

A deck lands in the uploads bucket, S3 notifies this function, and the same
pipeline the website uses runs against it. The fixed deck and its report are
written to the processed bucket under the same key, alongside a small JSON file
saying whether it worked, so the frontend can show the result without needing
to talk to the function itself.

Nothing here is specific to Lambda beyond moving files in and out of S3. The
actual work is `z_reorder.process_one_file`, exactly as `app.py` calls it.
"""

import os
import json
import shutil
import logging
import urllib.parse

import boto3

from z_reorder import process_one_file

# z_reorder calls logging.basicConfig(filename=...) when it is imported, which
# is a no-op here: the Lambda runtime has already put a handler on the root
# logger, so no file is opened and every log line goes to CloudWatch instead.
# The level still has to be set, because Lambda defaults the root logger to
# WARNING and the pipeline says everything useful at INFO.
logging.getLogger().setLevel(os.environ.get("LOG_LEVEL", "INFO"))

PROCESSED_BUCKET = os.environ.get("PROCESSED_BUCKET", "")

# the only writable place in a Lambda container
WORK_ROOT = "/tmp/work"

ALLOWED_EXTENSIONS = (".pptx", ".ppt")

# process_one_file also moves the original upload into the output folder, but
# that copy is already sitting in the uploads bucket, so it is not sent back.
RESULT_SUFFIXES = ("_updated.pptx", "_alt_text")

s3 = boto3.client("s3")


def status_key(key):
    return key + ".status.json"


def write_status(key, was_successful, message, steps):
    """Put a small JSON file next to the results saying how the deck went.

    The website reads this instead of guessing from which files exist, so a
    failure explains itself the same way it does in the browser today.
    """
    body = json.dumps({
        "source_key": key,
        "status": "done" if was_successful else "error",
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
        logging.warning(f"Ignoring {key}: not a .pptx or .ppt file.")
        return None

    input_folder = os.path.join(WORK_ROOT, "input")
    output_folder = os.path.join(WORK_ROOT, "output")

    # a warm container reuses /tmp, so start from an empty pair of folders every
    # time rather than picking up whatever the last invocation left behind
    shutil.rmtree(WORK_ROOT, ignore_errors=True)
    os.makedirs(input_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)

    try:
        download_path = os.path.join(input_folder, filename)
        s3.download_file(bucket, key, download_path)
        logging.info(f"Downloaded s3://{bucket}/{key}")

        was_successful, message, steps = process_one_file(filename, input_folder, output_folder)

        if was_successful:
            upload_results(key, output_folder)

        write_status(key, was_successful, message, steps)

        return was_successful, message, steps
    finally:
        shutil.rmtree(WORK_ROOT, ignore_errors=True)


def upload_results(key, output_folder):
    """Copy the fixed deck and its report to the processed bucket.

    They keep the uploaded key's folder, so a deck at `abc123/lecture.pptx`
    produces `abc123/lecture_updated.pptx` and `abc123/lecture_alt_text`.
    """
    prefix = os.path.dirname(key)

    for produced in sorted(os.listdir(output_folder)):
        if not produced.endswith(RESULT_SUFFIXES):
            continue

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
