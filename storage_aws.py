import os
import time
import logging
from datetime import datetime, timedelta
from flask import redirect

import boto3
from boto3.dynamodb.conditions import Key

UPLOADS_BUCKET = os.environ.get("UPLOADS_BUCKET", "")
PROCESSED_BUCKET = os.environ.get("PROCESSED_BUCKET", "")
SUBMISSIONS_TABLE = os.environ.get("SUBMISSIONS_TABLE", "")
SECRET_KEY_PARAMETER = os.environ.get("SECRET_KEY_PARAMETER", "")

# the index that lets one visitor's submissions be looked up without reading the
# whole table. Its name has to match the one Terraform creates.
OWNER_INDEX = "owner_id-submitted_at-index"

# S3 lifecycle rules and the table's own time to live delete old submissions, so
# nothing here has to sweep them up. This only has to agree with the Terraform.
RETENTION_DAYS = 30

# how long a download link stays valid. Long enough to click, short enough that a
# copied link is not a lasting way in.
DOWNLOAD_LINK_SECONDS = 300

# a submission left processing for longer than this is assumed dead, because the
# Lambda that had it cannot still be running
STUCK_AFTER_SECONDS = int(os.environ.get("PIPELINE_TIMEOUT_SECONDS", "900")) + 120

# built on first use so importing this module needs no credentials, which keeps
# the tests free of AWS
_s3 = None
_table = None


def get_s3():
    global _s3

    if _s3 is None:
        _s3 = boto3.client("s3")

    return _s3


def get_table():
    global _table

    if _table is None:
        _table = boto3.resource("dynamodb").Table(SUBMISSIONS_TABLE)

    return _table


def setup():
    missing = []

    for name in ("UPLOADS_BUCKET", "PROCESSED_BUCKET", "SUBMISSIONS_TABLE"):
        if not os.environ.get(name):
            missing.append(name)

    if missing:
        raise RuntimeError("Running on AWS storage needs these set: " + ", ".join(missing))

    logging.info(f"Using S3 buckets {UPLOADS_BUCKET} and {PROCESSED_BUCKET}, table {SUBMISSIONS_TABLE}.")


# the cookie signing key is shared by every instance, so it comes from Parameter
# Store rather than being made up locally. Instances would otherwise each sign
# with their own key and visitors would lose their list on every other request.
def get_secret_key():
    from_environment = os.environ.get("SECRET_KEY")

    if from_environment:
        return from_environment

    if not SECRET_KEY_PARAMETER:
        raise RuntimeError("Set SECRET_KEY_PARAMETER or SECRET_KEY so the session cookie can be signed.")

    ssm = boto3.client("ssm")
    answer = ssm.get_parameter(Name=SECRET_KEY_PARAMETER, WithDecryption=True)

    return answer["Parameter"]["Value"]


def expires_at():
    return int(time.time()) + (RETENTION_DAYS * 24 * 60 * 60)


# DynamoDB will not store an empty string, and leaves a missing field out of the
# item entirely, so the pages get the same shape they get from sqlite
def item_to_submission(item):
    submission = {
        "id": item["id"],
        "owner_id": item.get("owner_id"),
        "original_filename": item.get("original_filename", ""),
        "saved_filename": item.get("saved_filename", ""),
        "status": item.get("status", "error"),
        "alt_text_filename": item.get("alt_text_filename"),
        "error_message": item.get("error_message"),
        "error_steps": item.get("error_steps"),
        "retry_count": int(item.get("retry_count", 0)),
        "submitted_at": item.get("submitted_at", ""),
        "captions_done": int(item.get("captions_done", 0)),
        "captions_total": int(item.get("captions_total", 0)),
    }

    if item.get("status") == "processing" and is_stuck(item):
        mark_stuck(submission)

    return submission


def is_stuck(item):
    started = item.get("started_at")

    if not started:
        return False

    return (time.time() - float(started)) > STUCK_AFTER_SECONDS


# nothing is watching the pipeline from the outside, so a submission whose run
# died silently is spotted here, the next time somebody looks at it
def mark_stuck(submission):
    message = "This file was being worked on for longer than we allow, so it was stopped."
    steps = ["Upload the presentation again.", "If it fails again, save it from PowerPoint first and upload that copy."]

    update_submission(submission["id"], {
        "status": "error",
        "error_message": message,
        "error_steps": steps,
    })

    submission["status"] = "error"
    submission["error_message"] = message
    submission["error_steps"] = steps

    logging.warning(f"Submission {submission['id']} was still processing after {STUCK_AFTER_SECONDS}s, marked as failed.")


def add_submission(submission):
    item = {
        "id": submission["id"],
        "original_filename": submission["original_filename"],
        "saved_filename": submission["saved_filename"],
        "status": submission["status"],
        "retry_count": submission["retry_count"],
        "submitted_at": submission["submitted_at"],
        "expires_at": expires_at(),
    }

    # the index needs a value to sort on, and DynamoDB skips items missing the key
    if submission["owner_id"]:
        item["owner_id"] = submission["owner_id"]

    for name in ("alt_text_filename", "error_message", "error_steps"):
        if submission.get(name):
            item[name] = submission[name]

    get_table().put_item(Item=item)


def get_submission(submission_id):
    answer = get_table().get_item(Key={"id": submission_id})
    item = answer.get("Item")

    if item is None:
        return None

    return item_to_submission(item)


def get_submissions_for_owner(owner_id):
    answer = get_table().query(
        IndexName=OWNER_INDEX,
        KeyConditionExpression=Key("owner_id").eq(owner_id),
        ScanIndexForward=False,
    )

    submissions = []

    for item in answer.get("Items", []):
        submissions.append(item_to_submission(item))

    return submissions


def update_submission(submission_id, changes):
    if not changes:
        return

    # status is a reserved word in DynamoDB, so every name is passed separately
    # rather than written into the expression
    assignments = []
    remove = []
    names = {}
    values = {}
    number = 0

    for field, value in changes.items():
        number += 1
        placeholder = f"#f{number}"
        names[placeholder] = field

        if value is None or value == "":
            remove.append(placeholder)
            continue

        assignments.append(f"{placeholder} = :v{number}")
        values[f":v{number}"] = value

    expression = ""

    if assignments:
        expression = "SET " + ", ".join(assignments)

    if remove:
        expression = (expression + " REMOVE " + ", ".join(remove)).strip()

    arguments = {
        "Key": {"id": submission_id},
        "UpdateExpression": expression,
        "ExpressionAttributeNames": names,
    }

    if values:
        arguments["ExpressionAttributeValues"] = values

    get_table().update_item(**arguments)


def delete_submission(submission_id):
    get_table().delete_item(Key={"id": submission_id})


def processed_names(saved_name):
    name_without_extension = os.path.splitext(saved_name)[0]

    return {
        "upload": saved_name,
        "copy": saved_name,
        "fixed": name_without_extension + "_updated.pptx",
        "report": name_without_extension + "_alt_text",
    }


def delete_submission_files(submission):
    names = processed_names(submission["saved_filename"])
    s3 = get_s3()

    delete_object(s3, UPLOADS_BUCKET, names["upload"])

    for key in ("copy", "fixed", "report"):
        delete_object(s3, PROCESSED_BUCKET, names[key])


def delete_object(s3, bucket, key):
    try:
        s3.delete_object(Bucket=bucket, Key=key)
    except Exception as e:
        logging.warning(f"Could not delete s3://{bucket}/{key}. Error: {e}")


# putting the deck in the uploads bucket is also what starts the pipeline: the
# bucket tells the Lambda about the new object
def save_upload(uploaded_file, saved_name):
    get_s3().upload_fileobj(uploaded_file, UPLOADS_BUCKET, saved_name)


def upload_exists(saved_name):
    return object_exists(UPLOADS_BUCKET, saved_name)


def object_exists(bucket, key):
    try:
        get_s3().head_object(Bucket=bucket, Key=key)
        return True
    except Exception:
        return False


def read_report(submission):
    if not submission["alt_text_filename"]:
        return None

    try:
        answer = get_s3().get_object(Bucket=PROCESSED_BUCKET, Key=submission["alt_text_filename"])
    except Exception as e:
        logging.error(f"Could not read the report for submission {submission['id']}. Error: {e}")
        return None

    return answer["Body"].read().decode("utf-8")


# the browser fetches the deck from S3 itself, so a 250 MB download never goes
# through the web server
def send_processed_file(submission):
    names = processed_names(submission["saved_filename"])
    # the result is always a .pptx, even when an old .ppt was uploaded
    download_name = "accessible_" + os.path.splitext(submission["original_filename"])[0] + ".pptx"

    link = get_s3().generate_presigned_url(
        "get_object",
        Params={
            "Bucket": PROCESSED_BUCKET,
            "Key": names["fixed"],
            "ResponseContentDisposition": f'attachment; filename="{download_name}"',
        },
        ExpiresIn=DOWNLOAD_LINK_SECONDS,
    )

    return redirect(link)


def processed_file_is_ready(submission):
    names = processed_names(submission["saved_filename"])

    return object_exists(PROCESSED_BUCKET, names["fixed"])


# copying the object over itself makes S3 announce it again, which is what starts
# the pipeline. Nothing else in the bucket changes.
def restart_processing(saved_name):
    get_s3().copy_object(
        Bucket=UPLOADS_BUCKET,
        Key=saved_name,
        CopySource={"Bucket": UPLOADS_BUCKET, "Key": saved_name},
        MetadataDirective="REPLACE",
        Metadata={"retried_at": str(int(time.time()))},
    )

    logging.info(f"Asked the pipeline to run again for {saved_name}.")


# there is no in-memory queue to lose, and a run that dies is caught by the check
# in item_to_submission, so there is nothing to recover at startup
def recover_stuck_submissions():
    return


# S3 lifecycle rules and the table's time to live do this, so the web server does
# not have to walk the whole table on every restart
def delete_old_submissions():
    return


def retention_cutoff():
    return datetime.now() - timedelta(days=RETENTION_DAYS)
