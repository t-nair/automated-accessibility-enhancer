import os
import json
import shutil
import uuid
import sqlite3
import logging
from datetime import datetime, timedelta
from flask import send_from_directory

UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"
DATA_FOLDER = "data"
DATABASE_FILE = "data/submissions.db"
OLD_SUBMISSIONS_FILE = "data/submissions.json"
SECRET_KEY_FILE = "data/secret_key.txt"

# submissions and their files are thrown away after this long, so the disk does not
# fill up on a server that is left running
RETENTION_DAYS = 30


def setup():
    # these folders are not stored in git, so make sure they exist on a fresh copy
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(PROCESSED_FOLDER, exist_ok=True)
    os.makedirs(DATA_FOLDER, exist_ok=True)

    connection = get_connection()
    connection.execute("""
        CREATE TABLE IF NOT EXISTS submissions (
            id TEXT PRIMARY KEY,
            owner_id TEXT,
            original_filename TEXT,
            saved_filename TEXT,
            status TEXT,
            alt_text_filename TEXT,
            error_message TEXT,
            error_steps TEXT,
            retry_count INTEGER,
            submitted_at TEXT,
            language TEXT
        )
    """)

    # databases made before the language option have no such column
    columns = [row["name"] for row in connection.execute("PRAGMA table_info(submissions)")]

    if "language" not in columns:
        connection.execute("ALTER TABLE submissions ADD COLUMN language TEXT")

    connection.commit()
    connection.close()

    import_old_submissions()
    delete_old_submissions()


# the key that signs the cookie has to stay the same between restarts, or everyone
# loses track of the files they submitted
def get_secret_key():
    from_environment = os.environ.get("SECRET_KEY")

    if from_environment:
        return from_environment

    if os.path.exists(SECRET_KEY_FILE):
        with open(SECRET_KEY_FILE, "r") as f:
            return f.read().strip()

    new_key = uuid.uuid4().hex + uuid.uuid4().hex

    with open(SECRET_KEY_FILE, "w") as f:
        f.write(new_key)

    return new_key


def get_connection():
    # a new connection each time, because the background worker runs on its own
    # thread and sqlite connections cannot be shared between threads
    connection = sqlite3.connect(DATABASE_FILE)
    connection.row_factory = sqlite3.Row

    return connection


# the pages expect a dictionary, and the steps are stored as one piece of text
def row_to_submission(row):
    submission = dict(row)

    if submission["error_steps"]:
        submission["error_steps"] = json.loads(submission["error_steps"])
    else:
        submission["error_steps"] = None

    return submission


def add_submission(submission):
    connection = get_connection()
    connection.execute("""
        INSERT INTO submissions
        (id, owner_id, original_filename, saved_filename, status,
         alt_text_filename, error_message, error_steps, retry_count, submitted_at, language)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        submission["id"],
        submission["owner_id"],
        submission["original_filename"],
        submission["saved_filename"],
        submission["status"],
        submission["alt_text_filename"],
        submission["error_message"],
        json.dumps(submission["error_steps"]) if submission["error_steps"] else None,
        submission["retry_count"],
        submission["submitted_at"],
        submission.get("language"),
    ))
    connection.commit()
    connection.close()


def get_submission(submission_id):
    connection = get_connection()
    row = connection.execute("SELECT * FROM submissions WHERE id = ?", (submission_id,)).fetchone()
    connection.close()

    if row is None:
        return None

    return row_to_submission(row)


def get_submissions_for_owner(owner_id):
    connection = get_connection()
    rows = connection.execute(
        "SELECT * FROM submissions WHERE owner_id = ? ORDER BY submitted_at DESC, rowid DESC",
        (owner_id,)
    ).fetchall()
    connection.close()

    submissions = []

    for row in rows:
        submissions.append(row_to_submission(row))

    return submissions


# changes is a dictionary of column names and their new values
def update_submission(submission_id, changes):
    if not changes:
        return

    if "error_steps" in changes:
        steps = changes["error_steps"]
        changes = dict(changes)
        changes["error_steps"] = json.dumps(steps) if steps else None

    assignments = []
    values = []

    for column, value in changes.items():
        assignments.append(column + " = ?")
        values.append(value)

    values.append(submission_id)

    connection = get_connection()
    connection.execute("UPDATE submissions SET " + ", ".join(assignments) + " WHERE id = ?", values)
    connection.commit()
    connection.close()


def delete_submission(submission_id):
    connection = get_connection()
    connection.execute("DELETE FROM submissions WHERE id = ?", (submission_id,))
    connection.commit()
    connection.close()


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

    paths = [
        os.path.join(UPLOAD_FOLDER, names["upload"]),
        os.path.join(PROCESSED_FOLDER, names["copy"]),
        os.path.join(PROCESSED_FOLDER, names["fixed"]),
        os.path.join(PROCESSED_FOLDER, names["report"]),
    ]

    for path in paths:
        if not os.path.exists(path):
            continue

        try:
            os.remove(path)
        except OSError as e:
            logging.warning(f"Could not delete {path}. Error: {e}")

    # a PDF's results are a whole folder (HTML, PDF, manifest, image review)
    pdf_folder = pdf_output_folder(submission["saved_filename"])

    if os.path.exists(pdf_folder):
        shutil.rmtree(pdf_folder, ignore_errors=True)


def is_pdf(filename):
    return filename.lower().endswith(".pdf")


def pdf_output_folder(saved_name):
    return os.path.join(PROCESSED_FOLDER, os.path.splitext(saved_name)[0] + "_pdf")


# the folder holding the original upload, or None if it is gone. Older versions moved
# it into the processed folder once it was done, so look there too.
def original_upload_folder(saved_name):
    for folder in (UPLOAD_FOLDER, PROCESSED_FOLDER):
        if os.path.exists(os.path.join(folder, saved_name)):
            return folder

    return None


# where the finished file is kept, and the name it is downloaded as
def processed_file_for(submission):
    saved_stem = os.path.splitext(submission["saved_filename"])[0]
    original_stem = os.path.splitext(submission["original_filename"])[0]

    if is_pdf(submission["saved_filename"]):
        return pdf_output_folder(submission["saved_filename"]), saved_stem + ".accessible.pdf", "accessible_" + original_stem + ".pdf"

    # the result is always a .pptx, even when an old .ppt was uploaded
    return PROCESSED_FOLDER, saved_stem + "_updated.pptx", "accessible_" + original_stem + ".pptx"


def save_upload(uploaded_file, saved_name):
    uploaded_file.save(os.path.join(UPLOAD_FOLDER, saved_name))


def upload_exists(saved_name):
    return original_upload_folder(saved_name) is not None


def read_report(submission):
    if not submission["alt_text_filename"]:
        return None

    path = os.path.join(PROCESSED_FOLDER, submission["alt_text_filename"])

    if not os.path.exists(path):
        return None

    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def send_processed_file(submission):
    folder, filename, download_name = processed_file_for(submission)

    return send_from_directory(folder, filename, as_attachment=True, download_name=download_name)


def processed_file_is_ready(submission):
    folder, filename, _ = processed_file_for(submission)

    return os.path.exists(os.path.join(folder, filename))


# the project used to keep submissions in a JSON file, so bring those across once
def import_old_submissions():
    if not os.path.exists(OLD_SUBMISSIONS_FILE):
        return

    connection = get_connection()
    already_there = connection.execute("SELECT COUNT(*) FROM submissions").fetchone()[0]
    connection.close()

    if already_there > 0:
        return

    with open(OLD_SUBMISSIONS_FILE, "r") as f:
        old_submissions = json.load(f)

    for old in old_submissions:
        add_submission({
            "id": old["id"],
            # these were submitted before there were separate visitors
            "owner_id": None,
            "original_filename": old.get("original_filename", ""),
            "saved_filename": old.get("saved_filename", ""),
            "status": old.get("status", "error"),
            "alt_text_filename": old.get("alt_text_filename"),
            "error_message": old.get("error_message"),
            "error_steps": old.get("error_steps"),
            "retry_count": old.get("retry_count", 0),
            "submitted_at": old.get("submitted_at", ""),
        })

    logging.info(f"Copied {len(old_submissions)} older submission(s) from {OLD_SUBMISSIONS_FILE} into the database.")


# anything waiting or half done when the app stopped, oldest first, so it can be
# queued again: the queue only lives in memory
def unfinished_submissions():
    connection = get_connection()
    rows = connection.execute(
        "SELECT id, saved_filename FROM submissions WHERE status IN ('queued', 'processing') ORDER BY rowid"
    ).fetchall()
    connection.close()

    return [dict(row) for row in rows]


def delete_old_submissions():
    cutoff = datetime.now() - timedelta(days=RETENTION_DAYS)
    connection = get_connection()
    rows = connection.execute("SELECT * FROM submissions").fetchall()
    connection.close()

    removed = 0

    for row in rows:
        submission = row_to_submission(row)

        try:
            submitted = datetime.strptime(submission["submitted_at"], "%Y-%m-%d %H:%M")
        except ValueError:
            continue

        if submitted < cutoff:
            delete_submission_files(submission)
            delete_submission(submission["id"])
            removed += 1

    if removed:
        logging.info(f"Deleted {removed} submission(s) older than {RETENTION_DAYS} days.")
