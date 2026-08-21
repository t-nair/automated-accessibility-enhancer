import os
import json
import uuid
import queue
import sqlite3
import logging
import threading
from datetime import datetime, timedelta
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, jsonify, session
from werkzeug.utils import secure_filename
from z_reorder import process_one_file

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"
DATA_FOLDER = "data"
DATABASE_FILE = "data/submissions.db"
OLD_SUBMISSIONS_FILE = "data/submissions.json"
SECRET_KEY_FILE = "data/secret_key.txt"
ALLOWED_EXTENSIONS = (".pptx", ".ppt")
MAX_FILE_SIZE_MB = 50

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# these folders are not stored in git, so make sure they exist on a fresh copy
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)
os.makedirs(DATA_FOLDER, exist_ok=True)

# how far along each file being processed right now is. This is only kept in memory,
# because it is not worth saving to disk once the file is finished.
processing_progress = {}

# files wait here to be processed. One worker takes them one at a time, so several
# uploads can never load the captioning model at once and run the GPU out of memory.
work_queue = queue.Queue()


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


app.secret_key = get_secret_key()
app.permanent_session_lifetime = timedelta(days=30)


def get_connection():
    # a new connection each time, because the background worker runs on its own
    # thread and sqlite connections cannot be shared between threads
    connection = sqlite3.connect(DATABASE_FILE)
    connection.row_factory = sqlite3.Row

    return connection


def setup_database():
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
            submitted_at TEXT
        )
    """)
    connection.commit()
    connection.close()


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
         alt_text_filename, error_message, error_steps, retry_count, submitted_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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


def delete_submission(submission_id):
    connection = get_connection()
    connection.execute("DELETE FROM submissions WHERE id = ?", (submission_id,))
    connection.commit()
    connection.close()


def delete_submission_files(submission):
    saved_name = submission["saved_filename"]
    name_without_extension = os.path.splitext(saved_name)[0]

    paths = [
        os.path.join(UPLOAD_FOLDER, saved_name),
        os.path.join(PROCESSED_FOLDER, saved_name),
        os.path.join(PROCESSED_FOLDER, name_without_extension + "_updated.pptx"),
        os.path.join(PROCESSED_FOLDER, name_without_extension + "_alt_text"),
    ]

    for path in paths:
        if not os.path.exists(path):
            continue

        try:
            os.remove(path)
        except OSError as e:
            logging.warning(f"Could not delete {path}. Error: {e}")


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


setup_database()
import_old_submissions()


# each browser gets a random id in a cookie, which is how one person's list of files
# is kept separate from everyone else's. There are no accounts or passwords.
def get_owner_id():
    if "owner_id" not in session:
        session["owner_id"] = str(uuid.uuid4())
        session.permanent = True

    return session["owner_id"]


# returns nothing if the submission belongs to someone else, so the pages treat it
# the same as one that does not exist
def get_my_submission(submission_id):
    submission = get_submission(submission_id)

    if submission is None:
        return None

    if submission["owner_id"] != get_owner_id():
        return None

    return submission


def file_is_too_big(file):
    # move to the end of the file to check its size, then move back
    file.seek(0, os.SEEK_END)
    size_mb = file.tell() / (1024 * 1024)
    file.seek(0)

    return size_mb > MAX_FILE_SIZE_MB


def set_submission_status(submission_id, new_status):
    update_submission(submission_id, {"status": new_status})


def process_one_submission(submission_id, saved_name):
    def update_progress(captions_done, total_to_caption):
        processing_progress[submission_id] = {"done": captions_done, "total": total_to_caption}

    update_progress(0, 0)
    set_submission_status(submission_id, "processing")

    was_successful, error_message, error_steps = process_one_file(
        saved_name, UPLOAD_FOLDER, PROCESSED_FOLDER, update_progress
    )

    if not was_successful:
        logging.error(f"Submission {submission_id} ({saved_name}) failed processing: {error_message}")

    if was_successful:
        alt_text_filename = os.path.splitext(saved_name)[0] + "_alt_text"
    else:
        alt_text_filename = None

    update_submission(submission_id, {
        "status": "done" if was_successful else "error",
        "error_message": error_message,
        "error_steps": error_steps,
        "alt_text_filename": alt_text_filename,
    })

    processing_progress.pop(submission_id, None)


def worker_loop():
    while True:
        submission_id, saved_name = work_queue.get()

        try:
            process_one_submission(submission_id, saved_name)
        except Exception as e:
            # a crash here must not kill the worker, or nothing else would ever process
            logging.error(f"Unexpected problem while processing submission {submission_id}. Error: {e}")
            set_submission_status(submission_id, "error")

        work_queue.task_done()


# daemon means this thread does not keep the app running when it is shut down
worker_thread = threading.Thread(target=worker_loop, daemon=True)
worker_thread.start()


def start_processing(submission_id, saved_name):
    work_queue.put((submission_id, saved_name))


@app.route("/")
def home():
    return render_template("index.html", max_file_size_mb=MAX_FILE_SIZE_MB)


# gives back the new submission id, or None and a message saying what was wrong
def save_one_upload(uploaded_file):
    filename = uploaded_file.filename

    if not filename.lower().endswith(ALLOWED_EXTENSIONS):
        logging.warning(f"Upload rejected: {filename} is not a .pptx or .ppt file.")
        return None, f"{filename}: only .pptx and .ppt files are accepted."

    if file_is_too_big(uploaded_file):
        logging.warning(f"Upload rejected: {filename} is over the {MAX_FILE_SIZE_MB}MB size limit.")
        return None, f"{filename}: too large (the limit is {MAX_FILE_SIZE_MB} MB)."

    submission_id = str(uuid.uuid4())[:8]
    saved_name = submission_id + "_" + secure_filename(filename)
    save_path = os.path.join(app.config["UPLOAD_FOLDER"], saved_name)

    try:
        uploaded_file.save(save_path)
    except OSError as e:
        logging.error(f"Could not save uploaded file {filename} to {save_path}. Error: {e}")
        return None, f"{filename}: we could not save it. Please try again."

    add_submission({
        "id": submission_id,
        "owner_id": get_owner_id(),
        "original_filename": filename,
        "saved_filename": saved_name,
        "status": "queued",
        "alt_text_filename": None,
        "error_message": None,
        "error_steps": None,
        "retry_count": 0,
        "submitted_at": datetime.now().strftime("%Y-%m-%d %H:%M")
    })

    # the pipeline can take a while, so files are queued and worked through in the
    # background instead of making the browser wait
    start_processing(submission_id, saved_name)

    return submission_id, None


@app.route("/upload", methods=["POST"])
def upload():
    uploaded_files = request.files.getlist("presentation")
    chosen_files = []

    for uploaded_file in uploaded_files:
        if uploaded_file is not None and uploaded_file.filename != "":
            chosen_files.append(uploaded_file)

    if len(chosen_files) == 0:
        logging.warning("Upload rejected: no file was selected.")
        flash("Please choose at least one file before submitting.")
        return redirect(url_for("home"))

    accepted_ids = []
    problems = []

    for uploaded_file in chosen_files:
        submission_id, problem = save_one_upload(uploaded_file)

        if submission_id is None:
            problems.append(problem)
        else:
            accepted_ids.append(submission_id)

    for problem in problems:
        flash(problem)

    if len(accepted_ids) == 0:
        return redirect(url_for("home"))

    # one file goes straight to its own page, several are easier to watch on the status list
    if len(accepted_ids) == 1 and len(problems) == 0:
        return redirect(url_for("details", submission_id=accepted_ids[0]))

    flash(f"{len(accepted_ids)} file(s) added to the queue.")
    return redirect(url_for("status"))


@app.route("/status")
def status():
    # already newest first, and only the files this visitor sent in
    submissions = get_submissions_for_owner(get_owner_id())

    return render_template("status.html", submissions=submissions)


@app.route("/details/<submission_id>")
def details(submission_id):
    matching_submission = get_my_submission(submission_id)

    if matching_submission is None:
        logging.warning(f"Details requested for a submission that is unknown or belongs to someone else: {submission_id}")
        flash("That submission could not be found.")
        return redirect(url_for("status"))

    report_text = "No report is available for this submission."

    if matching_submission.get("alt_text_filename"):
        report_path = os.path.join(PROCESSED_FOLDER, matching_submission["alt_text_filename"])

        if os.path.exists(report_path):
            with open(report_path, "r", encoding="utf-8") as f:
                report_text = f.read()
        else:
            logging.error(f"Details page for submission {submission_id} expected a report at {report_path}, but it's missing.")

    return render_template("details.html", submission=matching_submission, report_text=report_text)


@app.route("/progress/<submission_id>")
def progress(submission_id):
    matching_submission = get_my_submission(submission_id)

    if matching_submission is None:
        return jsonify({"status": "unknown", "done": 0, "total": 0})

    counts = processing_progress.get(submission_id, {})

    return jsonify({
        "status": matching_submission["status"],
        "done": counts.get("done", 0),
        "total": counts.get("total", 0)
    })


@app.route("/retry/<submission_id>", methods=["POST"])
def retry_submission(submission_id):
    submission = get_my_submission(submission_id)

    if submission is None:
        flash("That submission could not be found, so it cannot be retried.")
        return redirect(url_for("status"))
    if submission["status"] != "error":
        flash("Only submissions with an error can be retried.")
        return redirect(url_for("details", submission_id=submission_id))

    filename = submission["saved_filename"]
    if not os.path.exists(os.path.join(UPLOAD_FOLDER, filename)):
        update_submission(submission_id, {
            "error_message": "The original upload is no longer available for retry.",
            "error_steps": ["Go back to the submit page and upload the presentation again."],
        })
        flash("We need the original upload to retry. Please submit the presentation again.")
        return redirect(url_for("details", submission_id=submission_id))

    retry_count = submission["retry_count"] + 1
    update_submission(submission_id, {"retry_count": retry_count, "status": "queued"})

    logging.info(f"Manual retry started for submission {submission_id} ({filename}), attempt {retry_count}.")
    start_processing(submission_id, filename)

    return redirect(url_for("details", submission_id=submission_id))


@app.route("/delete/<submission_id>", methods=["POST"])
def delete(submission_id):
    submission = get_my_submission(submission_id)

    if submission is None:
        flash("That submission could not be found.")
        return redirect(url_for("status"))

    # deleting a file while the worker is still using it would leave things half finished
    if submission["status"] == "queued" or submission["status"] == "processing":
        flash("That file is still being worked on. Wait until it finishes, then delete it.")
        return redirect(url_for("details", submission_id=submission_id))

    delete_submission_files(submission)
    delete_submission(submission_id)

    logging.info(f"Submission {submission_id} ({submission['original_filename']}) was deleted by the person who sent it.")
    flash(f"Deleted {submission['original_filename']}.")

    return redirect(url_for("status"))


@app.route("/download/<submission_id>")
def download(submission_id):
    matching_submission = get_my_submission(submission_id)

    if matching_submission is None:
        logging.warning(f"Download requested for a submission that is unknown or belongs to someone else: {submission_id}")
        flash("That submission could not be found.")
        return redirect(url_for("status"))

    if matching_submission["status"] != "done":
        logging.warning(f"Download requested for submission {submission_id} but its status is '{matching_submission['status']}'.")
        flash("That file isn't ready to download. It either failed processing or hasn't finished yet.")
        return redirect(url_for("details", submission_id=submission_id))

    saved_name = matching_submission["saved_filename"]
    processed_filename = os.path.splitext(saved_name)[0] + "_updated.pptx"

    # the result is always a .pptx, even when an old .ppt was uploaded
    original_stem = os.path.splitext(matching_submission["original_filename"])[0]
    download_name = "accessible_" + original_stem + ".pptx"

    processed_path = os.path.join(PROCESSED_FOLDER, processed_filename)
    if not os.path.exists(processed_path):
        logging.error(f"Download requested for submission {submission_id}, but {processed_path} is missing on disk.")
        flash("We couldn't find the processed file. Please try submitting it again.")
        return redirect(url_for("details", submission_id=submission_id))

    return send_from_directory(
        PROCESSED_FOLDER,
        processed_filename,
        as_attachment=True,
        download_name=download_name
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, port=port)
