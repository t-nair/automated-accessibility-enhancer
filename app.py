import os
import json
import time
import uuid
import queue
import logging
import secrets
import threading
from datetime import datetime, timedelta
from pathlib import Path
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, session
from werkzeug.utils import secure_filename
from z_reorder import process_one_file, severity_summary

# Where submissions and their files are kept. On AWS that is S3 and DynamoDB, and
# the pipeline runs in Lambda; on a laptop it is folders and a sqlite file, and the
# pipeline runs in a thread here. SUBMISSIONS_TABLE is only set on AWS.
if os.environ.get("SUBMISSIONS_TABLE"):
    import storage_aws as storage
else:
    import storage_local as storage

# on AWS the deck is processed by Lambda, which S3 starts on its own, so there is
# no queue and no worker thread in the web server
RUNS_PIPELINE_HERE = not os.environ.get("SUBMISSIONS_TABLE")

app = Flask(__name__)

# The project site serves this app at accessibilityenhancer.com/app and says so with
# X-Forwarded-Prefix, which makes url_for() build /app/... links so the stylesheet, nav
# and upload form resolve instead of being answered by the landing site.
PROXY_PREFIX = "/app"


# Only PROXY_PREFIX is honored. Render is reachable directly, so anyone can send this
# header; accepting whatever it says would let a caller choose the prefix every link on
# the page is built under. Accepting one known value makes that pointless, and leaves
# direct .onrender.com traffic serving itself at the root as before.
def _mount_under_known_prefix(wsgi_app):
    def wrapper(environ, start_response):
        if environ.get("HTTP_X_FORWARDED_PREFIX") == PROXY_PREFIX:
            environ["SCRIPT_NAME"] = PROXY_PREFIX
        return wsgi_app(environ, start_response)

    return wrapper


app.wsgi_app = _mount_under_known_prefix(app.wsgi_app)

ALLOWED_EXTENSIONS = (".pptx", ".ppt", ".pdf")
MAX_FILE_SIZE_MB = 50

# several files can be sent at once, so the whole request is allowed to be larger than
# one file. Anything past this is refused before it reaches the disk.
MAX_REQUEST_SIZE_MB = 250

SECONDS_IN_A_DAY = 24 * 60 * 60

# the languages offered on the upload form, as the code written into the deck and the
# name people see. Leaving it blank keeps whatever language each file already has.
LANGUAGES = {
    "en-US": "English (US)",
    "en-GB": "English (UK)",
    "es-ES": "Spanish",
    "fr-FR": "French",
    "de-DE": "German",
    "zh-CN": "Chinese (Simplified)",
    "ja-JP": "Japanese",
    "ko-KR": "Korean",
}

# FLASK_-prefixed environment variables set Flask settings, e.g. FLASK_SESSION_COOKIE_SECURE=true behind HTTPS
app.config.from_prefixed_env()

app.config["MAX_CONTENT_LENGTH"] = MAX_REQUEST_SIZE_MB * 1024 * 1024

storage.setup()

# how far along each file being processed right now is. Only used when the pipeline
# runs here, because it is not worth saving to disk once the file is finished. On AWS
# the Lambda writes its progress to the table instead.
processing_progress = {}

# files wait here to be processed. One worker takes them one at a time, so several
# uploads can never load the pipeline at once.
work_queue = queue.Queue()

app.secret_key = storage.get_secret_key()
app.permanent_session_lifetime = timedelta(days=30)


def is_pdf(filename):
    return filename.lower().endswith(".pdf")


ORIGINAL_GONE_MESSAGE = "The original upload is no longer available for retry."
ORIGINAL_GONE_STEPS = ["Go back to the submit page and upload the presentation again."]


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
    submission = storage.get_submission(submission_id)

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
    storage.update_submission(submission_id, {"status": new_status})


def process_one_submission(submission_id, saved_name):
    def update_progress(captions_done, total_to_caption):
        processing_progress[submission_id] = {"done": captions_done, "total": total_to_caption}

    logging.info(f"Submission {submission_id}: processing {saved_name} (queue depth {work_queue.qsize()}).")
    started_at = time.perf_counter()
    update_progress(0, 0)
    set_submission_status(submission_id, "processing")

    # read from the database so a retry sets the same language as the first attempt
    submission = storage.get_submission(submission_id)
    language = submission["language"]
    input_folder = storage.original_upload_folder(saved_name)

    if input_folder is None:
        was_successful, error_message, error_steps = False, ORIGINAL_GONE_MESSAGE, ORIGINAL_GONE_STEPS
    elif is_pdf(saved_name):
        was_successful, error_message, error_steps = process_one_pdf(
            saved_name, input_folder, submission["original_filename"], update_progress, language
        )
    else:
        was_successful, error_message, error_steps = process_one_file(
            saved_name, input_folder, storage.PROCESSED_FOLDER, update_progress, language
        )

    if not was_successful:
        logging.error(f"Submission {submission_id} ({saved_name}) failed processing: {error_message} Steps shown to user: {error_steps}")
    else:
        logging.info(f"Submission {submission_id} ({saved_name}) finished in {time.perf_counter() - started_at:.1f}s.")

    if was_successful:
        alt_text_filename = os.path.splitext(saved_name)[0] + "_alt_text"
    else:
        alt_text_filename = None

    storage.update_submission(submission_id, {
        "status": "done" if was_successful else "error",
        "error_message": error_message,
        "error_steps": error_steps,
        "alt_text_filename": alt_text_filename,
    })

    processing_progress.pop(submission_id, None)


# the PDF stages that count towards the progress bar, in the order they finish
PDF_STAGES = ("inspect", "image-remediation", "semantic-html", "tagged-pdf-export", "verify")


# the PDF version of process_one_file: gives back whether it worked, a message and steps to try
def process_one_pdf(saved_name, input_folder, original_filename, progress_callback, language=None):
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
    output_folder = Path(storage.pdf_output_folder(saved_name))

    try:
        manifest = remediate_pdf(
            Path(input_folder) / saved_name, output_folder, on_stage,
            fallback_title=os.path.splitext(original_filename)[0], language=language,
        )
    except Exception as e:
        logging.exception(f"PDF remediation of {saved_name} stopped. Error: {e}")
        return pdf_error_details(output_folder / (stem + ".remediation.json"), e)

    report_path = os.path.join(storage.PROCESSED_FOLDER, stem + "_alt_text")

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

    if isinstance(error, RuntimeError):
        # these messages come from the pipeline's own checks and say what stopped it
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


def worker_loop():
    last_cleanup = time.time()

    while True:
        submission_id, saved_name = work_queue.get()

        try:
            process_one_submission(submission_id, saved_name)
        except Exception as e:
            # a crash here must not kill the worker, or nothing else would ever process
            logging.exception(f"Unexpected problem while processing submission {submission_id}. Error: {e}")
            processing_progress.pop(submission_id, None)

            # without a message the page would show an error with no explanation
            storage.update_submission(submission_id, {
                "status": "error",
                "error_message": "Something unexpected went wrong while processing this file.",
                "error_steps": ["Please try again. If it keeps failing, contact the accessibility team."],
            })

        work_queue.task_done()

        # tidy up old files roughly once a day, so a server left running for months
        # does not slowly fill its disk
        if time.time() - last_cleanup > SECONDS_IN_A_DAY:
            last_cleanup = time.time()

            try:
                storage.delete_old_submissions()
            except Exception as e:
                logging.error(f"Could not tidy up old submissions. Error: {e}")


if RUNS_PIPELINE_HERE:
    # daemon means this thread does not keep the app running when it is shut down
    worker_thread = threading.Thread(target=worker_loop, daemon=True)
    worker_thread.start()


def start_processing(submission_id, saved_name):
    if RUNS_PIPELINE_HERE:
        work_queue.put((submission_id, saved_name))
        return

    # on AWS, putting the deck in the uploads bucket is what starts the pipeline,
    # so by the time we get here it is already on its way
    logging.info(f"Submission {submission_id} ({saved_name}) handed to the pipeline by S3.")


def restart_processing(submission_id, saved_name):
    if RUNS_PIPELINE_HERE:
        work_queue.put((submission_id, saved_name))
        return

    storage.restart_processing(saved_name)


# Flask refuses an oversized upload before our own checks get a chance to run,
# so this turns its error page into the same kind of message as everything else
@app.errorhandler(413)
def upload_was_too_large(error):
    logging.warning("Upload rejected: the request was larger than the limit.")
    flash(f"That upload was too large. Each file must be under {MAX_FILE_SIZE_MB} MB, and everything sent at once must be under {MAX_REQUEST_SIZE_MB} MB.")

    return redirect(url_for("home"))


# the queue only lives in memory, so anything waiting or half done when the app
# stopped is queued again, or marked as an error if its original is gone
def recover_unfinished_submissions():
    for row in storage.unfinished_submissions():
        if not storage.upload_exists(row["saved_filename"]):
            logging.error(f"Submission {row['id']} could not be resumed after a restart: {row['saved_filename']} is gone.")
            storage.update_submission(row["id"], {
                "status": "error",
                "error_message": ORIGINAL_GONE_MESSAGE,
                "error_steps": ORIGINAL_GONE_STEPS,
            })
            continue

        logging.info(f"Submission {row['id']} was unfinished when the app stopped, so it is queued again.")
        set_submission_status(row["id"], "queued")
        start_processing(row["id"], row["saved_filename"])


# every form carries this token, so another site can't make a visitor's browser
# upload, retry or delete files
def csrf_token():
    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_hex(32)

    return session["csrf_token"]


app.jinja_env.globals["csrf_token"] = csrf_token


@app.before_request
def check_csrf_token():
    if request.method != "POST":
        return None

    expected = session.get("csrf_token")
    sent = request.form.get("csrf_token", "")

    if expected and secrets.compare_digest(sent, expected):
        return None

    logging.warning(f"POST to {request.path} rejected: the form's security token was missing or wrong.")
    flash("That form had expired, so nothing was changed. Please try again.")
    return render_template("index.html", max_file_size_mb=MAX_FILE_SIZE_MB, languages=LANGUAGES), 400


@app.route("/")
def home():
    return render_template("index.html", max_file_size_mb=MAX_FILE_SIZE_MB, languages=LANGUAGES)


# gives back the new submission id, or None and a message saying what was wrong
def save_one_upload(uploaded_file, language=None):
    filename = uploaded_file.filename

    if not filename.lower().endswith(ALLOWED_EXTENSIONS):
        logging.warning(f"Upload rejected: {filename} is not a .pptx, .ppt or .pdf file.")
        return None, f"{filename}: only .pptx, .ppt and .pdf files are accepted."

    if is_pdf(filename) and not RUNS_PIPELINE_HERE:
        logging.warning(f"Upload rejected: {filename} is a PDF, and PDFs are not processed on this deployment.")
        return None, f"{filename}: PDFs are not supported on this server yet, only .pptx and .ppt files."

    if file_is_too_big(uploaded_file):
        logging.warning(f"Upload rejected: {filename} is over the {MAX_FILE_SIZE_MB}MB size limit.")
        return None, f"{filename}: too large (the limit is {MAX_FILE_SIZE_MB} MB)."

    submission_id = str(uuid.uuid4())[:8]
    saved_name = submission_id + "_" + secure_filename(filename)

    try:
        storage.save_upload(uploaded_file, saved_name)
    except Exception as e:
        logging.error(f"Could not save uploaded file {filename} as {saved_name}. Error: {e}")
        return None, f"{filename}: we could not save it. Please try again."

    storage.add_submission({
        "id": submission_id,
        "owner_id": get_owner_id(),
        "original_filename": filename,
        "saved_filename": saved_name,
        "status": "queued",
        "alt_text_filename": None,
        "error_message": None,
        "error_steps": None,
        "retry_count": 0,
        "submitted_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "language": language,
    })

    logging.info(f"Upload accepted: {filename} saved as {saved_name} (submission {submission_id}).")

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

    # blank means leave each file's language alone
    language = request.form.get("language", "") or None

    if language is not None and language not in LANGUAGES:
        logging.warning(f"Upload rejected: {language!r} is not one of the offered languages.")
        flash("Please choose a language from the list.")
        return redirect(url_for("home"))

    accepted_ids = []
    problems = []

    for uploaded_file in chosen_files:
        submission_id, problem = save_one_upload(uploaded_file, language)

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
    submissions = storage.get_submissions_for_owner(get_owner_id())

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
        saved_report = storage.read_report(matching_submission)

        if saved_report is not None:
            report_text = saved_report
        else:
            logging.error(f"Details page for submission {submission_id} expected a report, but it could not be read.")

    return render_template("details.html", submission=matching_submission, report_text=report_text)


@app.route("/progress/<submission_id>")
def progress(submission_id):
    matching_submission = get_my_submission(submission_id)

    if matching_submission is None:
        return jsonify({"status": "unknown", "done": 0, "total": 0})

    if RUNS_PIPELINE_HERE:
        counts = processing_progress.get(submission_id, {})
        done = counts.get("done", 0)
        total = counts.get("total", 0)
    else:
        done = int(matching_submission.get("captions_done") or 0)
        total = int(matching_submission.get("captions_total") or 0)

    return jsonify({
        "status": matching_submission["status"],
        "done": done,
        "total": total
    })


@app.route("/retry/<submission_id>", methods=["POST"])
def retry_submission(submission_id):
    submission = get_my_submission(submission_id)

    if submission is None:
        logging.warning(f"Retry requested for a submission that is unknown or belongs to someone else: {submission_id}")
        flash("That submission could not be found, so it cannot be retried.")
        return redirect(url_for("status"))
    if submission["status"] != "error":
        flash("Only submissions with an error can be retried.")
        return redirect(url_for("details", submission_id=submission_id))

    filename = submission["saved_filename"]
    if not storage.upload_exists(filename):
        logging.error(f"Retry of submission {submission_id} impossible: the original {filename} is gone.")
        storage.update_submission(submission_id, {
            "error_message": ORIGINAL_GONE_MESSAGE,
            "error_steps": ORIGINAL_GONE_STEPS,
        })
        flash("We need the original upload to retry. Please submit the presentation again.")
        return redirect(url_for("details", submission_id=submission_id))

    retry_count = submission["retry_count"] + 1
    storage.update_submission(submission_id, {"retry_count": retry_count, "status": "queued"})

    logging.info(f"Manual retry started for submission {submission_id} ({filename}), attempt {retry_count}.")
    restart_processing(submission_id, filename)

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

    storage.delete_submission_files(submission)
    storage.delete_submission(submission_id)

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

    if not storage.processed_file_is_ready(matching_submission):
        logging.error(f"Download requested for submission {submission_id}, but the processed file is missing.")
        flash("We couldn't find the processed file. Please try submitting it again.")
        return redirect(url_for("details", submission_id=submission_id))

    logging.info(f"Download of submission {submission_id}.")

    return storage.send_processed_file(matching_submission)


if __name__ == "__main__":
    from waitress import serve

    recover_unfinished_submissions()

    # one process with threads, never several processes: the queue and the progress
    # live in this process (see docs/deployment.md)
    serve(app, host=os.environ.get("HOST", "127.0.0.1"), port=int(os.environ.get("PORT", 5000)))
