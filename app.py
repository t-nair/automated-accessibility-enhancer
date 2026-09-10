import os
import json
import time
import uuid
import queue
import logging
import threading
from datetime import datetime, timedelta
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, session
from werkzeug.utils import secure_filename
from z_reorder import process_one_file

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

ALLOWED_EXTENSIONS = (".pptx", ".ppt")
MAX_FILE_SIZE_MB = 50

# several files can be sent at once, so the whole request is allowed to be larger than
# one file. Anything past this is refused before it reaches the disk.
MAX_REQUEST_SIZE_MB = 250

SECONDS_IN_A_DAY = 24 * 60 * 60

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

    update_progress(0, 0)
    set_submission_status(submission_id, "processing")

    was_successful, error_message, error_steps = process_one_file(
        saved_name, storage.UPLOAD_FOLDER, storage.PROCESSED_FOLDER, update_progress
    )

    if not was_successful:
        logging.error(f"Submission {submission_id} ({saved_name}) failed processing: {error_message}")

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


def worker_loop():
    last_cleanup = time.time()

    while True:
        submission_id, saved_name = work_queue.get()

        try:
            process_one_submission(submission_id, saved_name)
        except Exception as e:
            # a crash here must not kill the worker, or nothing else would ever process
            logging.error(f"Unexpected problem while processing submission {submission_id}. Error: {e}")
            set_submission_status(submission_id, "error")

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
    submissions = storage.get_submissions_for_owner(get_owner_id())

    return render_template("status.html", submissions=submissions)


# turns the saved report into a list of slides, so the page can show it a slide at a
# time instead of one long wall of text. Older reports were plain text, and those
# come back empty so the page falls back to showing them as they are.
def read_report(report_text):
    try:
        report = json.loads(report_text)
    except ValueError:
        return [], {}

    slides = []
    slides_by_number = {}
    described = 0
    already_had = 0
    no_description = 0
    words = 0
    busiest_slide = None
    busiest_words = 0

    def slide_entry(number):
        if number not in slides_by_number:
            slides_by_number[number] = {"number": number, "shapes": [], "problems": [], "images": 0, "added": 0, "words": 0}
            slides.append(slides_by_number[number])

        return slides_by_number[number]

    for shape in report.get("shapes", []):
        number = shape["slide"]
        slide_entry(number)

        slides_by_number[number]["shapes"].append(shape)

        if shape["kind"] == "image":
            slides_by_number[number]["images"] += 1

            if shape["we_added_it"]:
                slides_by_number[number]["added"] += 1
                described += 1
            else:
                already_had += 1
        elif shape["kind"] == "other":
            no_description += 1

        if shape["kind"] == "text" and shape["description"]:
            slides_by_number[number]["words"] += len(shape["description"].split())

    # problems are things a person has to fix, so they are listed against their slide
    # and counted separately from the descriptions we wrote
    problems = report.get("problems", [])

    for problem in problems:
        slide_entry(problem["slide"])["problems"].append(problem)

    slides.sort(key=lambda slide: slide["number"])

    # the busiest slide is worth pointing at, because a wall of text is hard to
    # follow whether you are listening to it or reading it
    for slide in slides:
        words += slide["words"]

        if slide["words"] > busiest_words:
            busiest_words = slide["words"]
            busiest_slide = slide["number"]

    words_per_slide = 0

    if slides:
        words_per_slide = round(words / len(slides))

    totals = {
        "slides": len(slides),
        "shapes": len(report.get("shapes", [])),
        "described": described,
        "already_had": already_had,
        "no_description": no_description,
        "problems": len(problems),
        "reading_level": report.get("reading_level"),
        "words": words,
        "words_per_slide": words_per_slide,
        "busiest_slide": busiest_slide,
        "busiest_words": busiest_words,
    }

    return slides, totals


@app.route("/details/<submission_id>")
def details(submission_id):
    matching_submission = get_my_submission(submission_id)

    if matching_submission is None:
        logging.warning(f"Details requested for a submission that is unknown or belongs to someone else: {submission_id}")
        flash("That submission could not be found.")
        return redirect(url_for("status"))

    report_text = "No report is available for this submission."
    slides = []
    totals = {}

    if matching_submission.get("alt_text_filename"):
        saved_report = storage.read_report(matching_submission)

        if saved_report is not None:
            report_text = saved_report
            slides, totals = read_report(report_text)
        else:
            logging.error(f"Details page for submission {submission_id} expected a report, but it could not be read.")

    return render_template(
        "details.html",
        submission=matching_submission,
        report_text=report_text,
        slides=slides,
        totals=totals
    )


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
        flash("That submission could not be found, so it cannot be retried.")
        return redirect(url_for("status"))
    if submission["status"] != "error":
        flash("Only submissions with an error can be retried.")
        return redirect(url_for("details", submission_id=submission_id))

    filename = submission["saved_filename"]
    if not storage.upload_exists(filename):
        storage.update_submission(submission_id, {
            "error_message": "The original upload is no longer available for retry.",
            "error_steps": ["Go back to the submit page and upload the presentation again."],
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

    return storage.send_processed_file(matching_submission)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))

    # debug mode opens a Python console in the browser whenever something breaks,
    # so it has to be off any time other people can reach the site
    debug_mode = os.environ.get("DEBUG", "true").lower() == "true"

    if not debug_mode:
        logging.info("Starting with debug mode off.")

    app.run(debug=debug_mode, port=port)
