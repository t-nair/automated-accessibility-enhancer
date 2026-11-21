import io
import os
import time
import uuid
import queue
import logging
import secrets
import threading
from datetime import datetime, timedelta
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, session, send_file
from werkzeug.utils import secure_filename
from z_reorder import process_one_file
from pdf_runner import process_one_pdf
import auth

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


# Only PROXY_PREFIX is honored. The ECS address is reachable directly, so anyone can send this
# header; accepting whatever it says would let a caller choose the prefix every link on
# the page is built under. Accepting one known value makes that pointless, and leaves
# direct traffic to the ECS address serving itself at the root as before.
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

# Lax still sends the cookie when Cognito sends someone back after sign-in, which
# is a top-level GET, and keeps it off requests other sites start. Secure whenever
# the public address is HTTPS. Either can still be set with FLASK_ variables.
app.config.setdefault("SESSION_COOKIE_SAMESITE", "Lax")

if "FLASK_SESSION_COOKIE_SECURE" not in os.environ:
    app.config["SESSION_COOKIE_SECURE"] = auth.PUBLIC_URL.startswith("https://")

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
# is kept separate from everyone else's. Signing in swaps it for the account's
# Cognito sub, so the list follows the person instead of the browser.
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
            saved_name, input_folder, storage.pdf_output_folder(saved_name), storage.PROCESSED_FOLDER,
            submission["original_filename"], update_progress, language
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


@app.context_processor
def sign_in_details():
    return {
        "sign_in_available": auth.ENABLED,
        "signed_in_email": session.get("email"),
    }


# the pages for signing in, creating an account and resetting a password
ACCOUNT_ENDPOINTS = {"sign_in", "sign_up", "confirm_email", "resend_code", "forgot_password",
                     "reset_password", "google_sign_in", "auth_callback"}

# pages anyone can reach: those, styles, and the health check
OPEN_ENDPOINTS = ACCOUNT_ENDPOINTS | {"static", "health"}


# With Cognito on, every page needs an account. Runs before the CSRF check, so
# someone signed out is sent to sign in instead of being told a form expired.
# Without Cognito (a laptop, a server of your own) the site stays open.
@app.before_request
def require_sign_in():
    if not auth.ENABLED or session.get("email") or request.endpoint in OPEN_ENDPOINTS:
        return None

    return redirect(url_for("sign_in"))


@app.route("/health")
def health():
    return "ok"


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

    if auth.ENABLED and not session.get("email"):
        return redirect(url_for("sign_in"))

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
    report = None

    if matching_submission.get("alt_text_filename"):
        saved_report = storage.read_report(matching_submission)

        if saved_report is not None:
            report_text = saved_report
        else:
            logging.error(f"Details page for submission {submission_id} expected a report, but it could not be read.")

        # the same findings arranged by slide. A PDF, or anything processed before
        # this existed, has none, and the page shows the text report instead.
        report = storage.read_report_data(matching_submission)

    return render_template("details.html", submission=matching_submission,
                           report_text=report_text, report=report)


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


# --- sign-in (required when Cognito is on; see auth.py) ---
#
# Email and password are on our own pages (templates/account.html), checked by
# Cognito's API. Google goes to Google and comes back to /auth/callback.

def account_page(form, status=200):
    return render_template("account.html", form=form, email=session.get("account_email", ""),
                           google_available=auth.GOOGLE_ENABLED), status


def form_email():
    email = request.form.get("email", "").strip()
    session["account_email"] = email
    return email


# Runs one of auth's Cognito calls. SignInError comes back for the page to show;
# anything else (Cognito down, Parameter Store unreadable) is logged and becomes a
# SignInError too, so no page has to know the difference.
def ask_cognito(action, *args):
    try:
        return action(*args)
    except auth.SignInError:
        raise
    except Exception as e:
        logging.error(f"Sign-in failed while talking to Cognito. Error: {e}")
        raise auth.SignInError("The sign-in service did not answer. Please try again in a moment.")


def finish_sign_in(user, with_google=False):
    # Only an anonymous browser's files move. If someone was already signed in here
    # (a shared computer, nobody signed out), their files stay in their account.
    browser_owner = None if session.get("email") else session.get("owner_id")

    # a fresh session, so nothing from before sign-in carries over
    session.clear()
    session["owner_id"] = user["sub"]
    session["email"] = user["email"]
    session["with_google"] = with_google
    session.permanent = True

    if browser_owner and browser_owner != user["sub"]:
        storage.move_submissions(browser_owner, user["sub"])

    logging.info(f"Signed in: account {user['sub']}.")
    flash(f"Signed in as {user['email']}.")

    return redirect(url_for("home"))


# the account pages make no sense once signed in, or with no accounts at all
@app.before_request
def skip_account_pages():
    if request.endpoint in ACCOUNT_ENDPOINTS and (not auth.ENABLED or session.get("email")):
        return redirect(url_for("home"))

    return None


@app.route("/sign-in", methods=["GET", "POST"])
def sign_in():
    if request.method == "GET":
        return account_page("sign_in")

    email = form_email()

    try:
        user = ask_cognito(auth.user_for_password, email, request.form.get("password", ""))
    except auth.SignInError as e:
        flash(e.message)

        if e.code == "UserNotConfirmedException":
            try:
                ask_cognito(auth.resend_code, email)
            except auth.SignInError as resend_problem:
                flash(resend_problem.message)

            return redirect(url_for("confirm_email"))

        if e.code == "PasswordResetRequiredException":
            return redirect(url_for("forgot_password"))

        return account_page("sign_in", 400)

    return finish_sign_in(user)


@app.route("/sign-up", methods=["GET", "POST"])
def sign_up():
    if request.method == "GET":
        return account_page("sign_up")

    email = form_email()

    try:
        sign_up_session = ask_cognito(auth.sign_up, email, request.form.get("password", ""))
    except auth.SignInError as e:
        flash(e.message)
        return account_page("sign_up", 400)

    # lets confirming the email sign them in too, without asking for the password
    # again. Useless to anyone else, and it expires in a few minutes.
    if sign_up_session:
        session["sign_up_session"] = {"email": email, "session": sign_up_session}

    flash(f"We sent a code to {email}. Enter it below to confirm your email.")
    return redirect(url_for("confirm_email"))


@app.route("/confirm-email", methods=["GET", "POST"])
def confirm_email():
    if request.method == "GET":
        return account_page("confirm_email")

    email = form_email()
    saved = session.get("sign_up_session") or {}
    sign_up_session = saved.get("session") if saved.get("email") == email else None

    try:
        sign_in_session = ask_cognito(auth.confirm_sign_up, email, request.form.get("code", "").strip(), sign_up_session)
    except auth.SignInError as e:
        flash(e.message)
        return account_page("confirm_email", 400)

    session.pop("sign_up_session", None)

    if sign_in_session:
        try:
            return finish_sign_in(ask_cognito(auth.user_for_sign_up_session, email, sign_in_session))
        except auth.SignInError:
            pass  # confirmed all the same; they sign in with their password

    # the password was never kept, so ask for it once more
    flash("Your email is confirmed. Sign in to start using the site.")
    return redirect(url_for("sign_in"))


@app.route("/confirm-email/resend", methods=["POST"])
def resend_code():
    email = form_email()

    try:
        ask_cognito(auth.resend_code, email)
        flash(f"We sent a new code to {email}.")
    except auth.SignInError as e:
        flash(e.message)

    return redirect(url_for("confirm_email"))


@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "GET":
        return account_page("forgot_password")

    email = form_email()

    try:
        ask_cognito(auth.forgot_password, email)
    except auth.SignInError as e:
        flash(e.message)
        return account_page("forgot_password", 400)

    flash(f"If there is an account for {email}, we sent it a code. Enter it below with a new password.")
    return redirect(url_for("reset_password"))


@app.route("/reset-password", methods=["GET", "POST"])
def reset_password():
    if request.method == "GET":
        return account_page("reset_password")

    email = form_email()

    try:
        ask_cognito(auth.reset_password, email, request.form.get("code", "").strip(), request.form.get("password", ""))
    except auth.SignInError as e:
        flash(e.message)
        return account_page("reset_password", 400)

    flash("Your password is changed. Sign in with the new one.")
    return redirect(url_for("sign_in"))


def sign_in_redirect_uri():
    if auth.PUBLIC_URL:
        return auth.PUBLIC_URL + "/auth/callback"

    return url_for("auth_callback", _external=True)


@app.route("/auth/google")
def google_sign_in():
    if not auth.GOOGLE_ENABLED:
        return redirect(url_for("sign_in"))

    state, verifier = auth.new_state_and_verifier()
    session["sign_in_state"] = state
    session["sign_in_verifier"] = verifier

    return redirect(auth.google_sign_in_url(sign_in_redirect_uri(), state, verifier))


@app.route("/auth/callback")
def auth_callback():
    expected_state = session.pop("sign_in_state", None)
    verifier = session.pop("sign_in_verifier", None)

    # someone who cancels at Google comes back with an error instead of a code
    if request.args.get("error"):
        logging.info(f"Sign-in did not finish: {request.args.get('error')}.")
        flash("You were not signed in. Please sign in to use the site.")
        return redirect(url_for("sign_in"))

    sent_state = request.args.get("state", "")

    if not expected_state or not verifier or not secrets.compare_digest(sent_state, expected_state):
        logging.warning("Sign-in callback rejected: the state did not match this browser's.")
        flash("That sign-in link had expired. Please sign in again.")
        return redirect(url_for("sign_in"))

    try:
        user = auth.user_for_code(request.args.get("code", ""), verifier, sign_in_redirect_uri())
    except Exception as e:
        logging.error(f"Sign-in failed while talking to Cognito. Error: {e}")
        flash("Sign-in did not work. Please try again.")
        return redirect(url_for("sign_in"))

    return finish_sign_in(user, with_google=True)


@app.route("/auth/sign-out", methods=["POST"])
def sign_out():
    with_google = session.get("with_google")
    session.clear()

    # Signing in with Google leaves a session at Cognito's domain too, which would
    # sign the next person at this browser straight back in. End it, and Cognito
    # sends the person back here. Password sign-ins leave nothing there.
    if auth.ENABLED and with_google:
        home_url = (auth.PUBLIC_URL + "/") if auth.PUBLIC_URL else url_for("home", _external=True)
        return redirect(auth.sign_out_url(home_url))

    return redirect(url_for("home"))


@app.route("/report/<submission_id>")
def download_report(submission_id):
    matching_submission = get_my_submission(submission_id)

    if matching_submission is None:
        logging.warning(f"Report requested for a submission that is unknown or belongs to someone else: {submission_id}")
        flash("That submission could not be found.")
        return redirect(url_for("status"))

    report_text = storage.read_report(matching_submission)

    if report_text is None:
        flash("There is no report to download for that submission.")
        return redirect(url_for("details", submission_id=submission_id))

    # the report is written without an extension, so give the download a .txt one
    # for the sake of whoever opens it
    # control characters would make the header invalid, so they are dropped
    stem = "".join(c for c in os.path.splitext(matching_submission["original_filename"])[0] if c.isprintable())
    filename = (stem or "presentation") + "_accessibility_report.txt"

    # send_file quotes and encodes the name itself, so a quote or an accented
    # letter in the uploaded name cannot break the header
    return send_file(io.BytesIO(report_text.encode("utf-8")), mimetype="text/plain",
                     as_attachment=True, download_name=filename)


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
