"""Bug reports from the "Report a bug" form.

Each report gets an id like BUG-20261120-7f3a and goes two places:

- an email to the team (BUG_REPORT_TO) with everything: who sent it, what they
  wrote and their attachments. The reporter is the Reply-To, so answering it
  reaches them.
- a public GitHub issue labelled bug, holding only facts the server knows itself
  (id, date, page, browser, attachment types and sizes). Nothing the person typed
  or uploaded goes there, because the repository is public and their decks can
  hold course material. Whoever reads the email adds a summary to the issue.

The email is the one that matters: if it fails the report fails. The issue is
best effort. With neither configured (a laptop) the report is only logged.
"""

import json
import logging
import mimetypes
import os
import secrets
import urllib.request
from datetime import datetime, timezone
from email.message import EmailMessage

SEND_TO = os.environ.get("BUG_REPORT_TO")
SEND_FROM = os.environ.get("BUG_REPORT_FROM")
GITHUB_REPOSITORY = os.environ.get("BUG_REPORT_GITHUB_REPO")

IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".gif", ".webp")
ALLOWED_EXTENSIONS = IMAGE_EXTENSIONS + (".pptx", ".ppt", ".pdf")
MAX_ATTACHMENTS = 3
MAX_SUBJECT_LENGTH = 150
MAX_DESCRIPTION_LENGTH = 5000

# SES refuses messages over 40 MB, and attachments grow by a third when encoded.
# ponytail: a deck bigger than this cannot be attached; upload to S3 and link it if that comes up.
MAX_ATTACHMENTS_MB = 25

_github_token = None


def new_bug_id():
    return f"BUG-{datetime.now(timezone.utc):%Y%m%d}-{secrets.token_hex(2)}"


# read once, on the first report, like the Cognito client secret in auth.py
def github_token():
    global _github_token

    if _github_token is None:
        if os.environ.get("GITHUB_ISSUES_TOKEN"):
            _github_token = os.environ["GITHUB_ISSUES_TOKEN"]
        elif os.environ.get("GITHUB_ISSUES_TOKEN_PARAMETER"):
            import boto3

            answer = boto3.client("ssm").get_parameter(
                Name=os.environ["GITHUB_ISSUES_TOKEN_PARAMETER"], WithDecryption=True
            )
            _github_token = answer["Parameter"]["Value"]

    return _github_token


# gives back a message saying what is wrong, or None if the report can be sent.
# attachments are (filename, bytes) pairs.
def problem_with(subject, description, attachments):
    if not subject:
        return "Add a one-line summary of the problem."

    if len(subject) > MAX_SUBJECT_LENGTH:
        return f"Keep the summary under {MAX_SUBJECT_LENGTH} characters and put the rest in the details."

    if len(description) > MAX_DESCRIPTION_LENGTH:
        return f"The details can be up to {MAX_DESCRIPTION_LENGTH} characters."

    if len(attachments) > MAX_ATTACHMENTS:
        return f"Attach at most {MAX_ATTACHMENTS} files."

    for filename, _ in attachments:
        if not filename.lower().endswith(ALLOWED_EXTENSIONS):
            return f"{filename}: only images (.png, .jpg, .gif, .webp), .pptx, .ppt and .pdf files can be attached."

    if sum(len(data) for _, data in attachments) > MAX_ATTACHMENTS_MB * 1024 * 1024:
        return f"The attachments add up to more than {MAX_ATTACHMENTS_MB} MB. Leave out the largest file and mention it in the details."

    return None


def describe_attachments(attachments):
    if not attachments:
        return "none"

    return ", ".join(
        f"{os.path.splitext(filename)[1].lower() or 'no extension'} ({len(data) / 1024 / 1024:.1f} MB)"
        for filename, data in attachments
    )


def issue_search_url(bug_id):
    return f"https://github.com/{GITHUB_REPOSITORY}/issues?q={bug_id}"


def build_email(bug_id, reporter_email, subject, description, attachments, facts):
    message = EmailMessage()
    message["Subject"] = f"[{bug_id}] {subject}"
    message["From"] = SEND_FROM
    message["To"] = SEND_TO

    if reporter_email:
        message["Reply-To"] = reporter_email

    lines = [
        f"Bug id: {bug_id}",
        f"From: {reporter_email or 'not signed in'}",
        *(f"{name}: {value}" for name, value in facts.items()),
    ]

    if GITHUB_REPOSITORY:
        lines.append(f"Public issue: {issue_search_url(bug_id)}")

    lines += ["", "Summary:", subject, "", "Details:", description or "(none given)"]
    message.set_content("\n".join(lines))

    for filename, data in attachments:
        content_type = mimetypes.guess_type(filename)[0] or "application/octet-stream"
        main_type, sub_type = content_type.split("/", 1)
        message.add_attachment(data, maintype=main_type, subtype=sub_type, filename=filename)

    return message


def build_issue(bug_id, facts):
    body = "\n".join(
        [
            "Filed automatically from the in-app bug report form.",
            "",
            *(f"- **{name}:** {value}" for name, value in facts.items()),
            "",
            "The reporter's description and files went to the team's inbox, not here, "
            f"because they can contain private course material. Search the inbox for `{bug_id}`, "
            "then replace this line with a summary that leaves out anything personal.",
        ]
    )
    return {"title": f"{bug_id}: bug report from the app", "body": body, "labels": ["bug", "needs-triage"]}


def send_email(message):
    import boto3

    boto3.client("sesv2").send_email(Content={"Raw": {"Data": message.as_bytes()}})


def open_issue(issue):
    request = urllib.request.Request(
        f"https://api.github.com/repos/{GITHUB_REPOSITORY}/issues",
        data=json.dumps(issue).encode(),
        headers={
            "Authorization": f"Bearer {github_token()}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        },
        method="POST",
    )

    with urllib.request.urlopen(request, timeout=10) as answer:
        return json.load(answer)["html_url"]


# sends the report and gives back its id. Raises if the email could not be sent,
# so the page can tell the person to write to the team directly instead.
def send(reporter_email, subject, description, attachments, facts):
    bug_id = new_bug_id()
    facts = {"Reported": f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC", **facts,
             "Attachments": describe_attachments(attachments)}

    if not SEND_TO:
        logging.info(f"Bug report {bug_id} (no BUG_REPORT_TO set, so it was only logged): {subject!r} {facts}")
        return bug_id

    send_email(build_email(bug_id, reporter_email, subject, description, attachments, facts))
    logging.info(f"Bug report {bug_id} emailed to the team.")

    # the token is read in here too, so a missing parameter only loses the issue, not the report
    try:
        if GITHUB_REPOSITORY and github_token():
            logging.info(f"Bug report {bug_id} opened as {open_issue(build_issue(bug_id, facts))}.")
    except Exception as e:
        logging.error(f"Bug report {bug_id} was emailed, but the GitHub issue could not be opened. Error: {e}")

    return bug_id
