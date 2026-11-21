import io
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as app_module
import bug_reports
from test_app import client  # noqa: F401 (client is a fixture)

SECRET_WORDS = "Prof. Lee's BIO 180 midterm"


@pytest.fixture
def sent(monkeypatch, client):
    sent = {"emails": [], "issues": []}
    monkeypatch.setattr(bug_reports, "SEND_TO", "hello@accessibilityenhancer.com")
    monkeypatch.setattr(bug_reports, "SEND_FROM", "no-reply@accessibilityenhancer.com")
    monkeypatch.setattr(bug_reports, "GITHUB_REPOSITORY", "t-nair/automated-accessibility-enhancer")
    monkeypatch.setattr(bug_reports, "github_token", lambda: "token")
    monkeypatch.setattr(bug_reports, "send_email", sent["emails"].append)
    monkeypatch.setattr(bug_reports, "open_issue", lambda issue: sent["issues"].append(issue) or "https://github.com/x/1")
    monkeypatch.setattr(app_module, "recent_bug_reports", {})

    with client.session_transaction() as browser_session:
        browser_session["email"] = "prof@uw.edu"

    return sent


def report(client, files=(), return_to="/status", subject=SECRET_WORDS):
    return client.post("/report-bug", content_type="multipart/form-data", data={
        "subject": subject,
        "description": f"Uploading {SECRET_WORDS}.pptx froze",
        "return_to": return_to,
        "attachments": [(io.BytesIO(data), name) for name, data in files],
    })


def test_a_report_is_emailed_in_full_and_the_public_issue_holds_none_of_it(client, sent):
    answer = report(client, files=[("shot.png", b"png"), ("midterm.pptx", b"deck")])

    assert answer.status_code == 302
    assert answer.headers["Location"].endswith("/status")

    email = sent["emails"][0]
    assert email["Reply-To"] == "prof@uw.edu"
    assert SECRET_WORDS in email["Subject"]
    assert [part.get_filename() for part in email.iter_attachments()] == ["shot.png", "midterm.pptx"]

    issue = sent["issues"][0]
    bug_id = issue["title"].split(":")[0]
    assert bug_id in email["Subject"]
    assert issue["labels"] == ["bug", "needs-triage"]
    for private in ("prof@uw.edu", "BIO 180", "midterm", "shot.png", "froze"):
        assert private not in issue["title"] + issue["body"]
    assert ".png" in issue["body"] and ".pptx" in issue["body"]


@pytest.mark.parametrize("files", [
    [(f"{n}.png", b"x") for n in range(4)],
    [("virus.exe", b"x")],
])
def test_too_many_or_the_wrong_files_send_nothing(client, sent, files):
    report(client, files=files)
    assert sent["emails"] == [] and sent["issues"] == []


def test_it_only_sends_people_back_to_this_site(client, sent):
    assert report(client, return_to="//evil.example").headers["Location"].endswith("/")
    assert "evil" not in report(client, return_to="https://evil.example").headers["Location"]


def test_a_failed_email_opens_no_issue_and_says_where_to_write(client, sent, monkeypatch):
    def broken(message):
        raise RuntimeError("SES is down")

    monkeypatch.setattr(bug_reports, "send_email", broken)
    report(client)

    assert sent["issues"] == []
    with client.session_transaction() as browser_session:
        assert "hello@accessibilityenhancer.com" in str(browser_session["_flashes"])


def test_reports_are_limited_per_hour(client, sent):
    for _ in range(app_module.BUG_REPORTS_PER_HOUR + 2):
        report(client)
    assert len(sent["emails"]) == app_module.BUG_REPORTS_PER_HOUR
