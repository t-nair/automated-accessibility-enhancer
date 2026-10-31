import os
import sys
import threading
import urllib.request

import pytest
from werkzeug.serving import make_server

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as app_module

# CI must run these, so only a local run without Playwright may skip them
if not os.environ.get("CI"):
    pytest.importorskip("playwright.sync_api")

from playwright.sync_api import sync_playwright

# the version the audit in docs/accessibility-audit.md was run with
AXE_URL = "https://cdnjs.cloudflare.com/ajax/libs/axe-core/4.10.2/axe.min.js"
OWNER_ID = "owner-being-audited"


def add_submission(submission_id, status, **extras):
    submission = {
        "id": submission_id,
        "owner_id": OWNER_ID,
        "original_filename": "Lecture.pptx",
        "saved_filename": submission_id + "_Lecture.pptx",
        "status": status,
        "alt_text_filename": None,
        "error_message": None,
        "error_steps": None,
        "retry_count": 0,
        "submitted_at": "2026-08-20 10:00",
    }
    submission.update(extras)
    app_module.add_submission(submission)


# the app on a real port, with one submission in each status. Nothing is ever processed,
# so the description model is never loaded.
@pytest.fixture(scope="module")
def base_url(tmp_path_factory):
    folder = tmp_path_factory.mktemp("axe")
    patch = pytest.MonkeyPatch()
    patch.setattr(app_module, "UPLOAD_FOLDER", str(folder))
    patch.setattr(app_module, "PROCESSED_FOLDER", str(folder))
    patch.setattr(app_module, "DATABASE_FILE", str(folder / "submissions.db"))
    patch.setitem(app_module.app.config, "UPLOAD_FOLDER", str(folder))
    patch.setattr(app_module, "start_processing", lambda submission_id, saved_name: None)
    app_module.setup_database()

    (folder / "done_report.txt").write_text("Slide 1: added a description to 1 picture.\n", encoding="utf-8")
    add_submission("queued01", "queued")
    add_submission("process1", "processing")
    add_submission("done0001", "done", alt_text_filename="done_report.txt")
    add_submission("error001", "error", error_message="The file could not be opened.",
                   error_steps=["Open it in PowerPoint", "Save it as .pptx", "Upload it again"])

    server = make_server("127.0.0.1", 0, app_module.app, threaded=True)
    threading.Thread(target=server.serve_forever, daemon=True).start()

    yield f"http://127.0.0.1:{server.server_port}"

    server.shutdown()
    patch.undo()


@pytest.fixture(scope="module")
def browser():
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch()
        yield browser
        browser.close()


@pytest.fixture(scope="module")
def axe_source():
    with urllib.request.urlopen(AXE_URL, timeout=30) as response:
        return response.read().decode("utf-8")


# form.submit() skips the page's own checks, the same as a browser without JavaScript
def submit_with_no_file(page):
    page.evaluate("document.getElementById('upload-form').submit()")


def submit_with_stale_token(page):
    page.evaluate("""() => {
        const form = document.getElementById('upload-form');
        form.querySelector('[name=csrf_token]').value = 'stale';
        form.submit();
    }""")


# every page and state in docs/accessibility-audit.md. The delete confirmation is the
# browser's own dialog, so there is no page content for axe to check.
STATES = [
    ("upload empty", "/", None, False),
    ("upload validation error", "/", submit_with_no_file, True),
    ("upload flash message", "/", submit_with_stale_token, True),
    ("status no submissions", "/status", None, False),
    ("status every status", "/status", None, True),
    ("status flash message", "/details/nosuchid", None, True),
    ("details queued", "/details/queued01", None, True),
    ("details processing", "/details/process1", None, True),
    ("details done with report", "/details/done0001", None, True),
    ("details failed", "/details/error001", None, True),
]


@pytest.mark.parametrize("path, action, as_owner", [state[1:] for state in STATES], ids=[state[0] for state in STATES])
def test_page_has_no_axe_violations(base_url, browser, axe_source, path, action, as_owner):
    # a zero-width viewport hides colours from axe, so always set a real one
    context = browser.new_context(viewport={"width": 1280, "height": 900})

    if as_owner:
        cookie = app_module.app.session_interface.get_signing_serializer(app_module.app).dumps({"owner_id": OWNER_ID})
        context.add_cookies([{"name": app_module.app.config["SESSION_COOKIE_NAME"], "value": cookie, "url": base_url}])

    page = context.new_page()
    page.goto(base_url + path)

    if action:
        with page.expect_navigation():
            action(page)

    page.evaluate(axe_source)
    violations = page.evaluate("async () => (await axe.run()).violations")
    context.close()

    assert violations == [], "\n".join(
        f"{v['id']} ({v['impact']}): {v['help']} at {[node['target'] for node in v['nodes']]}" for v in violations
    )
