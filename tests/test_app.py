import io
import os
import sys

import pytest

# make sure the tests can import the project files from the folder above
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as app_module

FIXTURE_FOLDER = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Error Test PPTX")

MY_OWNER_ID = "owner-doing-the-testing"
SOMEONE_ELSE = "a-different-visitor"


@pytest.fixture
def client(tmp_path, monkeypatch):
    uploads = tmp_path / "uploads"
    processed = tmp_path / "processed"
    data = tmp_path / "data"
    uploads.mkdir()
    processed.mkdir()
    data.mkdir()

    monkeypatch.setattr(app_module, "UPLOAD_FOLDER", str(uploads))
    monkeypatch.setattr(app_module, "PROCESSED_FOLDER", str(processed))
    monkeypatch.setattr(app_module, "DATABASE_FILE", str(data / "submissions.db"))
    app_module.app.config["UPLOAD_FOLDER"] = str(uploads)

    app_module.setup_database()

    # the real pipeline is slow and needs a GPU, so tests only check that files get queued
    queued = []
    monkeypatch.setattr(app_module, "start_processing", lambda submission_id, saved_name: queued.append(submission_id))

    app_module.app.config["TESTING"] = True
    test_client = app_module.app.test_client()
    test_client.queued = queued

    # pretend this browser has already been given an id
    with test_client.session_transaction() as browser_session:
        browser_session["owner_id"] = MY_OWNER_ID

    return test_client


def make_upload(filename):
    path = os.path.join(FIXTURE_FOLDER, filename)

    with open(path, "rb") as f:
        contents = f.read()

    return (io.BytesIO(contents), filename)


def add_submission(submission_id="abc12345", owner_id=MY_OWNER_ID, status="done", **extras):
    submission = {
        "id": submission_id,
        "owner_id": owner_id,
        "original_filename": "Lecture.pptx",
        "saved_filename": submission_id + "_Lecture.pptx",
        "status": status,
        "alt_text_filename": submission_id + "_Lecture_alt_text",
        "error_message": None,
        "error_steps": None,
        "retry_count": 0,
        "submitted_at": "2026-08-20 10:00",
    }
    submission.update(extras)
    app_module.add_submission(submission)

    return submission


def copy_into(source_name, target_path):
    with open(os.path.join(FIXTURE_FOLDER, source_name), "rb") as f:
        contents = f.read()

    with open(target_path, "wb") as f:
        f.write(contents)


def test_a_submission_can_be_read_back(client):
    add_submission("read0001")

    assert app_module.get_submission("read0001")["original_filename"] == "Lecture.pptx"


def test_an_unknown_submission_is_none(client):
    assert app_module.get_submission("nosuchid") is None


def test_error_steps_survive_being_stored(client):
    add_submission("step0001", error_steps=["First thing", "Second thing"])

    assert app_module.get_submission("step0001")["error_steps"] == ["First thing", "Second thing"]


def test_updating_a_submission_changes_it(client):
    add_submission("edit0001", status="queued")
    app_module.update_submission("edit0001", {"status": "done"})

    assert app_module.get_submission("edit0001")["status"] == "done"


def test_only_my_submissions_are_listed(client):
    add_submission("mine0001", owner_id=MY_OWNER_ID)
    add_submission("their001", owner_id=SOMEONE_ELSE)

    mine = app_module.get_submissions_for_owner(MY_OWNER_ID)
    ids = []

    for submission in mine:
        ids.append(submission["id"])

    assert ids == ["mine0001"]


def test_home_page_loads(client):
    response = client.get("/")

    assert response.status_code == 200
    assert b"Drag and drop" in response.data


def test_home_page_allows_more_than_one_file(client):
    response = client.get("/")

    assert b"multiple" in response.data


def test_status_page_loads_when_there_are_no_submissions(client):
    response = client.get("/status")

    assert response.status_code == 200
    assert b"not submitted any files" in response.data


def test_status_page_shows_my_files(client):
    add_submission("mine0002")

    response = client.get("/status")

    assert b"Lecture.pptx" in response.data


def test_status_page_hides_other_peoples_files(client):
    add_submission("their002", owner_id=SOMEONE_ELSE, original_filename="Secret.pptx")

    response = client.get("/status")

    assert b"Secret.pptx" not in response.data
    assert b"not submitted any files" in response.data


def test_uploading_nothing_is_rejected(client):
    response = client.post("/upload", data={}, follow_redirects=True)

    assert b"choose at least one file" in response.data


def test_uploading_a_wrong_file_type_is_rejected(client):
    data = {"presentation": (io.BytesIO(b"just some notes"), "notes.txt")}
    response = client.post("/upload", data=data, follow_redirects=True)

    assert b"only .pptx and .ppt files are accepted" in response.data
    assert len(client.queued) == 0


def test_uploading_a_pptx_creates_a_submission(client):
    data = {"presentation": make_upload("control_no_issues.pptx")}
    response = client.post("/upload", data=data, follow_redirects=True)

    assert response.status_code == 200

    submissions = app_module.get_submissions_for_owner(MY_OWNER_ID)

    assert len(submissions) == 1
    assert submissions[0]["original_filename"] == "control_no_issues.pptx"
    assert submissions[0]["status"] == "queued"


def test_an_upload_belongs_to_the_person_who_sent_it(client):
    data = {"presentation": make_upload("control_no_issues.pptx")}
    client.post("/upload", data=data)

    submissions = app_module.get_submissions_for_owner(MY_OWNER_ID)

    assert submissions[0]["owner_id"] == MY_OWNER_ID


def test_an_uploaded_file_is_queued_for_processing(client):
    data = {"presentation": make_upload("control_no_issues.pptx")}
    client.post("/upload", data=data)

    assert len(client.queued) == 1


def test_uploading_several_files_at_once(client):
    data = {
        "presentation": [
            make_upload("control_no_issues.pptx"),
            make_upload("issue_no_title_shape.pptx"),
            make_upload("issue_missing_alt_text.pptx"),
        ]
    }
    response = client.post("/upload", data=data, follow_redirects=True)

    assert b"3 file(s) added to the queue" in response.data
    assert len(app_module.get_submissions_for_owner(MY_OWNER_ID)) == 3
    assert len(client.queued) == 3


def test_a_bad_file_does_not_stop_the_good_ones(client):
    data = {
        "presentation": [
            (io.BytesIO(b"not a presentation"), "notes.txt"),
            make_upload("control_no_issues.pptx"),
        ]
    }
    response = client.post("/upload", data=data, follow_redirects=True)

    assert b"notes.txt" in response.data
    assert len(app_module.get_submissions_for_owner(MY_OWNER_ID)) == 1
    assert len(client.queued) == 1


def test_a_file_that_is_too_big_is_rejected(client, monkeypatch):
    monkeypatch.setattr(app_module, "MAX_FILE_SIZE_MB", 0)

    data = {"presentation": make_upload("control_no_issues.pptx")}
    response = client.post("/upload", data=data, follow_redirects=True)

    assert b"too large" in response.data
    assert len(client.queued) == 0


def test_details_page_shows_a_submission(client):
    add_submission("abc12345")

    response = client.get("/details/abc12345")

    assert response.status_code == 200
    assert b"Lecture.pptx" in response.data


def test_details_page_for_an_unknown_id_goes_back_to_status(client):
    response = client.get("/details/nosuchid", follow_redirects=True)

    assert b"could not be found" in response.data


def test_details_page_refuses_someone_elses_submission(client):
    add_submission("their003", owner_id=SOMEONE_ELSE, original_filename="Secret.pptx")

    response = client.get("/details/their003", follow_redirects=True)

    assert b"Secret.pptx" not in response.data
    assert b"could not be found" in response.data


def test_details_page_shows_the_error_and_the_steps(client):
    add_submission(
        "err00001",
        status="error",
        alt_text_filename=None,
        error_message="This file is empty (0 bytes), so there's nothing to open.",
        error_steps=["Check the file on your computer.", "Upload it again."],
    )

    response = client.get("/details/err00001")

    assert b"nothing to open" in response.data
    assert b"Check the file on your computer." in response.data


def test_details_page_shows_a_progress_bar_while_queued(client):
    add_submission("wait0001", status="queued")

    response = client.get("/details/wait0001")

    assert b"progress-bar" in response.data


def test_progress_reports_the_status(client):
    add_submission("prog0001", status="queued")

    body = client.get("/progress/prog0001").get_json()

    assert body["status"] == "queued"
    assert body["done"] == 0


def test_progress_for_an_unknown_id_says_unknown(client):
    assert client.get("/progress/nosuchid").get_json()["status"] == "unknown"


def test_progress_for_someone_elses_submission_says_unknown(client):
    add_submission("their004", owner_id=SOMEONE_ELSE, status="queued")

    assert client.get("/progress/their004").get_json()["status"] == "unknown"


def test_downloading_a_file_that_is_not_ready_is_refused(client):
    add_submission("busy0001", status="queued")

    response = client.get("/download/busy0001", follow_redirects=True)

    # the page escapes apostrophes, so only match the plain part of the message
    assert b"ready to download" in response.data


def test_downloading_an_unknown_id_goes_back_to_status(client):
    response = client.get("/download/nosuchid", follow_redirects=True)

    assert b"could not be found" in response.data


def test_downloading_someone_elses_file_is_refused(client):
    add_submission("their005", owner_id=SOMEONE_ELSE)
    copy_into("control_no_issues.pptx", os.path.join(app_module.PROCESSED_FOLDER, "their005_Lecture_updated.pptx"))

    response = client.get("/download/their005", follow_redirects=True)

    assert b"could not be found" in response.data


def test_downloading_when_the_processed_file_is_missing(client):
    add_submission("gone0001")

    response = client.get("/download/gone0001", follow_redirects=True)

    assert b"find the processed file" in response.data


def test_downloading_a_finished_file_sends_it(client):
    add_submission("good0001")
    copy_into("control_no_issues.pptx", os.path.join(app_module.PROCESSED_FOLDER, "good0001_Lecture_updated.pptx"))

    response = client.get("/download/good0001")

    assert response.status_code == 200
    assert "accessible_Lecture.pptx" in response.headers["Content-Disposition"]


def test_retrying_a_finished_file_is_refused(client):
    add_submission("fine0001")

    response = client.post("/retry/fine0001", follow_redirects=True)

    assert b"Only submissions with an error" in response.data


def test_retrying_someone_elses_file_is_refused(client):
    add_submission("their006", owner_id=SOMEONE_ELSE, status="error")

    response = client.post("/retry/their006", follow_redirects=True)

    assert b"could not be found" in response.data
    assert len(client.queued) == 0


def test_retrying_without_the_original_upload_asks_for_it_again(client):
    add_submission("lost0001", status="error")

    response = client.post("/retry/lost0001", follow_redirects=True)

    assert b"need the original upload" in response.data


def test_retrying_queues_the_file_again(client):
    add_submission("redo0001", status="error")

    # the retry route needs the original upload to still be there
    copy_into("control_no_issues.pptx", os.path.join(app_module.UPLOAD_FOLDER, "redo0001_Lecture.pptx"))

    client.post("/retry/redo0001", follow_redirects=True)

    submission = app_module.get_submission("redo0001")

    assert submission["status"] == "queued"
    assert submission["retry_count"] == 1
    assert len(client.queued) == 1


def test_deleting_removes_the_submission(client):
    add_submission("del00001")

    response = client.post("/delete/del00001", follow_redirects=True)

    assert b"Deleted Lecture.pptx" in response.data
    assert app_module.get_submission("del00001") is None


def test_deleting_also_removes_the_files(client):
    add_submission("del00002")

    upload_path = os.path.join(app_module.UPLOAD_FOLDER, "del00002_Lecture.pptx")
    updated_path = os.path.join(app_module.PROCESSED_FOLDER, "del00002_Lecture_updated.pptx")
    report_path = os.path.join(app_module.PROCESSED_FOLDER, "del00002_Lecture_alt_text")

    copy_into("control_no_issues.pptx", upload_path)
    copy_into("control_no_issues.pptx", updated_path)

    with open(report_path, "w") as f:
        f.write("a report")

    client.post("/delete/del00002", follow_redirects=True)

    assert not os.path.exists(upload_path)
    assert not os.path.exists(updated_path)
    assert not os.path.exists(report_path)


def test_deleting_someone_elses_submission_is_refused(client):
    add_submission("their007", owner_id=SOMEONE_ELSE)

    response = client.post("/delete/their007", follow_redirects=True)

    assert b"could not be found" in response.data
    assert app_module.get_submission("their007") is not None


def test_deleting_an_unknown_submission_is_refused(client):
    response = client.post("/delete/nosuchid", follow_redirects=True)

    assert b"could not be found" in response.data


def test_a_file_being_worked_on_cannot_be_deleted(client):
    add_submission("busy0002", status="processing")

    response = client.post("/delete/busy0002", follow_redirects=True)

    assert b"still being worked on" in response.data
    assert app_module.get_submission("busy0002") is not None


def test_a_failed_submission_can_be_deleted(client):
    add_submission("bad00001", status="error")

    client.post("/delete/bad00001", follow_redirects=True)

    assert app_module.get_submission("bad00001") is None


def test_the_status_page_offers_a_delete_button(client):
    add_submission("show0001")

    response = client.get("/status")

    assert b"/delete/show0001" in response.data


def test_the_status_page_hides_delete_while_a_file_is_queued(client):
    add_submission("hide0001", status="queued")

    response = client.get("/status")

    assert b"/delete/hide0001" not in response.data


def test_a_new_visitor_gets_their_own_id(tmp_path, monkeypatch):
    data = tmp_path / "data2"
    data.mkdir()
    monkeypatch.setattr(app_module, "DATABASE_FILE", str(data / "submissions.db"))
    app_module.setup_database()

    fresh_browser = app_module.app.test_client()
    fresh_browser.get("/status")

    with fresh_browser.session_transaction() as browser_session:
        assert browser_session.get("owner_id")


def test_two_visitors_do_not_see_each_other(client):
    add_submission("mine0003", original_filename="MyDeck.pptx")

    other_browser = app_module.app.test_client()
    response = other_browser.get("/status")

    assert b"MyDeck.pptx" not in response.data
