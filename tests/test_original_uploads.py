import os

import pytest
from pptx.presentation import Presentation

# test_app puts the project folder on the path, so it is imported first
from test_app import add_submission, client, copy_into  # noqa: F401 (client is a fixture)
import app as app_module
import z_reorder


@pytest.fixture
def fake_captions(monkeypatch):
    monkeypatch.setattr(z_reorder, "generate_image_caption", lambda shape, slide_text="": "A test caption")


def read_bytes(path):
    with open(path, "rb") as f:
        return f.read()


def upload_path(submission_id):
    return os.path.join(app_module.UPLOAD_FOLDER, submission_id + "_Lecture.pptx")


def test_a_finished_file_keeps_its_original_upload(client, fake_captions):
    add_submission("keep0001", status="queued", alt_text_filename=None)
    copy_into("control_no_issues.pptx", upload_path("keep0001"))
    before = read_bytes(upload_path("keep0001"))

    app_module.process_one_submission("keep0001", "keep0001_Lecture.pptx")

    assert app_module.get_submission("keep0001")["status"] == "done"
    assert read_bytes(upload_path("keep0001")) == before
    assert not os.path.exists(os.path.join(app_module.PROCESSED_FOLDER, "keep0001_Lecture.pptx"))


def test_a_failed_file_can_be_retried_and_its_original_is_never_changed(client, fake_captions, monkeypatch):
    add_submission("fail0001", status="queued", alt_text_filename=None)
    copy_into("control_no_issues.pptx", upload_path("fail0001"))
    before = read_bytes(upload_path("fail0001"))

    real_save = Presentation.save
    disk_full = True

    def save(prs, path):
        if disk_full:
            raise OSError("disk full")
        real_save(prs, path)

    monkeypatch.setattr(Presentation, "save", save)
    app_module.process_one_submission("fail0001", "fail0001_Lecture.pptx")

    assert app_module.get_submission("fail0001")["status"] == "error"
    assert read_bytes(upload_path("fail0001")) == before

    client.post("/retry/fail0001")
    assert client.queued == ["fail0001"]

    disk_full = False
    app_module.process_one_submission("fail0001", "fail0001_Lecture.pptx")

    assert app_module.get_submission("fail0001")["status"] == "done"
    assert read_bytes(upload_path("fail0001")) == before


def test_retry_finds_an_original_an_older_version_moved_into_processed(client):
    add_submission("old00001", status="error")
    copy_into("control_no_issues.pptx", os.path.join(app_module.PROCESSED_FOLDER, "old00001_Lecture.pptx"))

    client.post("/retry/old00001")

    assert client.queued == ["old00001"]


def test_startup_requeues_unfinished_work_and_flags_lost_originals(client, monkeypatch):
    add_submission("wait0001", status="queued")
    add_submission("half0001", status="processing")
    add_submission("gone0001", status="processing")
    add_submission("done0001", status="done")
    copy_into("control_no_issues.pptx", upload_path("wait0001"))
    copy_into("control_no_issues.pptx", upload_path("half0001"))

    requeued = []
    monkeypatch.setattr(app_module, "start_processing", lambda submission_id, saved_name: requeued.append(submission_id))

    app_module.recover_unfinished_submissions()

    assert requeued == ["wait0001", "half0001"]
    assert app_module.get_submission("half0001")["status"] == "queued"
    assert app_module.get_submission("gone0001")["status"] == "error"
    assert "no longer available" in app_module.get_submission("gone0001")["error_message"]
    assert app_module.get_submission("done0001")["status"] == "done"
