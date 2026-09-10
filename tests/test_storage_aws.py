import os
import sys
import time

import pytest

# make sure the tests can import the project files from the folder above
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import storage_aws
import lambda_handler


# stands in for the DynamoDB table, so these tests need no AWS account and cost
# nothing. It remembers what it was asked to do so the tests can look at it.
class FakeTable:
    def __init__(self, item=None):
        self.item = item
        self.updates = []
        self.deleted = []
        self.put = []

    def get_item(self, Key):
        if self.item is None:
            return {}

        return {"Item": self.item}

    def put_item(self, Item):
        self.put.append(Item)

    def update_item(self, **arguments):
        self.updates.append(arguments)

    def delete_item(self, Key):
        self.deleted.append(Key["id"])


@pytest.fixture
def table(monkeypatch):
    fake = FakeTable()
    monkeypatch.setattr(storage_aws, "get_table", lambda: fake)

    return fake


def a_row(**extras):
    row = {
        "id": "abc12345",
        "owner_id": "someone",
        "original_filename": "Lecture.pptx",
        "saved_filename": "abc12345_Lecture.pptx",
        "status": "done",
        "submitted_at": "2026-01-01 09:00",
        "retry_count": 0,
    }
    row.update(extras)

    return row


def test_a_row_becomes_the_shape_the_pages_expect():
    submission = storage_aws.item_to_submission(a_row())

    assert submission["id"] == "abc12345"
    assert submission["original_filename"] == "Lecture.pptx"
    # DynamoDB leaves a field out entirely rather than storing nothing in it
    assert submission["error_message"] is None
    assert submission["error_steps"] is None


def test_error_steps_come_back_as_a_list():
    submission = storage_aws.item_to_submission(a_row(
        status="error",
        error_message="It broke.",
        error_steps=["Try again.", "Then give up."],
    ))

    assert submission["error_steps"] == ["Try again.", "Then give up."]


def test_a_run_still_going_is_left_alone(table):
    submission = storage_aws.item_to_submission(a_row(status="processing", started_at=int(time.time())))

    assert submission["status"] == "processing"
    assert table.updates == []


def test_a_run_that_died_is_marked_as_failed(table):
    long_ago = int(time.time()) - storage_aws.STUCK_AFTER_SECONDS - 60
    submission = storage_aws.item_to_submission(a_row(status="processing", started_at=long_ago))

    assert submission["status"] == "error"
    assert "longer than we allow" in submission["error_message"]
    assert submission["error_steps"]
    # and the row itself is corrected, not just the copy the page is shown
    assert len(table.updates) == 1


def test_a_field_set_to_nothing_is_removed_rather_than_stored(table):
    storage_aws.update_submission("abc12345", {"status": "done", "error_message": None})

    expression = table.updates[0]["UpdateExpression"]

    assert "SET" in expression
    assert "REMOVE" in expression
    # status is a reserved word, so it can only appear as a placeholder
    assert "status" not in expression
    assert "status" in table.updates[0]["ExpressionAttributeNames"].values()


def test_updating_nothing_does_nothing(table):
    storage_aws.update_submission("abc12345", {})

    assert table.updates == []


def test_a_submission_without_an_owner_is_not_given_an_empty_one(table):
    storage_aws.add_submission({
        "id": "abc12345",
        "owner_id": None,
        "original_filename": "Old.pptx",
        "saved_filename": "abc12345_Old.pptx",
        "status": "error",
        "alt_text_filename": None,
        "error_message": None,
        "error_steps": None,
        "retry_count": 0,
        "submitted_at": "2026-01-01 09:00",
    })

    stored = table.put[0]

    assert "owner_id" not in stored
    assert "error_message" not in stored
    # the row deletes itself once this passes, which is how old files are cleared
    assert stored["expires_at"] > time.time()


def test_the_pipeline_finds_the_submission_id_in_the_key():
    assert lambda_handler.submission_id_from_key("abc12345_Lecture.pptx") == "abc12345"
    assert lambda_handler.submission_id_from_key("folder/abc12345_Lecture.pptx") == "abc12345"


def test_a_file_nobody_uploaded_through_the_site_has_no_id():
    # someone putting a deck in the bucket by hand still gets it processed, but
    # there is no row to update
    assert lambda_handler.submission_id_from_key("Lecture.pptx") is None


def test_the_pipeline_does_not_fall_over_when_there_is_no_table(monkeypatch):
    monkeypatch.setattr(lambda_handler, "SUBMISSIONS_TABLE", "")

    lambda_handler.update_row("abc12345", {"status": "processing"})


def test_a_table_problem_does_not_fail_the_deck(monkeypatch):
    class BrokenTable:
        def update_item(self, **arguments):
            raise RuntimeError("no")

    monkeypatch.setattr(lambda_handler, "get_table", lambda: BrokenTable())

    # a caption that has been written matters more than the progress bar
    lambda_handler.update_row("abc12345", {"captions_done": 1})


def test_a_finished_deck_gets_the_name_of_its_report(monkeypatch):
    written = []
    monkeypatch.setattr(lambda_handler, "update_row", lambda submission_id, changes: written.append(changes))

    lambda_handler.write_outcome("abc12345", "abc12345_Lecture.pptx", True, None, None)

    assert written[0]["status"] == "done"
    assert written[0]["alt_text_filename"] == "abc12345_Lecture_alt_text"


def test_a_failed_deck_keeps_its_explanation(monkeypatch):
    written = []
    monkeypatch.setattr(lambda_handler, "update_row", lambda submission_id, changes: written.append(changes))

    lambda_handler.write_outcome("abc12345", "abc12345_Lecture.pptx", False, "This file is empty.", ["Check the file."])

    assert written[0]["status"] == "error"
    assert written[0]["error_message"] == "This file is empty."
    assert written[0]["error_steps"] == ["Check the file."]


def test_a_failure_with_nothing_to_say_still_says_something(monkeypatch):
    written = []
    monkeypatch.setattr(lambda_handler, "update_row", lambda submission_id, changes: written.append(changes))

    lambda_handler.write_outcome("abc12345", "abc12345_Lecture.pptx", False, None, None)

    assert written[0]["error_message"]
    assert written[0]["error_steps"]
