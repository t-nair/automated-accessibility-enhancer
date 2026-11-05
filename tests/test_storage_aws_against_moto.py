import io
import os
import sys

import pytest

# make sure the tests can import the project files from the folder above
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

boto3 = pytest.importorskip("boto3")
moto = pytest.importorskip("moto")

from moto import mock_aws

import storage_aws

UPLOADS = "test-uploads"
PROCESSED = "test-processed"
TABLE = "test-submissions"

# written out rather than read from storage_aws, so that renaming the index in the
# code without renaming it in the Terraform fails here instead of in production.
# This has to match infra/modules/dynamodb/main.tf.
OWNER_INDEX = "owner_id-submitted_at-index"


# moto stands in for S3 and DynamoDB in this process, so these run the real boto3
# calls without an AWS account and without costing anything. The point is to catch
# the things a hand written fake cannot: a wrong index name, an update expression
# DynamoDB rejects, a key that does not match where the file was put.
@pytest.fixture
def aws(monkeypatch):
    monkeypatch.setenv("AWS_ACCESS_KEY_ID", "testing")
    monkeypatch.setenv("AWS_SECRET_ACCESS_KEY", "testing")
    monkeypatch.setenv("AWS_DEFAULT_REGION", "us-west-2")

    with mock_aws():
        s3 = boto3.client("s3", region_name="us-west-2")

        for bucket in (UPLOADS, PROCESSED):
            s3.create_bucket(
                Bucket=bucket,
                CreateBucketConfiguration={"LocationConstraint": "us-west-2"},
            )

        dynamodb = boto3.client("dynamodb", region_name="us-west-2")
        dynamodb.create_table(
            TableName=TABLE,
            BillingMode="PAY_PER_REQUEST",
            AttributeDefinitions=[
                {"AttributeName": "id", "AttributeType": "S"},
                {"AttributeName": "owner_id", "AttributeType": "S"},
                {"AttributeName": "submitted_at", "AttributeType": "S"},
            ],
            KeySchema=[{"AttributeName": "id", "KeyType": "HASH"}],
            GlobalSecondaryIndexes=[{
                "IndexName": OWNER_INDEX,
                "KeySchema": [
                    {"AttributeName": "owner_id", "KeyType": "HASH"},
                    {"AttributeName": "submitted_at", "KeyType": "RANGE"},
                ],
                "Projection": {"ProjectionType": "ALL"},
            }],
        )

        # the module builds its clients once and keeps them, so clear them out and
        # point it at the buckets and table moto just made
        monkeypatch.setattr(storage_aws, "_s3", None)
        monkeypatch.setattr(storage_aws, "_table", None)
        monkeypatch.setattr(storage_aws, "UPLOADS_BUCKET", UPLOADS)
        monkeypatch.setattr(storage_aws, "PROCESSED_BUCKET", PROCESSED)
        monkeypatch.setattr(storage_aws, "SUBMISSIONS_TABLE", TABLE)

        yield s3


def a_submission(submission_id="abc12345", owner_id="visitor-one", submitted_at="2026-01-01 09:00"):
    return {
        "id": submission_id,
        "owner_id": owner_id,
        "original_filename": "Lecture.pptx",
        "saved_filename": submission_id + "_Lecture.pptx",
        "status": "queued",
        "alt_text_filename": None,
        "error_message": None,
        "error_steps": None,
        "retry_count": 0,
        "submitted_at": submitted_at,
    }


def test_a_submission_survives_a_round_trip(aws):
    storage_aws.add_submission(a_submission())

    came_back = storage_aws.get_submission("abc12345")

    assert came_back["original_filename"] == "Lecture.pptx"
    assert came_back["status"] == "queued"
    assert came_back["error_steps"] is None


def test_an_unknown_submission_comes_back_as_nothing(aws):
    assert storage_aws.get_submission("nope") is None


def test_the_index_returns_one_visitor_newest_first(aws):
    storage_aws.add_submission(a_submission("aaa11111", "visitor-one", "2026-01-01 09:00"))
    storage_aws.add_submission(a_submission("bbb22222", "visitor-one", "2026-01-03 09:00"))
    storage_aws.add_submission(a_submission("ccc33333", "visitor-two", "2026-01-02 09:00"))

    mine = storage_aws.get_submissions_for_owner("visitor-one")

    assert len(mine) == 2
    assert mine[0]["id"] == "bbb22222"
    assert mine[1]["id"] == "aaa11111"


def test_a_visitor_with_nothing_gets_an_empty_list(aws):
    assert storage_aws.get_submissions_for_owner("nobody") == []


def test_updating_changes_only_what_it_is_given(aws):
    storage_aws.add_submission(a_submission())

    storage_aws.update_submission("abc12345", {"status": "done", "alt_text_filename": "abc12345_Lecture_alt_text"})
    came_back = storage_aws.get_submission("abc12345")

    assert came_back["status"] == "done"
    assert came_back["alt_text_filename"] == "abc12345_Lecture_alt_text"
    assert came_back["original_filename"] == "Lecture.pptx"


def test_error_steps_go_in_and_come_out_as_a_list(aws):
    storage_aws.add_submission(a_submission())

    storage_aws.update_submission("abc12345", {
        "status": "error",
        "error_message": "This file is empty.",
        "error_steps": ["Check the file.", "Upload it again."],
    })

    came_back = storage_aws.get_submission("abc12345")

    assert came_back["error_steps"] == ["Check the file.", "Upload it again."]


def test_clearing_a_field_removes_it(aws):
    storage_aws.add_submission(a_submission())
    storage_aws.update_submission("abc12345", {"error_message": "It broke."})

    storage_aws.update_submission("abc12345", {"status": "done", "error_message": None})
    came_back = storage_aws.get_submission("abc12345")

    assert came_back["status"] == "done"
    assert came_back["error_message"] is None


def test_deleting_removes_the_row(aws):
    storage_aws.add_submission(a_submission())

    storage_aws.delete_submission("abc12345")

    assert storage_aws.get_submission("abc12345") is None


def test_an_upload_lands_where_the_pipeline_will_look_for_it(aws):
    storage_aws.save_upload(io.BytesIO(b"pretend this is a deck"), "abc12345_Lecture.pptx")

    assert storage_aws.upload_exists("abc12345_Lecture.pptx")

    stored = aws.get_object(Bucket=UPLOADS, Key="abc12345_Lecture.pptx")

    assert stored["Body"].read() == b"pretend this is a deck"


def test_a_missing_upload_is_reported_as_missing(aws):
    assert storage_aws.upload_exists("never_uploaded.pptx") is False


def test_the_report_is_read_back_from_the_processed_bucket(aws):
    aws.put_object(Bucket=PROCESSED, Key="abc12345_Lecture_alt_text", Body=b'{"shapes": []}')

    submission = a_submission()
    submission["alt_text_filename"] = "abc12345_Lecture_alt_text"

    assert storage_aws.read_report(submission) == '{"shapes": []}'


def test_a_missing_report_does_not_raise(aws):
    submission = a_submission()
    submission["alt_text_filename"] = "not_there_alt_text"

    assert storage_aws.read_report(submission) is None


def test_the_finished_deck_is_found_under_the_name_the_pipeline_writes(aws):
    aws.put_object(Bucket=PROCESSED, Key="abc12345_Lecture_updated.pptx", Body=b"fixed deck")

    assert storage_aws.processed_file_is_ready(a_submission()) is True


def test_a_deck_that_is_not_finished_is_not_offered(aws):
    assert storage_aws.processed_file_is_ready(a_submission()) is False


def test_the_download_link_points_at_the_finished_deck(aws):
    aws.put_object(Bucket=PROCESSED, Key="abc12345_Lecture_updated.pptx", Body=b"fixed deck")

    response = storage_aws.send_processed_file(a_submission())
    link = response.headers["Location"]

    assert response.status_code == 302
    assert "abc12345_Lecture_updated.pptx" in link
    # the browser is told to save it under the name the person uploaded
    assert "accessible_Lecture.pptx" in link


def test_deleting_a_submission_clears_every_file_it_made(aws):
    storage_aws.save_upload(io.BytesIO(b"deck"), "abc12345_Lecture.pptx")
    aws.put_object(Bucket=PROCESSED, Key="abc12345_Lecture.pptx", Body=b"copy")
    aws.put_object(Bucket=PROCESSED, Key="abc12345_Lecture_updated.pptx", Body=b"fixed")
    aws.put_object(Bucket=PROCESSED, Key="abc12345_Lecture_alt_text", Body=b"report")

    storage_aws.delete_submission_files(a_submission())

    assert storage_aws.upload_exists("abc12345_Lecture.pptx") is False
    assert aws.list_objects_v2(Bucket=PROCESSED).get("KeyCount") == 0


def test_a_retry_puts_the_upload_back_so_the_pipeline_runs_again(aws):
    storage_aws.save_upload(io.BytesIO(b"deck"), "abc12345_Lecture.pptx")

    storage_aws.restart_processing("abc12345_Lecture.pptx")

    # the deck is still there, which is what S3 announces a second time
    stored = aws.get_object(Bucket=UPLOADS, Key="abc12345_Lecture.pptx")

    assert stored["Body"].read() == b"deck"
    assert "retried_at" in stored["Metadata"]
