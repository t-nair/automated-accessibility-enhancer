import os
import sys
import shutil
import json

import pytest
from pptx import Presentation
from pptx.util import Inches

# make sure the tests can import the project files from the folder above
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import z_reorder

FIXTURE_FOLDER = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Error Test PPTX")

# these fixture files should all fail to open
BROKEN_FILES = [
    "corrupt_empty_file.pptx",
    "corrupt_random_bytes.pptx",
    "corrupt_plain_text_renamed.pptx",
    "corrupt_zip_missing_pptx_parts.pptx",
    "corrupt_truncated_valid_file.pptx",
]

# these fixture files should all process successfully
WORKING_FILES = [
    "control_no_issues.pptx",
    "issue_missing_alt_text.pptx",
    "issue_no_title_shape.pptx",
    "issue_title_not_first_in_order.pptx",
    "issue_empty_slide_no_shapes.pptx",
    "issue_multiple_images_no_alt_text.pptx",
    "issue_all_issues_combined.pptx",
]


@pytest.fixture
def fake_captions(monkeypatch):
    def fake_caption(shape, slide_text=""):
        return "A test caption"

    monkeypatch.setattr(z_reorder, "generate_image_caption", fake_caption)


@pytest.fixture
def folders(tmp_path):
    input_folder = tmp_path / "uploads"
    output_folder = tmp_path / "processed"
    input_folder.mkdir()
    output_folder.mkdir()

    return str(input_folder), str(output_folder)


def copy_fixture(filename, input_folder):
    source = os.path.join(FIXTURE_FOLDER, filename)
    shutil.copy(source, os.path.join(input_folder, filename))


def test_tidy_caption_removes_filler_at_the_start():
    assert z_reorder.tidy_caption("The image shows a red car") == "A red car"


def test_tidy_caption_capitalises_the_first_letter():
    assert z_reorder.tidy_caption("there are two people talking") == "Two people talking"


def test_tidy_caption_leaves_a_normal_sentence_alone():
    assert z_reorder.tidy_caption("A bronze statue on a pedestal") == "A bronze statue on a pedestal"


def test_tidy_caption_handles_an_empty_caption():
    assert z_reorder.tidy_caption("") == ""


def test_build_caption_prompt_without_slide_text():
    prompt = z_reorder.build_caption_prompt("")

    assert prompt == z_reorder.CAPTION_PROMPT


def test_build_caption_prompt_includes_the_slide_text():
    prompt = z_reorder.build_caption_prompt("Root Causes of the Uprising")

    assert "Root Causes of the Uprising" in prompt


def test_get_slide_text_collects_the_words_on_the_slide():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
    box.text_frame.text = "Soviet Response"

    assert "Soviet Response" in z_reorder.get_slide_text(slide)


def test_get_slide_text_is_cut_off_when_the_slide_is_very_long():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
    box.text_frame.text = "word " * 500

    slide_text = z_reorder.get_slide_text(slide)

    assert len(slide_text) <= z_reorder.SLIDE_CONTEXT_MAX_CHARS


def test_get_slide_text_is_empty_for_a_slide_with_no_text():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    assert z_reorder.get_slide_text(slide) == ""


JUNK_ALT_TEXTS = [
    "Image Google Shape;1447;p111",
    "Three Right Hand Rules of ...",
    "Motor Speed ...",
    "Erno Gero | Turtledove | Fandom",
    "International Baccalaureate - Wikipedia",
    "How Electric Motors Work | HowStuffWorks",
    "https://example.com/diagram",
    "diagram.png",
    "",
    "   ",
]

REAL_ALT_TEXTS = [
    "A blue square used as a placeholder image",
    "A constant current I flows in the long straight wire in the direction shown.",
    "a woman sitting at a desk with a computer",
    "Two silhouettes of human heads, one pink and one blue",
    "A bronze bust of a man on a pedestal",
]


@pytest.mark.parametrize("alt_text", JUNK_ALT_TEXTS)
def test_unhelpful_alt_text_is_spotted(alt_text):
    assert z_reorder.looks_like_junk_alt_text(alt_text, "Picture 2") is True


@pytest.mark.parametrize("alt_text", REAL_ALT_TEXTS)
def test_real_alt_text_is_left_alone(alt_text):
    assert z_reorder.looks_like_junk_alt_text(alt_text, "Picture 2") is False


def test_our_own_placeholder_counts_as_unhelpful():
    assert z_reorder.looks_like_junk_alt_text("Image Picture 4", "Picture 4") is True


def test_a_description_that_mentions_a_picture_is_kept():
    # this should not be mistaken for the placeholder above, because the shape differs
    assert z_reorder.looks_like_junk_alt_text("Image Picture 4", "Picture 9") is False


def get_problems(filename, folders):
    input_folder, output_folder = folders
    copy_fixture(filename, input_folder)
    z_reorder.process_one_file(filename, input_folder, output_folder)

    report_path = os.path.join(output_folder, os.path.splitext(filename)[0] + "_alt_text")

    with open(report_path, encoding="utf-8") as f:
        return json.load(f)["problems"]


def kinds_found(problems):
    kinds = []

    for problem in problems:
        kinds.append(problem["kind"])

    return kinds


def test_a_slide_with_no_title_is_reported(folders, fake_captions):
    problems = get_problems("issue_checks_sampler.pptx", folders)

    assert "no title" in kinds_found(problems)


def test_a_repeated_title_is_reported(folders, fake_captions):
    problems = get_problems("issue_checks_sampler.pptx", folders)

    assert "repeated title" in kinds_found(problems)


def test_an_unclear_link_is_reported(folders, fake_captions):
    problems = get_problems("issue_checks_sampler.pptx", folders)

    assert "unclear link" in kinds_found(problems)


def test_a_table_without_a_header_row_is_reported(folders, fake_captions):
    problems = get_problems("issue_checks_sampler.pptx", folders)

    assert "table without a header row" in kinds_found(problems)


def test_every_problem_says_which_slide_it_is_on(folders, fake_captions):
    problems = get_problems("issue_checks_sampler.pptx", folders)

    for problem in problems:
        assert problem["slide"] > 0
        assert problem["detail"]


def test_a_tidy_deck_has_nothing_to_report(folders, fake_captions):
    problems = get_problems("control_no_issues.pptx", folders)

    assert kinds_found(problems) == []


def test_an_empty_slide_is_reported(folders, fake_captions):
    problems = get_problems("issue_empty_slide_no_shapes.pptx", folders)

    assert "empty slide" in kinds_found(problems)


def test_link_text_that_says_something_is_left_alone():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    run = box.text_frame.paragraphs[0].add_run()
    run.text = "the course syllabus"
    run.hyperlink.address = "https://example.com/syllabus"

    assert z_reorder.find_vague_links(slide) == []


def get_first_picture(path):
    prs = Presentation(path)

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.shape_type == 13:
                return shape

    return None


def test_missing_alt_text_is_read_as_empty():
    picture = get_first_picture(os.path.join(FIXTURE_FOLDER, "issue_missing_alt_text.pptx"))

    assert not z_reorder.get_picture_alt_text(picture)


def test_existing_alt_text_is_read_back():
    picture = get_first_picture(os.path.join(FIXTURE_FOLDER, "control_no_issues.pptx"))

    assert z_reorder.get_picture_alt_text(picture)


def test_alt_text_can_be_set_and_read_again():
    picture = get_first_picture(os.path.join(FIXTURE_FOLDER, "issue_missing_alt_text.pptx"))

    z_reorder.set_picture_alt_text(picture, "A blue square")

    assert z_reorder.get_picture_alt_text(picture) == "A blue square"


def test_count_images_needing_captions():
    prs = Presentation(os.path.join(FIXTURE_FOLDER, "issue_multiple_images_no_alt_text.pptx"))

    assert z_reorder.count_images_needing_captions(prs) == 3


def test_pictures_that_already_have_alt_text_are_not_counted():
    prs = Presentation(os.path.join(FIXTURE_FOLDER, "control_no_issues.pptx"))

    assert z_reorder.count_images_needing_captions(prs) == 0


@pytest.mark.parametrize("filename", BROKEN_FILES)
def test_broken_files_fail_with_a_message_and_steps(filename, folders, fake_captions):
    input_folder, output_folder = folders
    copy_fixture(filename, input_folder)

    was_successful, message, steps = z_reorder.process_one_file(filename, input_folder, output_folder)

    assert was_successful is False
    assert message
    assert len(steps) > 0


def test_an_empty_file_says_it_is_empty(folders, fake_captions):
    input_folder, output_folder = folders
    copy_fixture("corrupt_empty_file.pptx", input_folder)

    was_successful, message, steps = z_reorder.process_one_file("corrupt_empty_file.pptx", input_folder, output_folder)

    assert was_successful is False
    assert "empty" in message.lower()


def test_a_missing_file_fails_politely(folders, fake_captions):
    input_folder, output_folder = folders

    was_successful, message, steps = z_reorder.process_one_file("not_here.pptx", input_folder, output_folder)

    assert was_successful is False
    assert message


@pytest.mark.parametrize("filename", WORKING_FILES)
def test_working_files_are_processed(filename, folders, fake_captions):
    input_folder, output_folder = folders
    copy_fixture(filename, input_folder)

    was_successful, message, steps = z_reorder.process_one_file(filename, input_folder, output_folder)

    assert was_successful is True
    assert message is None
    assert steps is None


def test_processing_writes_an_updated_file_and_a_report(folders, fake_captions):
    input_folder, output_folder = folders
    copy_fixture("issue_missing_alt_text.pptx", input_folder)

    z_reorder.process_one_file("issue_missing_alt_text.pptx", input_folder, output_folder)

    assert os.path.exists(os.path.join(output_folder, "issue_missing_alt_text_updated.pptx"))
    assert os.path.exists(os.path.join(output_folder, "issue_missing_alt_text_alt_text"))


def test_the_original_upload_is_moved_out_of_the_input_folder(folders, fake_captions):
    input_folder, output_folder = folders
    copy_fixture("control_no_issues.pptx", input_folder)

    z_reorder.process_one_file("control_no_issues.pptx", input_folder, output_folder)

    assert not os.path.exists(os.path.join(input_folder, "control_no_issues.pptx"))
    assert os.path.exists(os.path.join(output_folder, "control_no_issues.pptx"))


def test_a_missing_description_is_written_into_the_updated_file(folders, fake_captions):
    input_folder, output_folder = folders
    copy_fixture("issue_missing_alt_text.pptx", input_folder)

    z_reorder.process_one_file("issue_missing_alt_text.pptx", input_folder, output_folder)

    picture = get_first_picture(os.path.join(output_folder, "issue_missing_alt_text_updated.pptx"))

    assert z_reorder.get_picture_alt_text(picture) == "A test caption"


def test_alt_text_that_was_already_there_is_kept(folders, fake_captions):
    input_folder, output_folder = folders
    copy_fixture("control_no_issues.pptx", input_folder)

    z_reorder.process_one_file("control_no_issues.pptx", input_folder, output_folder)

    picture = get_first_picture(os.path.join(output_folder, "control_no_issues_updated.pptx"))

    assert z_reorder.get_picture_alt_text(picture) != "A test caption"


def test_the_title_is_moved_to_the_front_of_the_slide(folders, fake_captions):
    input_folder, output_folder = folders
    copy_fixture("issue_title_not_first_in_order.pptx", input_folder)

    z_reorder.process_one_file("issue_title_not_first_in_order.pptx", input_folder, output_folder)

    prs = Presentation(os.path.join(output_folder, "issue_title_not_first_in_order_updated.pptx"))
    shape_names = []

    for shape in prs.slides[0].shapes:
        shape_names.append(shape.name)

    assert "Title" in shape_names[0]


def test_a_slide_with_no_shapes_does_not_stop_processing(folders, fake_captions):
    input_folder, output_folder = folders
    copy_fixture("issue_empty_slide_no_shapes.pptx", input_folder)

    was_successful, message, steps = z_reorder.process_one_file(
        "issue_empty_slide_no_shapes.pptx", input_folder, output_folder
    )

    assert was_successful is True


def test_a_caption_failure_falls_back_to_a_placeholder(folders, monkeypatch):
    input_folder, output_folder = folders
    copy_fixture("issue_missing_alt_text.pptx", input_folder)

    def broken_caption(shape, slide_text=""):
        raise RuntimeError("pretend the model failed")

    monkeypatch.setattr(z_reorder, "generate_image_caption", broken_caption)

    was_successful, message, steps = z_reorder.process_one_file(
        "issue_missing_alt_text.pptx", input_folder, output_folder
    )

    assert was_successful is True

    picture = get_first_picture(os.path.join(output_folder, "issue_missing_alt_text_updated.pptx"))

    assert "Image" in z_reorder.get_picture_alt_text(picture)


def test_progress_is_reported_for_every_image(folders, fake_captions):
    input_folder, output_folder = folders
    copy_fixture("issue_multiple_images_no_alt_text.pptx", input_folder)

    updates = []

    def remember_progress(done, total):
        updates.append((done, total))

    z_reorder.process_one_file(
        "issue_multiple_images_no_alt_text.pptx", input_folder, output_folder, remember_progress
    )

    # one update before starting, then one after each of the three images
    assert updates[0] == (0, 3)
    assert updates[-1] == (3, 3)


def test_a_ppt_file_is_rejected_politely_when_libreoffice_is_missing(folders, monkeypatch, fake_captions):
    input_folder, output_folder = folders

    # a .ppt cannot be opened directly, so a copied .pptx stands in for the upload
    shutil.copy(
        os.path.join(FIXTURE_FOLDER, "control_no_issues.pptx"),
        os.path.join(input_folder, "old_slides.ppt")
    )

    monkeypatch.setattr(z_reorder, "find_libreoffice", lambda: None)

    was_successful, message, steps = z_reorder.process_one_file("old_slides.ppt", input_folder, output_folder)

    assert was_successful is False
    assert ".ppt" in message
    assert len(steps) > 0
