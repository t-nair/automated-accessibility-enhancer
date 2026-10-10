import os
import sys
import shutil

import pytest
from pptx import Presentation
from pptx.util import Inches, Pt

# make sure the tests can import the project files from the folder above
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import title_detection
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


def test_the_report_is_not_empty(folders, fake_captions):
    input_folder, output_folder = folders
    copy_fixture("issue_missing_alt_text.pptx", input_folder)

    z_reorder.process_one_file("issue_missing_alt_text.pptx", input_folder, output_folder)

    with open(os.path.join(output_folder, "issue_missing_alt_text_alt_text"), encoding="utf-8") as f:
        assert "Alt Text:" in f.read()


def test_processing_logs_a_summary_with_the_output_paths(folders, fake_captions, caplog):
    input_folder, output_folder = folders
    copy_fixture("issue_missing_alt_text.pptx", input_folder)

    with caplog.at_level("INFO"):
        z_reorder.process_one_file("issue_missing_alt_text.pptx", input_folder, output_folder)

    summary = [m for m in caplog.messages if m.startswith("Summary:")]
    assert len(summary) == 1
    assert "image(s) captioned" in summary[0]
    assert "issue_missing_alt_text_updated.pptx" in summary[0]


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


# --- title detection and reading order ---------------------------------------


def build_slide(names):
    """A blank slide holding plain text boxes, in the order given."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    for name in names:
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
        box.name = name

    return slide


def names_in_order(slide):
    return [shape.name for shape in slide.shapes]


def test_a_title_that_is_not_first_is_moved_to_the_front():
    slide = build_slide(["Body 1", "Title 1"])

    assert z_reorder.move_titles_to_front(slide) == 1
    assert names_in_order(slide) == ["Title 1", "Body 1"]


def test_two_titles_keep_their_relative_order():
    # the old code re-read slide.shapes[0] on every pass while reordering, so the
    # second title anchored against the first one after it had already moved and
    # these came back as Title B, Title A
    slide = build_slide(["Body 1", "Title A", "Title B"])

    assert z_reorder.move_titles_to_front(slide) == 2
    assert names_in_order(slide) == ["Title A", "Title B", "Body 1"]


def test_a_title_already_at_the_front_is_left_alone():
    slide = build_slide(["Title 1", "Body 1"])

    assert z_reorder.move_titles_to_front(slide) == 0
    assert names_in_order(slide) == ["Title 1", "Body 1"]


def test_reordering_a_second_time_changes_nothing():
    slide = build_slide(["Body 1", "Title 1"])

    z_reorder.move_titles_to_front(slide)

    assert z_reorder.move_titles_to_front(slide) == 0
    assert names_in_order(slide) == ["Title 1", "Body 1"]


def test_a_lowercase_title_name_is_still_a_title():
    slide = build_slide(["Body 1", "title 1"])

    assert z_reorder.move_titles_to_front(slide) == 1
    assert names_in_order(slide) == ["title 1", "Body 1"]


def test_a_subtitle_is_not_treated_as_a_title():
    slide = build_slide(["Body 1", "Subtitle 2"])

    assert z_reorder.move_titles_to_front(slide) == 0
    assert names_in_order(slide) == ["Body 1", "Subtitle 2"]


def test_a_slide_with_no_title_is_left_alone():
    slide = build_slide(["Body 1", "Body 2"])

    assert z_reorder.move_titles_to_front(slide) == 0
    assert names_in_order(slide) == ["Body 1", "Body 2"]


def test_a_real_title_placeholder_is_found_even_when_renamed():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    title = slide.shapes.title
    title.name = "Header"

    assert z_reorder.is_title(title)


# --- titles that are not marked as titles --------------------------------------

BODY_TEXT = ("Evaporation, condensation and precipitation move water between the oceans, "
             "the air and the land, and the cycle repeats continuously.")


def build_text_slide(boxes):
    """A blank slide with plain text boxes given as (name, text, top in inches, font size), in z-order."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    for name, text, top, size in boxes:
        box = slide.shapes.add_textbox(Inches(0.7), Inches(top), Inches(8.6), Inches(1))
        box.name = name
        box.text_frame.text = text
        box.text_frame.paragraphs[0].runs[0].font.size = Pt(size)

    return slide


def test_a_title_typed_into_a_text_box_is_found_and_moved_to_the_front():
    # nothing here says "title": not the placeholder type, not the shape name
    slide = build_text_slide([("Body", BODY_TEXT, 2.5, 18), ("TextBox 7", "The Water Cycle", 0.4, 40)])

    assert z_reorder.move_titles_to_front(slide) == 1
    assert names_in_order(slide) == ["TextBox 7", "Body"]


def test_a_title_below_a_small_kicker_line_is_still_found():
    slide = build_text_slide([("Body", BODY_TEXT, 3.0, 18),
                              ("Kicker", "CHAPTER 3", 0.2, 12),
                              ("Heading", "Evaporation and Rain", 0.7, 36)])

    z_reorder.move_titles_to_front(slide)

    assert names_in_order(slide)[0] == "Heading"


def test_a_real_title_is_never_second_guessed(monkeypatch):
    monkeypatch.setattr(z_reorder, "guess_title", lambda slide: pytest.fail("guess_title was called"))
    slide = build_slide(["Body 1", "Title 1"])

    assert z_reorder.move_titles_to_front(slide) == 1


def test_a_slide_of_only_body_paragraphs_gets_no_guessed_title():
    slide = build_text_slide([("A", BODY_TEXT, 1.0, 18), ("B", BODY_TEXT, 3.5, 18)])

    assert z_reorder.move_titles_to_front(slide) == 0
    assert names_in_order(slide) == ["A", "B"]


def test_a_master_without_text_styles_does_not_break_font_size_lookup():
    # p:txStyles is optional; a placeholder with no size anywhere used to hit None.tag
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    master = slide.part.slide_layout.slide_master._element
    master.remove(master.find(title_detection.qn("p:txStyles")))

    assert title_detection.effective_pt(slide.shapes.title) == title_detection.DEFAULT_PT


def test_a_missing_title_model_leaves_the_slide_alone(monkeypatch):
    monkeypatch.setattr(title_detection, "_model", False)
    slide = build_text_slide([("Body", BODY_TEXT, 2.5, 18), ("TextBox 7", "The Water Cycle", 0.4, 40)])

    assert z_reorder.move_titles_to_front(slide) == 0
    assert names_in_order(slide) == ["Body", "TextBox 7"]


# --- alt text that is not really alt text ------------------------------------


@pytest.mark.parametrize("alt_text", [
    "diagram.png",
    "Screenshot 2024.PNG",
    "chart.jpeg",
    " figure.svg ",
])
def test_filename_style_alt_text_is_not_a_description(alt_text):
    assert z_reorder.is_placeholder_alt_text(alt_text, "Picture 3")


@pytest.mark.parametrize("name", [
    "Picture 3",
    "Image 5",
    "Graphic 7",
    "(Picture 12)",
    "Picture Placeholder 2",
    "Content Placeholder 4",
])
def test_alt_text_repeating_an_auto_generated_shape_name_is_not_a_description(name):
    assert z_reorder.is_placeholder_alt_text(name, name)


@pytest.mark.parametrize("name", ["(Picture 3", "Picture 3)"])
def test_an_unbalanced_parenthesis_is_not_an_auto_generated_name(name):
    assert not z_reorder.is_placeholder_alt_text(name, name)


@pytest.mark.parametrize("name", [
    "Water cycle diagram",
    "Free body diagram of the beam",
    "Beam",
])
def test_a_description_matching_an_author_chosen_shape_name_is_kept(name):
    # someone who renames a shape and writes the same text as alt text meant it.
    # Replacing real alt text is worse than leaving a thin description alone.
    assert not z_reorder.is_placeholder_alt_text(name, name)


def test_a_real_description_on_an_auto_named_shape_is_kept():
    assert not z_reorder.is_placeholder_alt_text("A blue square", "Picture 3")


@pytest.mark.parametrize("alt_text", [
    "A bar chart of enrolment by year",
    "",
    None,
])
def test_a_real_description_is_left_alone(alt_text):
    assert not z_reorder.is_placeholder_alt_text(alt_text, "Picture 3")


def set_alt_text_on_every_picture(prs, alt_text):
    for slide in prs.slides:
        for shape in slide.shapes:
            if z_reorder.is_picture(shape):
                z_reorder.set_picture_alt_text(shape, alt_text)


def test_a_picture_whose_alt_text_is_a_filename_still_needs_a_caption():
    prs = Presentation(os.path.join(FIXTURE_FOLDER, "issue_missing_alt_text.pptx"))
    set_alt_text_on_every_picture(prs, "diagram.png")

    assert z_reorder.count_images_needing_captions(prs) == 1


def test_a_picture_with_a_real_description_does_not_need_a_caption():
    prs = Presentation(os.path.join(FIXTURE_FOLDER, "issue_missing_alt_text.pptx"))
    set_alt_text_on_every_picture(prs, "A blue square on a white background")

    assert z_reorder.count_images_needing_captions(prs) == 0


# --- shape types python-pptx cannot resolve ----------------------------------


def test_a_text_box_is_not_a_picture():
    slide = build_slide(["Body 1"])

    assert not z_reorder.is_picture(slide.shapes[0])


def test_an_unresolvable_shape_type_does_not_raise():
    class Unresolvable:
        name = "Odd 1"

        @property
        def shape_type(self):
            raise NotImplementedError

    assert z_reorder.get_shape_type(Unresolvable()) is None
    assert not z_reorder.is_picture(Unresolvable())
    assert not z_reorder.needs_caption(Unresolvable())


def test_a_whole_folder_is_processed_in_one_run(folders, fake_captions):
    input_folder, output_folder = folders
    for filename in WORKING_FILES + BROKEN_FILES:
        copy_fixture(filename, input_folder)

    z_reorder.accessibility_processor(input_folder, output_folder)

    for filename in WORKING_FILES:
        assert os.path.exists(os.path.join(output_folder, filename.replace(".pptx", "_updated.pptx")))

    for filename in BROKEN_FILES:
        assert not os.path.exists(os.path.join(output_folder, filename.replace(".pptx", "_updated.pptx")))

    prs = Presentation(os.path.join(output_folder, "issue_title_not_first_in_order_updated.pptx"))
    assert [shape.name for shape in prs.slides[0].shapes] == ["Title 1", "Body Text 1"]


def make_deck_with_untitled_slides(path):
    prs = Presentation()
    titled = prs.slides.add_slide(prs.slide_layouts[5])
    titled.shapes.title.text = "Has a title"
    prs.slides.add_slide(prs.slide_layouts[5])  # title placeholder left empty
    prs.slides.add_slide(prs.slide_layouts[6])  # blank layout, no shapes
    prs.save(path)


def test_slides_without_a_title_are_listed_in_the_report(folders, fake_captions):
    input_folder, output_folder = folders
    make_deck_with_untitled_slides(os.path.join(input_folder, "untitled.pptx"))

    z_reorder.process_one_file("untitled.pptx", input_folder, output_folder)

    with open(os.path.join(output_folder, "untitled_alt_text"), encoding="utf-8") as f:
        report = f.read()

    assert "Slides with no title (WCAG 2.4.2): 2, 3" in report


def test_a_deck_where_every_slide_has_a_title_reports_nothing(folders, fake_captions):
    input_folder, output_folder = folders
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[5]).shapes.title.text = "Only slide"
    prs.save(os.path.join(input_folder, "titled.pptx"))

    z_reorder.process_one_file("titled.pptx", input_folder, output_folder)

    with open(os.path.join(output_folder, "titled_alt_text"), encoding="utf-8") as f:
        assert "no title" not in f.read()
