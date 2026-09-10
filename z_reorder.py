import os
import io
import json
import re
import time
import shutil
import logging
import subprocess
import torch
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from PIL import Image
from transformers import Qwen2VLForConditionalGeneration, AutoProcessor

logging.basicConfig(
    filename="pipeline.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# these libraries log every network request at INFO level, which just buries our own log lines
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("huggingface_hub").setLevel(logging.WARNING)
logging.getLogger("filelock").setLevel(logging.WARNING)

MAX_ATTEMPTS = 3
RETRY_DELAY_SECONDS = 2

# old .ppt files are a completely different format that python-pptx cannot read,
# so LibreOffice converts them to .pptx before the normal pipeline runs
LIBREOFFICE_LOCATIONS = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]
CONVERT_TIMEOUT_SECONDS = 120

# PowerPoint fills a picture's alt text with its source filename when the picture is
# inserted, so a deck can look fully described while telling a screen reader nothing.
# Alt text ending in one of these is treated as missing and gets a real caption.
IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff",
                    ".emf", ".wmf", ".svg", ".webp")

# The names PowerPoint gives a picture on its own, e.g. "Picture 3", "Image 5".
# Alt text that only repeats a name like this says nothing, but alt text that
# repeats a name the author chose deliberately may well be a real description,
# so only auto-generated names count as placeholder text.
# the parentheses have to balance: "(Picture 3" is not a name PowerPoint writes
_AUTO_SHAPE_NAME = r"(?:picture|image|graphic|picture placeholder|content placeholder)\s*\d+"
AUTO_SHAPE_NAME = re.compile(rf"^(?:{_AUTO_SHAPE_NAME}|\({_AUTO_SHAPE_NAME}\))$")

# a real title placeholder is authoritative, the shape name is only a fallback
TITLE_PLACEHOLDERS = (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)

CAPTION_MODEL_NAME = "Qwen/Qwen2-VL-2B-Instruct"

# the model is told once, here, what kind of description we want. Faculty never see
# or type this, it is just how we ask for a caption that works as alt text.
# course slides are full of diagrams, and asking about them directly gives much better
# descriptions than a general request does. Tested on physics diagrams from a real
# lecture deck: the model started naming what a diagram shows and reading its labels,
# instead of only describing the shapes on it. Photographs were not made any worse.
CAPTION_PROMPT = (
    "Write alt text for this image in two or three short sentences. "
    "If it is a diagram, chart or equation, say what kind it is, quote the labels that appear on it "
    "exactly as they are written, and say how the parts are arranged or connected. "
    "If it is a photograph, describe what is happening in it. "
    "Only describe what is actually visible. Do not invent labels, numbers or values."
)

# the words on the slide usually say what the image is there to show, which helps the
# model describe it correctly. Long slides get cut short so the prompt stays focused.
SLIDE_CONTEXT_MAX_CHARS = 400

# keeps a very large slide image from using far more GPU memory than it needs
CAPTION_MIN_PIXELS = 256 * 28 * 28
CAPTION_MAX_PIXELS = 768 * 28 * 28

CAPTION_DEVICE = torch.device("cuda" if torch.cuda.is_available() else "cpu")

# half precision saves GPU memory, but on CPU it is slower than normal precision
if CAPTION_DEVICE.type == "cuda":
    CAPTION_DTYPE = torch.float16
else:
    CAPTION_DTYPE = torch.float32

# loaded the first time we actually need to caption an image, so app startup stays fast
caption_processor = None
caption_model = None


def get_caption_model():
    global caption_processor, caption_model

    if caption_model is None:
        logging.info(f"Loading image captioning model ({CAPTION_MODEL_NAME}) on {CAPTION_DEVICE}. This can take a while the first time.")

        if CAPTION_DEVICE.type == "cuda":
            logging.info(f"Using GPU: {torch.cuda.get_device_name(0)}")
        else:
            logging.warning("No GPU available, so captions will be generated on the CPU. This is slower.")

        caption_processor = AutoProcessor.from_pretrained(
            CAPTION_MODEL_NAME,
            min_pixels=CAPTION_MIN_PIXELS,
            max_pixels=CAPTION_MAX_PIXELS
        )
        caption_model = Qwen2VLForConditionalGeneration.from_pretrained(CAPTION_MODEL_NAME, dtype=CAPTION_DTYPE)
        caption_model.to(CAPTION_DEVICE)
        caption_model.eval()

    return caption_processor, caption_model


def tidy_caption(caption):
    # the model often starts captions with filler like "the image shows", which
    # wastes time when a screen reader announces it
    filler_starts = [
        "the image is ",
        "the image shows ",
        "the image features ",
        "the image depicts ",
        "this image shows ",
        "there is ",
        "there are ",
        "an image of ",
        "a picture of ",
    ]

    caption = caption.strip()

    for filler in filler_starts:
        if caption.lower().startswith(filler):
            caption = caption[len(filler):]
            break

    caption = caption.strip()

    # removing the filler can leave the sentence starting with a lower case letter
    if len(caption) > 0:
        caption = caption[0].upper() + caption[1:]

    return caption


def get_slide_text(slide):
    pieces = []

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        text = shape.text.strip()

        if text:
            pieces.append(text)

    slide_text = " ".join(pieces)
    slide_text = " ".join(slide_text.split())

    if len(slide_text) > SLIDE_CONTEXT_MAX_CHARS:
        slide_text = slide_text[:SLIDE_CONTEXT_MAX_CHARS]

    return slide_text


# the same words, but not cut short, because the reading level is worked out from
# everything on the slide rather than from a prompt sized piece of it
def get_all_slide_text(slide):
    pieces = []

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        text = shape.text.strip()

        if text:
            pieces.append(text)

    return " ".join(pieces)


def build_caption_prompt(slide_text):
    if not slide_text:
        return CAPTION_PROMPT

    return (
        CAPTION_PROMPT
        + " For context, the slide this image appears on says: \""
        + slide_text
        + "\". Use that only to help you understand what you are looking at. "
        "Do not repeat the slide text, and do not claim anything from it that you cannot actually see in the image."
    )


def generate_image_caption(shape, slide_text=""):
    image_bytes = shape.image.blob
    image = Image.open(io.BytesIO(image_bytes)).convert("RGB")

    processor, model = get_caption_model()
    prompt = build_caption_prompt(slide_text)

    messages = [{"role": "user", "content": [{"type": "image"}, {"type": "text", "text": prompt}]}]
    chat_text = processor.apply_chat_template(messages, tokenize=False, add_generation_prompt=True)
    inputs = processor(text=[chat_text], images=[image], return_tensors="pt").to(CAPTION_DEVICE)

    start_time = time.perf_counter()

    with torch.inference_mode():
        output = model.generate(**inputs, max_new_tokens=70, do_sample=False)

    seconds_taken = time.perf_counter() - start_time

    # the model echoes the prompt back, so only keep the tokens it added
    new_tokens = output[0][len(inputs.input_ids[0]):]
    caption = processor.decode(new_tokens, skip_special_tokens=True)

    logging.info(f"Caption generated in {seconds_taken:.2f} seconds on {CAPTION_DEVICE}.")

    return tidy_caption(caption)


def accessibility_processor(directory, new_directory):
    files = os.listdir(directory)
    pptx_files = []

    for filename in files:
        if filename.endswith(".pptx"):
            pptx_files.append(filename)

    logging.info(f"Found {len(pptx_files)} pptx file(s) in {directory}")

    if len(pptx_files) == 0:
        logging.warning("No .pptx files found. Nothing to process.")
        return

    for filename in pptx_files:
        process_one_file(filename, directory, new_directory)


def find_libreoffice():
    for location in LIBREOFFICE_LOCATIONS:
        if os.path.exists(location):
            return location

    # fall back to whatever is on the system PATH
    return shutil.which("soffice")


# gives back the path to the converted file, or None and a message explaining why not
def convert_ppt_to_pptx(filepath, directory):
    soffice = find_libreoffice()

    if soffice is None:
        logging.error("Cannot convert a .ppt file because LibreOffice was not found.")
        return None, "This is an older .ppt file, and the converter we use for those isn't available right now.", [
            "Open the file in PowerPoint and use File > Save As to save it as a .pptx file.",
            "Upload the .pptx version instead.",
        ]

    logging.info(f"Converting {os.path.basename(filepath)} from .ppt to .pptx using LibreOffice.")

    try:
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pptx", "--outdir", directory, filepath],
            capture_output=True,
            timeout=CONVERT_TIMEOUT_SECONDS
        )
    except subprocess.TimeoutExpired:
        logging.error(f"Converting {filepath} took too long and was stopped.")
        return None, "This older .ppt file took too long to convert.", [
            "Open the file in PowerPoint and save it as a .pptx file.",
            "Upload the .pptx version instead.",
        ]

    converted_path = os.path.join(directory, os.path.splitext(os.path.basename(filepath))[0] + ".pptx")

    if not os.path.exists(converted_path):
        logging.error(f"Conversion of {filepath} did not produce a file. LibreOffice said: {result.stderr[:300]}")
        return None, "We could not convert this older .ppt file.", [
            "Open the file in PowerPoint to check that it still opens.",
            "Use File > Save As to save it as a .pptx file.",
            "Upload the .pptx version instead.",
        ]

    logging.info(f"Converted to {os.path.basename(converted_path)}")
    return converted_path, None, None


def get_open_error_details(error):
    if isinstance(error, FileNotFoundError):
        return (
            "We could no longer find the uploaded file.",
            [
                "This is usually a temporary hiccup.",
                "Upload the presentation again.",
            ]
        )

    if isinstance(error, PermissionError):
        return (
            "The file looks like it's open somewhere else, so we couldn't read it.",
            [
                "Close the file if it's open in PowerPoint or another program.",
                "Wait a few seconds, then try again.",
            ]
        )

    return (
        "This doesn't look like a valid PowerPoint (.pptx) file. It may be a different file type, or it may be corrupted.",
        [
            "Open the file in PowerPoint on your computer to check that it opens there.",
            "If it opens, use File > Save As and save a new copy as a .pptx file.",
            "If it does not open in PowerPoint either, the file itself is damaged. Try a backup copy or re-export it.",
            "Upload the new copy.",
        ]
    )


# gives back a (message, steps) pair if the file cannot be used, otherwise None
def check_upload_is_usable(filepath, filename):
    try:
        file_is_empty = os.path.getsize(filepath) == 0
    except FileNotFoundError:
        logging.error(f"{filename} was missing when we tried to check its size.")
        return "We could no longer find the uploaded file.", [
            "This is usually a temporary hiccup.",
            "Upload the presentation again.",
        ]

    if file_is_empty:
        logging.error(f"{filename} is an empty file (0 bytes).")
        return "This file is empty (0 bytes), so there's nothing to open.", [
            "Check the file on your computer to make sure it actually has content.",
            "If it's empty there too, re-export or re-save the presentation from PowerPoint.",
            "Upload the new file.",
        ]

    return None


# the file might still be open in PowerPoint or being scanned by antivirus,
# so give it a couple of tries before giving up
def open_presentation(filepath, filename):
    open_error = None

    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            return Presentation(filepath), None
        except PermissionError as e:
            open_error = e
            logging.warning(f"{filename} looks like it's in use (attempt {attempt}/{MAX_ATTEMPTS}). Retrying...")
            time.sleep(RETRY_DELAY_SECONDS)
        except Exception as e:
            return None, e

    return None, open_error


# same idea as opening: the upload might still be briefly locked
def move_original_to_output(directory, new_directory, filename):
    original_upload_path = os.path.join(directory, filename)
    moved_original_path = os.path.join(new_directory, filename)
    move_error = None

    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            os.rename(original_upload_path, moved_original_path)
            return None
        except PermissionError as e:
            move_error = e
            logging.warning(f"{filename} looks like it's in use while finishing up (attempt {attempt}/{MAX_ATTEMPTS}). Retrying...")
            time.sleep(RETRY_DELAY_SECONDS)
        except OSError as e:
            return e

    return move_error


def process_one_file(filename, directory, new_directory, progress_callback=None):
    os.makedirs(new_directory, exist_ok=True)
    filepath = os.path.join(directory, filename)
    logging.info(f"Starting: {filename}")

    problem = check_upload_is_usable(filepath, filename)

    if problem is not None:
        message, steps = problem
        return False, message, steps

    # an old .ppt has to be converted before python-pptx can read it. The converted
    # copy keeps the same name, so every output file below is named the same either way.
    converted_filepath = None

    if filename.lower().endswith(".ppt"):
        converted_filepath, convert_message, convert_steps = convert_ppt_to_pptx(filepath, directory)

        if converted_filepath is None:
            return False, convert_message, convert_steps

        filepath = converted_filepath

    prs, open_error = open_presentation(filepath, filename)

    if prs is None:
        logging.error(f"Could not open {filename}. Error: {open_error}")
        message, steps = get_open_error_details(open_error)
        return False, message, steps

    alt_text_output_file = os.path.join(new_directory, os.path.splitext(filename)[0] + "_alt_text")

    try:
        records, problems, reading_level = fix_titles_and_describe_shapes(prs, filename, progress_callback)

        with open(alt_text_output_file, "w", encoding="utf-8") as f:
            json.dump({
                "file": filename,
                "shapes": records,
                "problems": problems,
                "reading_level": reading_level,
            }, f, indent=2)
    except Exception as e:
        logging.error(f"Something went wrong processing shapes in {filename}. Error: {e}")
        return False, "We opened your file, but something went wrong while checking your slides for accessibility issues.", [
            "This can happen with unusual slide layouts or embedded objects.",
            "Try saving a fresh copy in PowerPoint and uploading it again.",
        ]

    new_filepath = os.path.join(new_directory, os.path.splitext(filename)[0] + "_updated.pptx")

    try:
        prs.save(new_filepath)
    except Exception as e:
        logging.error(f"Could not save updated file for {filename}. Error: {e}")
        return False, "We updated your slides, but something went wrong saving the new file.", [
            "This is often temporary. Please try submitting it again.",
        ]

    # if we converted a .ppt, the converted copy was only ever a working file
    if converted_filepath is not None:
        try:
            os.remove(converted_filepath)
        except OSError as e:
            logging.warning(f"Could not tidy up the converted copy of {filename}. Error: {e}")

    move_error = move_original_to_output(directory, new_directory, filename)

    if move_error is not None:
        logging.error(f"Could not move original file {filename} to output folder. Error: {move_error}")
        return False, "Your file was processed, but something went wrong while finishing up.", [
            "Please try submitting it again.",
        ]

    logging.info(f"Finished: {filename}")
    return True, None, None


# link text like "click here" tells a screen reader user nothing, because they often
# jump through a page link by link, away from the words around it
VAGUE_LINK_WORDS = ("click here", "here", "read more", "more", "link", "this", "this link", "click")


def get_slide_title(slide):
    try:
        placeholder = slide.shapes.title
    except Exception:
        placeholder = None

    if placeholder is not None and placeholder.has_text_frame and placeholder.text.strip():
        return placeholder.text.strip()

    # some decks do not use the title placeholder, but do name the shape "Title"
    for shape in slide.shapes:
        if "Title" in shape.name and shape.has_text_frame and shape.text.strip():
            return shape.text.strip()

    return None


def find_vague_links(slide):
    found = []

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if run.hyperlink is None or not run.hyperlink.address:
                    continue

                words = run.text.strip().lower().strip(".,:;!?")

                if words in VAGUE_LINK_WORDS:
                    found.append(run.text.strip())

    return found


def find_tables_without_a_header(slide):
    found = []

    for shape in slide.shapes:
        if shape.has_table and not shape.table.first_row:
            found.append(shape.name)

    return found


# shapes sitting completely outside the slide are invisible to everyone looking at it,
# but a screen reader still reads them out
def find_offslide_shapes(slide, slide_width, slide_height):
    found = []

    for shape in slide.shapes:
        if shape.left is None or shape.top is None or shape.width is None or shape.height is None:
            continue

        right = shape.left + shape.width
        bottom = shape.top + shape.height

        completely_outside = (
            right <= 0
            or bottom <= 0
            or shape.left >= slide_width
            or shape.top >= slide_height
        )

        if completely_outside:
            found.append(shape.name)

    return found


def shapes_with_a_position(slide):
    found = []

    for shape in slide.shapes:
        if shape.top is None or shape.left is None or shape.width is None or shape.height is None:
            continue

        if is_picture(shape) or (shape.has_text_frame and shape.text.strip()):
            found.append(shape)

    return found


# shapes are read out in the order they sit in the file, which is not always the order
# they are laid out in. We only report a pair we can actually defend: one shape read
# after another that it sits clearly above, or clearly to the left of. Two shapes in
# separate columns are left alone, because reading one column and then the other is a
# normal way to lay a slide out and we would only be guessing.
def find_reading_order_problem(slide):
    shapes = shapes_with_a_position(slide)
    slack = 457200  # half an inch in the units python-pptx uses

    for later in range(1, len(shapes)):
        for earlier in range(later):
            first = shapes[earlier]
            second = shapes[later]

            shares_a_column = (
                second.left < first.left + first.width
                and first.left < second.left + second.width
            )
            shares_a_row = (
                second.top < first.top + first.height
                and first.top < second.top + second.height
            )

            if shares_a_column and second.top + second.height <= first.top - slack:
                return f"{second.name} is read after {first.name}, but it sits above it on the slide."

            if shares_a_row and second.left + second.width <= first.left - slack:
                return f"{second.name} is read after {first.name}, but it sits to the left of it."

    return None


def find_table_cell_problems(slide):
    found = []

    for shape in slide.shapes:
        if not shape.has_table:
            continue

        merged = 0
        blank = 0

        for row in shape.table.rows:
            for cell in row.cells:
                if cell.is_merge_origin or cell.is_spanned:
                    merged += 1

                if not cell.text.strip():
                    blank += 1

        if merged:
            found.append(f"The table {shape.name} has merged cells, which screen readers read out of order.")

        if blank:
            found.append(f"The table {shape.name} has {blank} empty cell(s). Putting N/A in them makes the table easier to follow.")

    return found


def slide_has_notes(slide):
    if not slide.has_notes_slide:
        return False

    return bool(slide.notes_slide.notes_text_frame.text.strip())


# looks for the accessibility problems we can spot but should not quietly change,
# because only the person who wrote the slides knows what the right wording is
def find_slide_problems(slide, slide_number, title, seen_titles, slide_width=None, slide_height=None):
    problems = []

    if title is None:
        problems.append({
            "slide": slide_number,
            "kind": "no title",
            "detail": "This slide has no title, so a screen reader cannot announce what it is about.",
        })
    elif title.lower() in seen_titles:
        problems.append({
            "slide": slide_number,
            "kind": "repeated title",
            "detail": f"Another slide is also called \"{title}\", which makes them hard to tell apart.",
        })

    for link_text in find_vague_links(slide):
        problems.append({
            "slide": slide_number,
            "kind": "unclear link",
            "detail": f"The link says \"{link_text}\", which does not say where it goes.",
        })

    for table_name in find_tables_without_a_header(slide):
        problems.append({
            "slide": slide_number,
            "kind": "table without a header row",
            "detail": f"The table {table_name} has no header row, so its columns are not announced.",
        })

    for detail in find_table_cell_problems(slide):
        problems.append({
            "slide": slide_number,
            "kind": "table cells",
            "detail": detail,
        })

    if slide_width is not None and slide_height is not None:
        for shape_name in find_offslide_shapes(slide, slide_width, slide_height):
            problems.append({
                "slide": slide_number,
                "kind": "off the slide",
                "detail": f"{shape_name} sits outside the slide, so nobody sees it but a screen reader still reads it.",
            })

    order_problem = find_reading_order_problem(slide)

    if order_problem is not None:
        problems.append({
            "slide": slide_number,
            "kind": "reading order",
            "detail": order_problem,
        })

    if slide_has_notes(slide):
        problems.append({
            "slide": slide_number,
            "kind": "speaker notes",
            "detail": "This slide has speaker notes. They are not always read aloud, so put anything important on the slide itself.",
        })

    return problems


VOWELS = "aeiouy"


def count_syllables(word):
    word = word.lower().strip(".,:;!?\"'()")

    if not word:
        return 0

    syllables = 0
    previous_was_vowel = False

    for letter in word:
        is_vowel = letter in VOWELS

        if is_vowel and not previous_was_vowel:
            syllables += 1

        previous_was_vowel = is_vowel

    # a trailing "e" is usually silent, as in "make"
    if word.endswith("e") and syllables > 1:
        syllables -= 1

    if syllables == 0:
        syllables = 1

    return syllables


# the Flesch Kincaid grade level, which says roughly what school year someone would
# need to read the text comfortably. It is only a rough guide on slides, because
# slides are full of short fragments rather than full sentences.
def measure_reading_level(text):
    sentences = 0

    for mark in ".!?":
        sentences += text.count(mark)

    words = text.split()

    if len(words) < 30 or sentences == 0:
        return None

    syllables = 0

    for word in words:
        syllables += count_syllables(word)

    words_per_sentence = len(words) / sentences
    syllables_per_word = syllables / len(words)
    grade = (0.39 * words_per_sentence) + (11.8 * syllables_per_word) - 15.59

    return round(grade, 1)
def get_shape_type(shape):
    """Return the shape's MSO type, or None if python-pptx cannot resolve it.

    python-pptx raises for shape types it does not model. One exotic shape should
    not take down the whole deck, so callers treat None as "not a picture".
    """
    try:
        return shape.shape_type
    except (NotImplementedError, ValueError):
        return None


def is_picture(shape):
    return get_shape_type(shape) == MSO_SHAPE_TYPE.PICTURE


def is_title(shape):
    """True if the shape looks like a slide title.

    The placeholder type is checked first because it is authoritative, and the
    shape name is a fallback so plain text boxes acting as titles are caught too.
    The name check is case insensitive: matching only "Title" misses a real title
    placeholder named "title 1". Subtitles are excluded, since "subtitle"
    contains "title" but is not one.
    """
    try:
        if shape.is_placeholder and shape.placeholder_format.type in TITLE_PLACEHOLDERS:
            return True
    except (AttributeError, KeyError, ValueError):
        pass

    name = shape.name.lower()

    return "title" in name and "subtitle" not in name


def is_placeholder_alt_text(alt_text, shape_name):
    """True if alt text exists but is not a real description.

    Covers PowerPoint's auto-filled source filename, and alt text that only
    repeats a shape name PowerPoint generated itself.

    Alt text matching a name the author chose is left alone. Someone who renames
    a shape to "Water cycle diagram" and writes the same description meant it,
    and replacing real alt text is worse than leaving a thin description alone.
    """
    if not alt_text:
        return False

    candidate = alt_text.strip().lower()
    name = shape_name.strip().lower()

    if candidate.endswith(IMAGE_EXTENSIONS):
        return True

    return candidate == name and AUTO_SHAPE_NAME.match(name) is not None


def needs_caption(shape):
    """True if this is a picture with no alt text a screen reader could use."""
    if not is_picture(shape):
        return False

    return looks_like_junk_alt_text(get_picture_alt_text(shape), shape.name)


def move_titles_to_front(slide):
    """Move title shapes to the front of the slide's reading order.

    Returns the number of shapes moved. The shape list is snapshotted before any
    mutation: re-reading slide.shapes[0] while reordering re-anchors on a shape
    that has already moved, which reverses the order of a slide with more than
    one title.
    """
    shapes = list(slide.shapes)
    titles = [shape for shape in shapes if is_title(shape)]

    if not titles:
        return 0

    # already in the right order, so nothing to do. This also makes a second run
    # over an already-processed deck a no-op.
    if shapes[:len(titles)] == titles:
        return 0

    # Anchor on the first non-title shape rather than shapes[0]. If shapes[0] is
    # itself a title, inserting the remaining titles before it puts them in
    # reverse order. The early return above guarantees a non-title shape exists.
    anchor = next(shape for shape in shapes if not is_title(shape))._element

    for title in titles:
        anchor.addprevious(title._element)

    return len(titles)


def count_images_needing_captions(prs):
    total = 0

    for slide in prs.slides:
        for shape in slide.shapes:
            if needs_caption(shape):
                total += 1

    return total


# fixes the reading order and returns one record per shape, which becomes the report
def fix_titles_and_describe_shapes(prs, filename, progress_callback=None):
    total_to_caption = count_images_needing_captions(prs)
    captions_done = 0
    records = []
    problems = []
    seen_titles = set()
    all_words = []

    if progress_callback is not None:
        progress_callback(captions_done, total_to_caption)

    for slide_number, slide in enumerate(prs.slides, start=1):

        if len(slide.shapes) == 0:
            logging.warning(f"{filename}: slide {slide_number} has no shapes, skipping title fix.")
            problems.append({
                "slide": slide_number,
                "kind": "empty slide",
                "detail": "This slide has nothing on it.",
            })
            continue

        # the whole point of this tool is to put the title first, so do that before
        # checking the order, otherwise we report a problem we are about to fix
        moved = move_titles_to_front(slide)

        if moved:
            logging.info(f"{filename}: slide {slide_number} moved {moved} title shape(s) to the front.")

        title = get_slide_title(slide)
        problems.extend(find_slide_problems(
            slide, slide_number, title, seen_titles, prs.slide_width, prs.slide_height
        ))

        if title is not None:
            seen_titles.add(title.lower())

        slide_text = get_slide_text(slide)
        all_words.append(get_all_slide_text(slide))

        # read the shapes back after reordering, so the report lists them in the
        # order a screen reader will actually announce them
        for shape in list(slide.shapes):
            needed_caption = needs_caption(shape)

            records.append(describe_shape(shape, slide_number, slide_text))

            if needed_caption:
                captions_done += 1

                if progress_callback is not None:
                    progress_callback(captions_done, total_to_caption)

    reading_level = measure_reading_level(" ".join(all_words))

    return records, problems, reading_level


# PowerPoint and Google Slides often fill alt text in by themselves, usually with the
# title of the web page the picture was copied from. That tells a screen reader user
# nothing about the picture, so we treat it as missing and write a real description.
SITE_NAMES_IN_TITLES = ("wikipedia", "fandom", "howstuffworks", "magnum photos", "medium")
IMAGE_FILE_ENDINGS = (".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp")


def looks_like_junk_alt_text(alt_text, shape_name):
    if not alt_text or not alt_text.strip():
        return True

    text = alt_text.strip()
    lowered = text.lower()

    # Google Slides exports leave their own shape ids behind
    if lowered.startswith("image google shape"):
        return True

    # a placeholder this tool wrote itself when a description could not be made
    if text == "Image " + shape_name:
        return True

    # the source filename PowerPoint fills in, or an echo of a name it made up
    if is_placeholder_alt_text(text, shape_name):
        return True

    if lowered.startswith("http://") or lowered.startswith("https://"):
        return True

    if lowered.endswith(IMAGE_FILE_ENDINGS):
        return True

    # search result titles get cut short
    if text.endswith("...") or text.endswith("…"):
        return True

    # page titles usually carry the site name after a bar
    if " | " in text:
        return True

    for site in SITE_NAMES_IN_TITLES:
        if site in lowered:
            return True

    return False


def get_picture_alt_text(shape):
    # python-pptx doesn't expose a working .alt_text property in this version,
    # so we have to read the real "descr" attribute straight from the XML
    return shape._element.nvPicPr.cNvPr.get("descr")


def set_picture_alt_text(shape, alt_text):
    shape._element.nvPicPr.cNvPr.set("descr", alt_text)


# builds one line of the report, and writes a description into the file if the
# picture did not already have one
def describe_shape(shape, slide_number, slide_text=""):
    if is_picture(shape):
        existing_alt_text = get_picture_alt_text(shape)

        if not looks_like_junk_alt_text(existing_alt_text, shape.name):
            return {
                "slide": slide_number,
                "name": shape.name,
                "kind": "image",
                "description": existing_alt_text,
                "we_added_it": False,
            }

        if existing_alt_text and existing_alt_text.strip():
            logging.info(f"Replacing unhelpful alt text on {shape.name}: {existing_alt_text[:60]!r}")

        try:
            alt_text = generate_image_caption(shape, slide_text)
        except Exception as e:
            logging.warning(f"Could not generate a caption for {shape.name}. Falling back to a placeholder. Error: {e}")
            alt_text = f"Image {shape.name}"

        set_picture_alt_text(shape, alt_text)

        return {
            "slide": slide_number,
            "name": shape.name,
            "kind": "image",
            "description": alt_text,
            "we_added_it": True,
        }

    if shape.has_text_frame:
        return {
            "slide": slide_number,
            "name": shape.name,
            "kind": "text",
            "description": shape.text,
            "we_added_it": False,
        }

    return {
        "slide": slide_number,
        "name": shape.name,
        "kind": "other",
        "description": None,
        "we_added_it": False,
    }


if __name__ == "__main__":
    directory_path = "pptx_input"
    new_directory = "pptx_output"
    accessibility_processor(directory_path, new_directory)
