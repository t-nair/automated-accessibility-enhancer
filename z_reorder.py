import os
import io
import re
import time
import string
import shutil
import logging
import itertools
import colorsys
import subprocess
import torch
from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from PIL import Image
from title_detection import guess_title, _ph_chain, DEFAULT_PT, FURNITURE_PH
from transformers import Qwen2VLForConditionalGeneration, AutoProcessor

# file for the record, console so a failure is visible the moment it happens
logging.basicConfig(
    handlers=[logging.FileHandler("pipeline.log", encoding="utf-8"), logging.StreamHandler()],
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

# PowerPoint keeps the language as a lang attribute such as "en-US", on text runs and on
# the presentation's default text style. This is the shape of a usable value.
LANGUAGE_TAG = re.compile(r"^[a-z]{2,3}(-[a-z0-9]{2,8})*$", re.IGNORECASE)

# WCAG sets no minimum point size, but text this small fails 1.4.4 (Resize Text) in
# practice and cannot be read when projected. Set MIN_FONT_PT to use a different cutoff.
MIN_FONT_PT = float(os.environ.get("MIN_FONT_PT", "9"))

# WCAG 1.4.3: 4.5:1 for normal text, 3:1 for large text (18 pt, or 14 pt bold)
MIN_CONTRAST = 4.5
MIN_CONTRAST_LARGE = 3.0
LARGE_PT = 18
LARGE_BOLD_PT = 14

# WCAG 2.4.4: link text that says nothing about where the link goes. Compared after
# lowercasing and trimming punctuation, so "Click here!" matches "click here".
VAGUE_LINK_TEXT = {
    "", "click", "click here", "click this", "here", "this", "this link", "link", "link here",
    "go", "go here", "more", "read more", "learn more", "more info", "more information",
    "info", "details", "see more", "website", "this website", "page", "this page",
}
BARE_URL = re.compile(r"^(https?://|www\.)\S+$", re.IGNORECASE)

# a table cell holding only a number, e.g. "42", "-3.5", "$1,200" or "15%", is data, not a heading
NUMBER = re.compile(r"^[-+$€£]?\d[\d,.]*%?$")

# what a fill turns out to be when it isn't one solid colour
CLEAR = "clear"
UNKNOWN = "unknown"
WHITE = (1.0, 1.0, 1.0)

CAPTION_MODEL_NAME = "Qwen/Qwen2-VL-2B-Instruct"

# the model is told once, here, what kind of description we want. Faculty never see
# or type this, it is just how we ask for a caption that works as alt text.
CAPTION_PROMPT = (
    "Write alt text for this image in one or two short sentences. "
    "Describe what is visible, and if the image contains legible text, include it. "
    "Only state what you can clearly see. Describe things plainly instead of naming them if you are unsure. "
    "Do not add mood, interpretation, or invented detail."
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
        logging.info("Captioning model is loaded and ready.")

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
    logging.info(f"Captioning {shape.name} ({image.width}x{image.height}px, {len(image_bytes) // 1024} KB).")

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

    caption = tidy_caption(caption)
    logging.info(f"Caption generated in {seconds_taken:.2f} seconds on {CAPTION_DEVICE}: {caption!r}")

    # an empty caption would be written into the deck as blank alt text, so fail loudly instead
    if not caption:
        raise ValueError("the model returned an empty caption")

    return caption


def accessibility_processor(directory, new_directory, move_originals=False):
    files = os.listdir(directory)
    pptx_files = []

    for filename in files:
        if filename.endswith(".pptx"):
            pptx_files.append(filename)

    logging.info(f"Found {len(pptx_files)} pptx file(s) in {directory}")

    if len(pptx_files) == 0:
        logging.warning("No .pptx files found. Nothing to process.")
        return

    failed = []

    for filename in pptx_files:
        was_successful, message, _ = process_one_file(filename, directory, new_directory, move_originals=move_originals)

        if not was_successful:
            logging.error(f"{filename} failed: {message}")
            failed.append(filename)

    logging.info(
        f"Run summary: {len(pptx_files) - len(failed)}/{len(pptx_files)} file(s) processed into "
        f"{new_directory}. Failed: {', '.join(failed) or 'none'}")


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
        logging.error(
            f"Conversion of {filepath} did not produce a file (exit code {result.returncode}). "
            f"LibreOffice said: {result.stderr[:300]!r} {result.stdout[:300]!r}")
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
            logging.exception(f"{filename} could not be parsed as a presentation.")
            return None, e

    logging.error(f"{filename} is still locked after {MAX_ATTEMPTS} attempts.")
    return None, open_error


# same idea as opening: the upload might still be briefly locked
def move_original_to_output(directory, new_directory, filename):
    original_upload_path = os.path.join(directory, filename)
    moved_original_path = os.path.join(new_directory, filename)
    move_error = None
    linked = False

    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            # link then remove: os.link refuses an existing file in one step, while
            # os.rename would silently replace it on Linux and macOS
            if not linked:
                try:
                    os.link(original_upload_path, moved_original_path)
                except FileExistsError:
                    # an earlier attempt that could not remove the original leaves its
                    # link behind, which is the same file, so carry on from there
                    if not os.path.samefile(original_upload_path, moved_original_path):
                        raise

                linked = True

            os.remove(original_upload_path)
            return None
        except PermissionError as e:
            move_error = e
            logging.warning(f"{filename} looks like it's in use while finishing up (attempt {attempt}/{MAX_ATTEMPTS}). Retrying...")
            time.sleep(RETRY_DELAY_SECONDS)
        except OSError as e:
            return e

    return move_error


# the original is left where it is unless move_originals is set
def process_one_file(filename, directory, new_directory, progress_callback=None, language=None, move_originals=False):
    # language is optional. When given, it is set across the whole deck, e.g. "en-US".
    if language is not None and not is_language_tag(language):
        logging.error(f"{filename} not processed: {language!r} is not a language code.")
        return False, f"{language!r} is not a valid language code.", ["Use a code such as en-US or fr-FR."]

    os.makedirs(new_directory, exist_ok=True)
    filepath = os.path.join(directory, filename)
    started_at = time.perf_counter()
    logging.info(f"Starting: {filename}")

    problem = check_upload_is_usable(filepath, filename)

    if problem is not None:
        message, steps = problem
        logging.error(f"{filename} rejected before processing: {message}")
        return False, message, steps

    # an old .ppt has to be converted before python-pptx can read it. The converted
    # copy keeps the same name, so every output file below is named the same either way.
    converted_filepath = None

    if filename.lower().endswith(".ppt"):
        converted_filepath, convert_message, convert_steps = convert_ppt_to_pptx(filepath, directory)

        if converted_filepath is None:
            logging.error(f"{filename} could not be converted: {convert_message}")
            return False, convert_message, convert_steps

        filepath = converted_filepath

    prs, open_error = open_presentation(filepath, filename)

    if prs is None:
        logging.error(f"Could not open {filename}. Error: {open_error!r}")
        message, steps = get_open_error_details(open_error)
        return False, message, steps

    logging.info(f"{filename}: opened {len(prs.slides)} slide(s) in {time.perf_counter() - started_at:.2f}s, checking titles and alt text.")
    alt_text_output_file = os.path.join(new_directory, os.path.splitext(filename)[0] + "_alt_text")

    try:
        # the details go to memory first, so the score can be written above them
        details = io.StringIO()
        counts = fix_titles_and_log_alt_text(prs, details, filename, progress_callback)
        issues = counts["issues"] + report_and_set_language(prs, details, filename, language)

        with open(alt_text_output_file, "w", encoding="utf-8") as f:
            f.write(severity_summary(issues))
            f.write(details.getvalue())
    except Exception as e:
        logging.exception(f"Something went wrong processing shapes in {filename}. Error: {e}")
        return False, "We opened your file, but something went wrong while checking your slides for accessibility issues.", [
            "This can happen with unusual slide layouts or embedded objects.",
            "Try saving a fresh copy in PowerPoint and uploading it again.",
        ]

    new_filepath = os.path.join(new_directory, os.path.splitext(filename)[0] + "_updated.pptx")

    try:
        prs.save(new_filepath)
    except Exception as e:
        logging.exception(f"Could not save updated file for {filename}. Error: {e}")
        return False, "We updated your slides, but something went wrong saving the new file.", [
            "This is often temporary. Please try submitting it again.",
        ]

    # if we converted a .ppt, the converted copy was only ever a working file
    if converted_filepath is not None:
        try:
            os.remove(converted_filepath)
        except OSError as e:
            logging.warning(f"Could not tidy up the converted copy of {filename}. Error: {e}")

    move_error = move_original_to_output(directory, new_directory, filename) if move_originals else None

    if move_error is not None:
        logging.error(f"Could not move original file {filename} to output folder. Error: {move_error}")
        return False, "Your file was processed, but something went wrong while finishing up.", [
            "Please try submitting it again.",
        ]

    logging.info(
        f"Summary: {filename} | {counts['slides']} slide(s), {counts['titles_moved']} title(s) moved, "
        f"{counts['captioned']} image(s) captioned, {counts['small_text']} shape(s) with small text, "
        f"{counts['low_contrast']} shape(s) with low contrast | report: {alt_text_output_file} | "
        f"updated deck: {new_filepath} | {time.perf_counter() - started_at:.2f}s")
    return True, None, None


def get_shape_type(shape):
    """Return the shape's MSO type, or None if python-pptx cannot resolve it.

    python-pptx raises for shape types it does not model. One exotic shape should
    not take down the whole deck, so callers treat None as "not a picture".
    """
    try:
        return shape.shape_type
    except (NotImplementedError, ValueError) as e:
        logging.debug(f"Shape type unavailable for {getattr(shape, 'name', '?')}: {e!r}")
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
    except (AttributeError, KeyError, ValueError) as e:
        logging.debug(f"Could not read placeholder type of {shape.name}: {e!r}")

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

    alt_text = get_picture_alt_text(shape)

    return not alt_text or is_placeholder_alt_text(alt_text, shape.name)


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
        # no title placeholder and no shape named "title", so the author probably typed the
        # title into a text box or a body placeholder. Ask the model which shape reads like one.
        guess = guess_title(slide)
        guessed_name = guess.name if guess is not None else "no title"
        logging.info(f"No title placeholder; model guessed {guessed_name!r}.")
        titles = [shape for shape in shapes if guess is not None and shape._element is guess._element]

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


def has_title(slide):
    """True if the slide has a title shape with actual text in it.

    An empty title placeholder ("Click to add title") counts as no title: it
    announces nothing to a screen reader, and PowerPoint's checker flags it too.
    """
    for shape in slide.shapes:
        if is_title(shape) and shape.has_text_frame and shape.text_frame.text.strip():
            return True

    return False


def is_language_tag(language):
    return LANGUAGE_TAG.match(language) is not None


def get_default_language(prs):
    """The language on the presentation's default text style, or None if it has none."""
    default = prs.part._element.find(
        qn("p:defaultTextStyle") + "/" + qn("a:defPPr") + "/" + qn("a:defRPr"))

    return default.get("lang") if default is not None else None


def get_run_languages(prs):
    """Map each lang value set on a slide's text runs to the slide numbers that use it.

    Runs with no lang attribute inherit the default, so they are left out.
    """
    found = {}

    for slide_number, slide in enumerate(prs.slides, start=1):
        for run_props in slide._element.iter(qn("a:rPr")):
            language = run_props.get("lang")

            if language and slide_number not in found.setdefault(language, []):
                found[language].append(slide_number)

    return found

def get_shapes(shapes):
    """Yield every shape, with groups opened up into the shapes inside them."""
    for shape in shapes:
        if get_shape_type(shape) == MSO_SHAPE_TYPE.GROUP:
            yield from get_shapes(shape.shapes)
        else:
            yield shape


def get_text_frames(shapes):
    """Yield (shape, text_frame) for every text box, placeholder and table cell, inside groups too."""
    for shape in get_shapes(shapes):
        if shape.has_text_frame:
            yield shape, shape.text_frame
        elif shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    yield shape, cell.text_frame


def get_style_chain(slide, shape, text_frame):
    """Elements that can set a size for this text, nearest first: the shape's own
    list style, then the layout and master placeholders, then the master's text styles."""
    chain = [text_frame._txBody.find(qn("a:lstStyle"))]
    style_name = "otherStyle"

    if shape.is_placeholder:
        for element in _ph_chain(shape):
            if element.tag != qn("p:txStyles"):
                chain.append(element.find(".//" + qn("a:lstStyle")))

        if shape.placeholder_format.type in TITLE_PLACEHOLDERS:
            style_name = "titleStyle"
        elif shape.placeholder_format.type not in FURNITURE_PH:
            style_name = "bodyStyle"

    master_styles = slide.slide_layout.slide_master._element.find(qn("p:txStyles"))

    if master_styles is not None:
        chain.append(master_styles.find(qn("p:" + style_name)))

    return chain


def get_run_pt(run, level, chain, scale):
    """Size in points the run will render at, after any autofit shrinking."""
    if run.font.size is not None:
        return run.font.size.pt * scale

    for style in chain:
        if style is None:
            continue

        default_run = style.find(qn(f"a:lvl{level + 1}pPr") + "/" + qn("a:defRPr"))

        if default_run is not None and default_run.get("sz"):
            return int(default_run.get("sz")) / 100 * scale

    # ponytail: ignores the deck's defaultTextStyle, which is 18 pt in practice
    return DEFAULT_PT * scale


def find_small_text(slide, min_pt=MIN_FONT_PT):
    """Return (shape name, smallest size in pt, text) for each shape with a run below min_pt."""
    found = []

    for shape, text_frame in get_text_frames(slide.shapes):
        chain = get_style_chain(slide, shape, text_frame)
        fit = text_frame._txBody.find(qn("a:bodyPr") + "/" + qn("a:normAutofit"))
        scale = int(fit.get("fontScale")) / 100000 if fit is not None and fit.get("fontScale") else 1.0
        sizes = []
        words = []

        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                pt = get_run_pt(run, paragraph.level, chain, scale)

                if run.text.strip() and pt < min_pt:
                    sizes.append(pt)
                    words.append(run.text.strip())

        if sizes:
            found.append((shape.name, min(sizes), " ".join(words)))

    return found


def get_theme(slide):
    """The master's theme and colour map, which scheme colours like tx1 and accent1 resolve through."""
    master = slide.slide_layout.slide_master
    theme = etree.fromstring(master.part.part_related_by(RT.THEME).blob)

    return theme, master._element.find(qn("p:clrMap"))


def scheme_rgb(name, theme, clr_map):
    # ponytail: ignores a clrMapOvr on the slide or layout, which few decks use
    name = clr_map.get(name, name) if clr_map is not None else name
    slot = theme.find(".//" + qn("a:clrScheme") + "/" + qn("a:" + name))

    return color_to_rgb(slot[0], theme, clr_map) if slot is not None and len(slot) else None


def color_to_rgb(color, theme, clr_map):
    """(r, g, b) from 0 to 1 for a colour element, or None if it can't be worked out."""
    tag = etree.QName(color).localname

    if tag == "schemeClr":
        rgb = scheme_rgb(color.get("val"), theme, clr_map)
    elif tag in ("srgbClr", "sysClr"):
        hex_value = color.get("val") if tag == "srgbClr" else color.get("lastClr")
        rgb = tuple(int(hex_value[i:i + 2], 16) / 255 for i in (0, 2, 4)) if hex_value else None
    else:
        return None

    if rgb is None:
        return None

    lum_mod, lum_off = 1.0, 0.0

    for modifier in color:
        name = etree.QName(modifier).localname

        if name == "lumMod":
            lum_mod = int(modifier.get("val")) / 100000
        elif name == "lumOff":
            lum_off = int(modifier.get("val")) / 100000
        else:
            # ponytail: tint, shade, alpha and the rest are not applied, so the colour is unknown
            return None

    if lum_mod != 1.0 or lum_off:
        hue, lightness, saturation = colorsys.rgb_to_hls(*rgb)
        rgb = colorsys.hls_to_rgb(hue, max(0.0, min(1.0, lightness * lum_mod + lum_off)), saturation)

    return rgb


def get_fill(props, theme, clr_map):
    """The fill set directly on props: an RGB tuple, CLEAR for no fill, UNKNOWN for a fill
    that isn't one solid colour, or None if props sets no fill at all."""
    if props is None:
        return None

    for child in props:
        name = etree.QName(child).localname

        if name == "solidFill":
            return (color_to_rgb(child[0], theme, clr_map) if len(child) else None) or UNKNOWN
        if name == "noFill":
            return CLEAR
        if name in ("gradFill", "blipFill", "pattFill", "grpFill"):
            return UNKNOWN

    return None


def get_style_ref_fill(ref, theme, clr_map):
    """Fill for a fillRef or bgRef, which point into the theme's fill styles by index."""
    index = int(ref.get("idx", "0"))

    if index == 0:
        return CLEAR

    style_list = theme.find(".//" + qn("a:bgFillStyleLst" if index >= 1001 else "a:fillStyleLst"))
    styles = list(style_list) if style_list is not None else []
    position = index - 1001 if index >= 1001 else index - 1

    if position >= len(styles):
        return UNKNOWN

    style = styles[position]

    if etree.QName(style).localname != "solidFill" or not len(style) or not len(ref):
        return get_fill([style], theme, clr_map) or UNKNOWN

    # the style's phClr stands for the colour given on the reference itself
    color = ref[0] if style[0].get("val") == "phClr" and not len(style[0]) else style[0]

    return color_to_rgb(color, theme, clr_map) or UNKNOWN


def get_shape_fill(element, theme, clr_map):
    if element.tag == qn("p:pic"):
        return UNKNOWN

    fill = get_fill(element.find(qn("p:spPr")), theme, clr_map)

    if fill is None:
        ref = element.find(qn("p:style") + "/" + qn("a:fillRef"))
        fill = get_style_ref_fill(ref, theme, clr_map) if ref is not None else CLEAR

    return fill


def get_slide_background(slide, theme, clr_map):
    for source in (slide, slide.slide_layout, slide.slide_layout.slide_master):
        background = source._element.find(qn("p:cSld") + "/" + qn("p:bg"))

        if background is None:
            continue

        background_props = background.find(qn("p:bgPr"))

        if background_props is not None:
            fill = get_fill(background_props, theme, clr_map)
            return WHITE if fill == CLEAR else fill or UNKNOWN

        ref = background.find(qn("p:bgRef"))

        if ref is not None:
            return get_style_ref_fill(ref, theme, clr_map)

    return WHITE


def covers(behind, shape):
    edges = (behind.left, behind.top, behind.width, behind.height, shape.left, shape.top, shape.width, shape.height)

    return None not in edges and (
        behind.left <= shape.left and behind.top <= shape.top
        and behind.left + behind.width >= shape.left + shape.width
        and behind.top + behind.height >= shape.top + shape.height)


def get_background(slide, shape, text_frame, theme, clr_map):
    """Colour behind the text: the shape's own fill, else a filled shape under it, else the slide."""
    if shape.has_table:
        # ponytail: table styles are not read, so only cells with their own solid fill are checked
        fill = get_fill(text_frame._txBody.getparent().find(qn("a:tcPr")), theme, clr_map)
        return fill if fill not in (None, CLEAR) else UNKNOWN

    fill = get_shape_fill(shape._element, theme, clr_map)

    if fill != CLEAR:
        return fill

    # ponytail: only looks under top-level shapes, and only at shapes that fully cover it
    shapes = list(slide.shapes)

    if shape in shapes:
        for behind in reversed(shapes[:shapes.index(shape)]):
            if covers(behind, shape):
                fill = get_shape_fill(behind._element, theme, clr_map)

                if fill != CLEAR:
                    return fill

    return get_slide_background(slide, theme, clr_map)


def get_level_defaults(chain, level):
    """The defRPr for this paragraph level from each style in the chain, nearest first."""
    return [style.find(qn(f"a:lvl{level + 1}pPr") + "/" + qn("a:defRPr")) if style is not None else None
            for style in chain]


def get_run_color(run, defaults, shape, theme, clr_map):
    """Text colour as an RGB tuple, CLEAR for hidden text, or UNKNOWN."""
    font_ref = shape._element.find(qn("p:style") + "/" + qn("a:fontRef"))

    # the shape style's font colour beats the layout and master, but not the shape's own list style
    for props in [run._r.find(qn("a:rPr")), defaults[0], font_ref] + defaults[1:]:
        if props is not None and props is font_ref:
            return (color_to_rgb(font_ref[0], theme, clr_map) if len(font_ref) else None) or UNKNOWN

        fill = get_fill(props, theme, clr_map)

        if fill is not None:
            return fill

    return scheme_rgb("tx1", theme, clr_map) or UNKNOWN


def is_bold(run, defaults):
    if run.font.bold is not None:
        return run.font.bold

    return next((d.get("b") in ("1", "true") for d in defaults if d is not None and d.get("b") is not None), False)


def luminance(rgb):
    channels = [c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4 for c in rgb]

    return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2]


def contrast_ratio(first, second):
    lighter, darker = sorted((luminance(first) + 0.05, luminance(second) + 0.05), reverse=True)

    return lighter / darker


def find_low_contrast(slide):
    """Return (shape name, worst ratio, ratio needed, text) for each shape with text below WCAG 1.4.3.

    The ratios are None when the shape could not be checked, because the background or
    text colour is a gradient, picture, table style or a colour effect that isn't read.
    """
    found = []
    theme, clr_map = get_theme(slide)

    for shape, text_frame in get_text_frames(slide.shapes):
        if not text_frame.text.strip():
            continue

        chain = get_style_chain(slide, shape, text_frame)
        fit = text_frame._txBody.find(qn("a:bodyPr") + "/" + qn("a:normAutofit"))
        scale = int(fit.get("fontScale")) / 100000 if fit is not None and fit.get("fontScale") else 1.0
        background = get_background(slide, shape, text_frame, theme, clr_map)
        unchecked = background == UNKNOWN
        failures = []

        for paragraph in text_frame.paragraphs:
            defaults = get_level_defaults(chain, paragraph.level)

            for run in paragraph.runs:
                if unchecked or not run.text.strip():
                    continue

                color = get_run_color(run, defaults, shape, theme, clr_map)

                if color == UNKNOWN:
                    unchecked = True
                    continue

                if color == CLEAR:
                    continue

                pt = get_run_pt(run, paragraph.level, chain, scale)
                large = pt >= LARGE_PT or (pt >= LARGE_BOLD_PT and is_bold(run, defaults))
                needed = MIN_CONTRAST_LARGE if large else MIN_CONTRAST
                ratio = contrast_ratio(color, background)

                if ratio < needed:
                    failures.append((ratio, needed, run.text.strip()))

        if failures:
            ratio, needed, _ = min(failures)
            found.append((shape.name, ratio, needed, " ".join(text for _, _, text in failures)))
        elif unchecked:
            found.append((shape.name, None, None, text_frame.text.strip()))

    return found


def is_vague_link_text(text):
    text = " ".join(text.split())
    return text.lower().strip(string.punctuation) in VAGUE_LINK_TEXT or bool(BARE_URL.match(text))


def find_vague_links(slide):
    """Return (shape name, link text, address) for each link whose text doesn't say where it goes.

    Links are on text runs, or on a whole shape when clicking the shape follows it. A
    shape's link text is its own text, or its alt text when it has none, as a screen
    reader would announce it. Jumps to another slide have no address and are skipped.
    """
    found = []

    for shape, text_frame in get_text_frames(slide.shapes):
        for paragraph in text_frame.paragraphs:
            # one link can be split over several runs, e.g. by a bold word
            for address, runs in itertools.groupby(paragraph.runs, key=lambda run: run.hyperlink.address):
                text = "".join(run.text for run in runs)

                if address is not None and is_vague_link_text(text):
                    found.append((shape.name, text.strip(), address))

    for shape in get_shapes(slide.shapes):
        address = shape.click_action.hyperlink.address

        if address is None:
            continue

        text = shape.text_frame.text if shape.has_text_frame else ""
        text = text.strip() or (shape._element._nvXxPr.cNvPr.get("descr") or "").strip()

        if is_vague_link_text(text):
            found.append((shape.name, text, address))

    return found


def looks_like_header_row(row):
    """A first row reads as headings when every cell has text and none of it is a number."""
    texts = [cell.text.strip() for cell in row.cells]
    return all(texts) and not any(NUMBER.match(text) for text in texts)


def fix_table_headers(slide):
    """Turn on the header row flag where the first row reads as headings.

    Returns three lists of table names: headers turned on, tables still without a
    header row because the first row holds data, and tables with merged or split
    cells, which screen readers cannot line up with their headers.
    """
    fixed, no_header, merged = [], [], []

    for shape in get_shapes(slide.shapes):
        if not shape.has_table:
            continue

        table = shape.table

        if not table.first_row:
            if looks_like_header_row(table.rows[0]):
                table.first_row = True
                fixed.append(shape.name)
            else:
                no_header.append(shape.name)

        # a split cell is stored as its neighbours being merged, so one check finds both
        if any(cell.is_merge_origin for row in table.rows for cell in row.cells):
            merged.append(shape.name)

    return fixed, no_header, merged


def find_language_problems(prs):
    """List what is wrong with the deck's language, as short sentences for the report.

    WCAG 3.1.1 wants one language for the page, and 3.1.2 wants any passage in a
    different language marked. A deck mixing languages is reported rather than
    called wrong, because a marked French quotation is exactly what 3.1.2 asks for.
    The text itself is not checked, so a deck wrongly marked en-US is not caught.
    """
    problems = []
    default = get_default_language(prs)
    in_use = get_run_languages(prs)

    if not default:
        problems.append("the presentation has no default language, so screen readers have to guess")
    elif not is_language_tag(default):
        problems.append(f"the presentation's default language {default!r} is not a valid language code")

    for language, slides in in_use.items():
        if not is_language_tag(language):
            slide_list = ", ".join(str(n) for n in sorted(slides))
            problems.append(f"{language!r} is not a valid language code (slides {slide_list})")

    valid = {language: slides for language, slides in in_use.items() if is_language_tag(language)}
    families = {language.split("-")[0].lower() for language in valid}

    if default and is_language_tag(default):
        families.add(default.split("-")[0].lower())

    if len(families) > 1:
        # the language on the most slides is the main one, so only the others list slide numbers
        main = max(valid, key=lambda language: len(valid[language]))
        parts = []

        for language, slides in valid.items():
            if language == main:
                parts.append(f"{language} on {len(slides)} slide(s)")
            else:
                parts.append(f"{language} on slide(s) {', '.join(str(n) for n in sorted(slides))}")

        problems.append("more than one language is used, check each is intended: " + "; ".join(parts))

    return problems


def set_language(prs, language):
    """Set one language across the deck: its default text style, slides, layouts and masters.

    Only lang attributes that already exist are replaced. A run without one inherits
    the default, which is set here too, so every run ends up in the chosen language.
    """
    root = prs.part._element
    parts = [root] + [slide._element for slide in prs.slides]

    for master in prs.slide_masters:
        parts.append(master._element)
        parts.extend(layout._element for layout in master.slide_layouts)

    for part in parts:
        for element in part.iter(qn("a:rPr"), qn("a:endParaRPr"), qn("a:defRPr")):
            if element.get("lang") is not None:
                element.set("lang", language)

    style = root.find(qn("p:defaultTextStyle"))

    if style is None:
        style = OxmlElement("p:defaultTextStyle")
        later = root.find(qn("p:modifyVerifier"))
        later = later if later is not None else root.find(qn("p:extLst"))

        if later is not None:
            later.addprevious(style)
        else:
            root.append(style)

    default = style.find(qn("a:defPPr"))

    if default is None:
        default = OxmlElement("a:defPPr")
        style.insert(0, default)

    run_defaults = default.find(qn("a:defRPr"))

    if run_defaults is None:
        run_defaults = OxmlElement("a:defRPr")
        default.append(run_defaults)

    run_defaults.set("lang", language)


# every problem a check finds gets one of these, worst first, so people know what to fix first
SEVERITIES = ("critical", "serious", "moderate", "minor")

# ponytail: flat points off per problem, so a long deck scores lower than a short one with
# the same habits. Scale by slide count if people find that unfair.
SEVERITY_POINTS = {"critical": 10, "serious": 5, "moderate": 2, "minor": 1}


def severity_summary(issues):
    """The report's opening lines: a score out of 100 and how many problems of each severity.

    issues has one severity per problem found in the file as it was uploaded, including
    the ones that were fixed, so the score says how accessible the original was.
    """
    counts = {severity: issues.count(severity) for severity in SEVERITIES}
    score = max(0, 100 - sum(SEVERITY_POINTS[severity] * n for severity, n in counts.items()))
    by_severity = ", ".join(f"{n} {severity}" for severity, n in counts.items())

    return f"Accessibility score: {score}/100 (as uploaded)\nIssues by severity: {by_severity}\n"


def report_and_set_language(prs, f, filename, language=None):
    # what was found is written first, so the report describes the deck as it was uploaded
    issues = []

    for problem in find_language_problems(prs):
        # two languages may be intended, a missing or broken code never is
        severity = "minor" if problem.startswith("more than one language") else "serious"
        issues.append(severity)
        logging.warning(f"{filename}: language: {problem}")
        f.write(f" \n [{severity}] Language (WCAG 3.1.1, 3.1.2): {problem} \n")

    if language:
        set_language(prs, language)
        logging.info(f"{filename}: language set to {language}.")
        f.write(f" \n Language set to {language} across the whole presentation. \n")

    return issues


def count_images_needing_captions(prs):
    total = 0

    for slide in prs.slides:
        for shape in slide.shapes:
            if needs_caption(shape):
                total += 1

    return total


def fix_titles_and_log_alt_text(prs, f, filename, progress_callback=None):
    total_to_caption = count_images_needing_captions(prs)
    captions_done = 0
    titles_moved = 0
    logging.info(f"{filename}: {total_to_caption} image(s) need a caption.")

    if progress_callback is not None:
        progress_callback(captions_done, total_to_caption)

    untitled_slides = []
    reordered_slides = []
    small_text_lines = []
    low_contrast_lines = []
    unchecked_contrast_lines = []
    vague_link_lines = []
    header_fixed_lines = []
    no_header_lines = []
    merged_cell_lines = []
    issues = []

    for slide_number, slide in enumerate(prs.slides, start=1):

        logging.debug(f"{filename}: slide {slide_number} has {len(slide.shapes)} shape(s).")

        if not has_title(slide):
            untitled_slides.append(slide_number)

        for name, pt, text in find_small_text(slide):
            small_text_lines.append(f"Slide {slide_number}, Shape: {name} - {pt:g} pt: {text[:60]!r}")

        for name, ratio, needed, text in find_low_contrast(slide):
            if ratio is None:
                unchecked_contrast_lines.append(f"Slide {slide_number}, Shape: {name}: {text[:60]!r}")
            else:
                low_contrast_lines.append(
                    f"Slide {slide_number}, Shape: {name} - {ratio:.2f}:1, needs {needed:g}:1: {text[:60]!r}")

        for name, text, address in find_vague_links(slide):
            vague_link_lines.append(f"Slide {slide_number}, Shape: {name} - {text[:60]!r} links to {address}")

        fixed, no_header, merged = fix_table_headers(slide)
        header_fixed_lines += [f"Slide {slide_number}, Shape: {name}" for name in fixed]
        no_header_lines += [f"Slide {slide_number}, Shape: {name}" for name in no_header]
        merged_cell_lines += [f"Slide {slide_number}, Shape: {name}" for name in merged]

        if len(slide.shapes) == 0:
            logging.warning(f"{filename}: slide {slide_number} has no shapes, skipping title fix.")
            continue

        slide_text = get_slide_text(slide)

        moved = move_titles_to_front(slide)

        titles_moved += moved

        if moved:
            reordered_slides.append(slide_number)
            logging.info(f"{filename}: slide {slide_number} moved {moved} title shape(s) to the front.")

        # read the shapes back after reordering, so the report lists them in the
        # order a screen reader will actually announce them
        for shape in list(slide.shapes):
            needed_caption = needs_caption(shape)

            severity = write_alt_text_line(shape, f, slide_text)

            if severity:
                issues.append(severity)

            if needed_caption:
                captions_done += 1

                if progress_callback is not None:
                    progress_callback(captions_done, total_to_caption)

    # already fixed, but listed so the score matches what was found
    if reordered_slides:
        issues += ["moderate"] * len(reordered_slides)
        slide_list = ", ".join(str(n) for n in reordered_slides)
        f.write(f" \n [moderate] Title was not read first, now fixed (WCAG 1.3.2): slides {slide_list} \n")

    # report only: a generated title would be a guess, so a person adds the real one
    if untitled_slides:
        issues += ["serious"] * len(untitled_slides)
        slide_list = ", ".join(str(n) for n in untitled_slides)
        logging.warning(f"{filename}: slides with no title: {slide_list}")
        f.write(f" \n [serious] Slides with no title (WCAG 2.4.2): {slide_list} \n")

    # report only: raising sizes can break the layout, so a person decides
    if small_text_lines:
        issues += ["minor"] * len(small_text_lines)
        logging.warning(f"{filename}: {len(small_text_lines)} shape(s) with text below {MIN_FONT_PT:g} pt.")
        f.write(f" \n [minor] Text below {MIN_FONT_PT:g} pt (WCAG 1.4.4): \n")

        for line in small_text_lines:
            f.write(f" - {line} \n")

    # report only: a replacement colour is a design choice, so a person picks it
    if low_contrast_lines:
        issues += ["serious"] * len(low_contrast_lines)
        logging.warning(f"{filename}: {len(low_contrast_lines)} shape(s) with low contrast text.")
        f.write(" \n [serious] Text with too little contrast (WCAG 1.4.3): \n")

        for line in low_contrast_lines:
            f.write(f" - {line} \n")

    if unchecked_contrast_lines:
        f.write(" \n Contrast not checked, the background or text is not one solid colour (WCAG 1.4.3): \n")

        for line in unchecked_contrast_lines:
            f.write(f" - {line} \n")

    # report only: better wording depends on the sentence around the link, so a person writes it
    if vague_link_lines:
        issues += ["moderate"] * len(vague_link_lines)
        logging.warning(f"{filename}: {len(vague_link_lines)} link(s) with text that does not say where they go.")
        f.write(" \n [moderate] Link text does not say where the link goes (WCAG 2.4.4): \n")

        for line in vague_link_lines:
            f.write(f" - {line} \n")

    # already fixed, but listed so the score matches what was found
    if header_fixed_lines:
        issues += ["serious"] * len(header_fixed_lines)
        f.write(" \n [serious] Table had no header row, first row now marked as headers (WCAG 1.3.1): \n")

        for line in header_fixed_lines:
            f.write(f" - {line} \n")

    # report only: a first row of data would be read out as headings, so a person adds a real one
    if no_header_lines:
        issues += ["serious"] * len(no_header_lines)
        logging.warning(f"{filename}: {len(no_header_lines)} table(s) with no header row.")
        f.write(" \n [serious] Table has no header row, and its first row looks like data (WCAG 1.3.1): \n")

        for line in no_header_lines:
            f.write(f" - {line} \n")

    # report only: unmerging changes the table's layout, so a person decides
    if merged_cell_lines:
        issues += ["moderate"] * len(merged_cell_lines)
        logging.warning(f"{filename}: {len(merged_cell_lines)} table(s) with merged or split cells.")
        f.write(" \n [moderate] Table has merged or split cells (WCAG 1.3.1): \n")

        for line in merged_cell_lines:
            f.write(f" - {line} \n")

    logging.info(f"{filename}: alt text pass done, {captions_done}/{total_to_caption} image(s) captioned.")

    return {"slides": len(prs.slides), "titles_moved": titles_moved, "captioned": captions_done,
            "small_text": len(small_text_lines), "low_contrast": len(low_contrast_lines), "issues": issues}


def get_picture_alt_text(shape):
    # python-pptx doesn't expose a working .alt_text property in this version,
    # so we have to read the real "descr" attribute straight from the XML
    return shape._element.nvPicPr.cNvPr.get("descr")


def set_picture_alt_text(shape, alt_text):
    shape._element.nvPicPr.cNvPr.set("descr", alt_text)


# gives back the severity of the problem it found and fixed, or None if there was none
def write_alt_text_line(shape, f, slide_text=""):
    if is_picture(shape):
        existing_alt_text = get_picture_alt_text(shape)

        if existing_alt_text and not is_placeholder_alt_text(existing_alt_text, shape.name):
            f.write(f" \n Shape: {shape.name} \n - Alt Text: {existing_alt_text} \n")
        else:
            # no description at all hides the image, a filename at least says it is there
            severity = "serious" if existing_alt_text else "critical"

            if existing_alt_text:
                logging.info(
                    f"{shape.name}: replacing filename-style alt text "
                    f"{existing_alt_text!r} with a real description.")

            try:
                alt_text = generate_image_caption(shape, slide_text)
            except Exception as e:
                alt_text = f"Image {shape.name}"
                logging.exception(f"Could not generate a caption for {shape.name}. Falling back to the placeholder {alt_text!r}. Error: {e}")

            set_picture_alt_text(shape, alt_text)
            f.write(f" \n Shape: {shape.name} \n - Alt Text: {alt_text} \n"
                    f" - [{severity}] The image had no real description, so this one was generated. Check it (WCAG 1.1.1). \n")
            return severity
    elif shape.has_text_frame:
        alt_text = "Text content: " + shape.text
        f.write(f" \n Shape: {shape.name} \n - Alt Text: {alt_text} \n")
    else:
        f.write(f" \n Shape: {shape.name} \n - No alt text available. \n")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Fix every .pptx in a folder. The originals are left where they are.")
    parser.add_argument("input_folder", nargs="?", default="pptx_input")
    parser.add_argument("output_folder", nargs="?", default="pptx_output")
    parser.add_argument("--move-originals", action="store_true",
                        help="also move each original into the output folder once it is done (never replaces an existing file)")
    args = parser.parse_args()
    accessibility_processor(args.input_folder, args.output_folder, move_originals=args.move_originals)
