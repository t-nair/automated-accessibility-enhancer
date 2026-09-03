import os
import io
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

# a real title placeholder is authoritative, the shape name is only a fallback
TITLE_PLACEHOLDERS = (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)

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
        with open(alt_text_output_file, "w", encoding="utf-8") as f:
            fix_titles_and_log_alt_text(prs, f, filename, progress_callback)
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
    repeats the shape name.
    """
    if not alt_text:
        return False

    candidate = alt_text.strip().lower()

    return (candidate.endswith(IMAGE_EXTENSIONS)
            or candidate == shape_name.strip().lower())


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


def fix_titles_and_log_alt_text(prs, f, filename, progress_callback=None):
    total_to_caption = count_images_needing_captions(prs)
    captions_done = 0

    if progress_callback is not None:
        progress_callback(captions_done, total_to_caption)

    for slide_number, slide in enumerate(prs.slides, start=1):

        if len(slide.shapes) == 0:
            logging.warning(f"{filename}: slide {slide_number} has no shapes, skipping title fix.")
            continue

        slide_text = get_slide_text(slide)

        moved = move_titles_to_front(slide)

        if moved:
            logging.info(f"{filename}: slide {slide_number} moved {moved} title shape(s) to the front.")

        # read the shapes back after reordering, so the report lists them in the
        # order a screen reader will actually announce them
        for shape in list(slide.shapes):
            needed_caption = needs_caption(shape)

            write_alt_text_line(shape, f, slide_text)

            if needed_caption:
                captions_done += 1

                if progress_callback is not None:
                    progress_callback(captions_done, total_to_caption)


def get_picture_alt_text(shape):
    # python-pptx doesn't expose a working .alt_text property in this version,
    # so we have to read the real "descr" attribute straight from the XML
    return shape._element.nvPicPr.cNvPr.get("descr")


def set_picture_alt_text(shape, alt_text):
    shape._element.nvPicPr.cNvPr.set("descr", alt_text)


def write_alt_text_line(shape, f, slide_text=""):
    if is_picture(shape):
        existing_alt_text = get_picture_alt_text(shape)

        if existing_alt_text and not is_placeholder_alt_text(existing_alt_text, shape.name):
            f.write(f" \n Shape: {shape.name} \n - Alt Text: {existing_alt_text} \n")
        else:
            if existing_alt_text:
                logging.info(
                    f"{shape.name}: replacing filename-style alt text "
                    f"{existing_alt_text!r} with a real description.")

            try:
                alt_text = generate_image_caption(shape, slide_text)
            except Exception as e:
                logging.warning(f"Could not generate a caption for {shape.name}. Falling back to a placeholder. Error: {e}")
                alt_text = f"Image {shape.name}"

            set_picture_alt_text(shape, alt_text)
            f.write(f" \n Shape: {shape.name} \n - Alt Text: {alt_text} \n")
    elif shape.has_text_frame:
        alt_text = "Text content: " + shape.text
        f.write(f" \n Shape: {shape.name} \n - Alt Text: {alt_text} \n")
    else:
        f.write(f" \n Shape: {shape.name} \n - No alt text available. \n")


if __name__ == "__main__":
    directory_path = "pptx_input"
    new_directory = "pptx_output"
    accessibility_processor(directory_path, new_directory)
