"""Offline OCR. Confidence is evidence, never an accessibility decision."""
from __future__ import annotations

import base64
import hashlib
import math
from dataclasses import asdict, dataclass
from contextlib import closing
from importlib.metadata import version
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Callable

from .models import DocumentInspection, PageRoute
from .providers import BoundingBox, OcrLine, OcrPage, ProviderUnavailable
from .setup_models import MODEL_PATH, MODEL_SHA256, verified_model

DPI = 300
MAX_PIXELS = 16_000_000
MIN_LINE_CONFIDENCE = 0.80
MIN_PAGE_CONFIDENCE = 0.90


class RapidOcrProvider:
    """Pinned bundled PP-OCR models; CPU only; no runtime downloads."""

    def __init__(self) -> None:
        try:
            import rapidocr_onnxruntime
            from rapidocr_onnxruntime import RapidOCR
        except ImportError as error:
            raise ProviderUnavailable("Install pipeline/requirements.txt in the project .venv to enable OCR.") from error
        if not verified_model():
            raise ProviderUnavailable("English OCR model is missing or changed. Run python -m pipeline.setup_models once.")
        self.engine = RapidOCR(rec_model_path=str(MODEL_PATH), text_score=0.0,
                               intra_op_num_threads=2, inter_op_num_threads=2)
        model_dir = Path(rapidocr_onnxruntime.__file__).parent / "models"
        self.metadata = {
            "engine": "rapidocr-onnxruntime",
            "version": version("rapidocr-onnxruntime"),
            "runtime": version("onnxruntime"),
            "models": {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(model_dir.glob("*.onnx")) if "rec_" not in p.name},
            "confidenceUnit": "line recognition score, 0 to 1; not calibrated probability",
            "coordinates": "rendered image pixels, origin top left",
        }
        self.metadata["models"][MODEL_PATH.name] = MODEL_SHA256
        self.metadata["recognitionLanguage"] = "en"

    def recognize(self, page_number: int, image_path: str, language_hints: list[str]) -> OcrPage:
        # Bundled recognizer is fixed; hints do not select or download a language model.
        result, _ = self.engine(image_path)
        lines = []
        for polygon, text, score in result or []:
            if not str(text).strip():
                continue
            x, y = zip(*polygon)
            lines.append(OcrLine(str(text).strip(), BoundingBox(float(min(x)), float(min(y)),
                                float(max(x)), float(max(y))), float(score)))
        lines.sort(key=lambda line: (line.bounds.y0, line.bounds.x0))
        reasons = ["Review transcription, reading order, tables, and non-text content against the source scan."]
        if not lines:
            reasons.append("No text recognized; this may be blank, unreadable, or unsupported content.")
        else:
            mean = sum(line.engine_confidence for line in lines) / len(lines)
            if mean < MIN_PAGE_CONFIDENCE:
                reasons.append("Mean line confidence is below 0.90.")
            if any(line.engine_confidence < MIN_LINE_CONFIDENCE for line in lines):
                reasons.append("One or more lines have confidence below 0.80.")
        # Do not invent word boxes from line-level recognition results.
        return OcrPage(page_number, [], "en", True, lines, reasons)


@dataclass
class OcrContent:
    page: OcrPage
    image_data: str


OCR_ROUTES = frozenset({PageRoute.SCANNED, PageRoute.HYBRID})


def recognize_raster_pages(
    source: Path,
    inspection: DocumentInspection,
    evidence: dict,
    report: Callable[[str, str, str], None],
) -> dict[int, OcrContent]:
    routed_pages = [page for page in inspection.pages if page.route in OCR_ROUTES]
    if not routed_pages:
        return {}
    import pypdfium2 as pdfium

    provider = RapidOcrProvider()
    evidence.update(provider.metadata)
    evidence.update({"requestedDpi": DPI, "maxPixelsPerPage": MAX_PIXELS,
                     "minimumLineConfidence": MIN_LINE_CONFIDENCE,
                     "minimumMeanConfidence": MIN_PAGE_CONFIDENCE, "pages": []})
    content = {}
    with TemporaryDirectory(prefix="uw-ocr-") as directory, closing(pdfium.PdfDocument(str(source))) as document:
        for item in routed_pages:
            name = f"ocr-page-{item.page_number}"
            report(name, "running", f"Recognizing {item.route} page {item.page_number} locally.")
            with closing(document[item.page_number - 1]) as page:
                width, height = page.get_size()
                scale = min(DPI / 72, math.sqrt(MAX_PIXELS / max(width * height, 1)))
                bitmap = page.render(scale=scale)
                try:
                    image = bitmap.to_pil().convert("RGB")
                finally:
                    bitmap.close()
            path = Path(directory) / f"page-{item.page_number}.png"
            image.save(path)
            result = provider.recognize(item.page_number, str(path), [])
            if item.route == PageRoute.HYBRID:
                result.review_reasons.append(
                    "Hybrid page: compare the OCR draft with the embedded text layer and source image."
                )
            record = asdict(result)
            record.update({"sourceRoute": item.route, "pixelWidth": image.width, "pixelHeight": image.height,
                           "effectiveDpi": round(scale * 72, 2),
                           "meanConfidence": (sum(x.engine_confidence for x in result.lines) / len(result.lines)) if result.lines else None})
            evidence["pages"].append(record)
            # Keep a visual reference in the HTML/PDF so figures are not silently discarded.
            image.thumbnail((1600, 2200))
            image.save(path, format="PNG")
            content[item.page_number] = OcrContent(result, base64.b64encode(path.read_bytes()).decode("ascii"))
            image.close()
            report(name, "complete", f"Page {item.page_number}: recovered {len(result.lines)} text lines; review required.")
    return content


# Compatibility for callers of the initial scanned-only integration.
recognize_scanned_pages = recognize_raster_pages
