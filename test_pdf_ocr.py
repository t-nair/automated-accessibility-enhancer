from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from PIL import Image, ImageDraw, ImageFont
from pypdf import PdfWriter
from pypdf.generic import BooleanObject, DictionaryObject, NameObject

from pipeline.cli import run
from pipeline.inspect_pdf import inspect_pdf
from pipeline.models import PageRoute
from pipeline.ocr import RapidOcrProvider, recognize_raster_pages
from pipeline.rebuild_html import build_semantic_html


def scan_fixture(path: Path, text: bool = True) -> None:
    image = Image.new("RGB", (1275, 1650), "white")
    if text:
        # Pillow can resolve DejaVu Sans on Windows, macOS, and Linux. Keeping
        # the fixture portable lets the branch run in CI as well as locally.
        font = ImageFont.truetype("DejaVuSans.ttf", 36)
        draw = ImageDraw.Draw(image)
        for index, line in enumerate([
            "University of Washington",
            "Scanned document accessibility project",
            "Students can read this text after OCR.",
            "Please review the original page for accuracy.",
        ]):
            draw.text((90, 100 + index * 75), line, font=font, fill="black")
    image.save(path, "PDF", resolution=150)
    image.close()


def tagged_fixture(_html: Path, output: Path) -> None:
    writer = PdfWriter()
    writer.add_blank_page(width=612, height=792)
    writer._root_object.update({
        NameObject("/MarkInfo"): DictionaryObject({NameObject("/Marked"): BooleanObject(True)}),
        NameObject("/StructTreeRoot"): DictionaryObject({NameObject("/Type"): NameObject("/StructTreeRoot")}),
    })
    writer.write(output)


class OcrPolicyTests(unittest.TestCase):
    def provider(self, results):
        provider = RapidOcrProvider.__new__(RapidOcrProvider)
        provider.engine = Mock(return_value=(results, None))
        return provider

    def test_low_confidence_is_preserved_and_flagged(self):
        result = self.provider([([[0, 0], [30, 0], [30, 10], [0, 10]], "uncertain", 0.4)]).recognize(1, "unused", [])
        self.assertEqual(result.lines[0].text, "uncertain")
        self.assertTrue(any("0.80" in reason for reason in result.review_reasons))
        self.assertTrue(any("0.90" in reason for reason in result.review_reasons))
        self.assertEqual(result.words, [])  # No invented word coordinates.

    def test_high_confidence_still_requires_review(self):
        result = self.provider([([[0, 0], [30, 0], [30, 10], [0, 10]], "text", 0.99)]).recognize(1, "unused", [])
        self.assertTrue(result.needs_review)
        self.assertEqual(len(result.review_reasons), 1)

    def test_no_text_is_not_silently_treated_as_blank(self):
        result = self.provider(None).recognize(1, "unused", [])
        self.assertTrue(result.needs_review)
        self.assertIn("No text recognized", result.review_reasons[-1])


class OcrIntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.directory = tempfile.TemporaryDirectory()
        cls.root = Path(cls.directory.name)
        cls.source = cls.root / "scan.pdf"
        scan_fixture(cls.source)
        cls.inspection = inspect_pdf(cls.source)
        cls.evidence = {}
        cls.content = recognize_raster_pages(cls.source, cls.inspection, cls.evidence, lambda *_: None)

    @classmethod
    def tearDownClass(cls):
        cls.directory.cleanup()

    def test_real_bundled_model_recovers_image_only_text(self):
        self.assertEqual(self.inspection.pages[0].route, PageRoute.SCANNED)
        self.assertEqual(self.inspection.pages[0].text_characters, 0)
        text = " ".join(line.text for line in self.content[1].page.lines)
        self.assertIn("University of Washington", text)
        # Recognition can vary only in letter case across platform font builds.
        self.assertIn("students can read this text after ocr", text.lower())
        self.assertEqual(len(self.evidence["models"]), 3)

    def test_html_keeps_source_and_escapes_recognized_text(self):
        line = self.content[1].page.lines[0]
        original = line.text
        try:
            line.text = '<script>alert("test")</script>'
            path = build_semantic_html(self.source, self.inspection, self.root / "draft.html", self.content)
            html = path.read_text(encoding="utf-8")
            self.assertNotIn('<script>', html)
            self.assertIn("&lt;script&gt;", html)
            self.assertIn("data:image/png;base64,", html)
            self.assertIn("HUMAN REVIEW REQUIRED", html)
        finally:
            line.text = original

    def test_cli_returns_review_draft_with_artifacts(self):
        with patch("pipeline.cli.recognize_raster_pages", return_value=self.content), patch("pipeline.cli.export_tagged_pdf", side_effect=tagged_fixture):
            manifest = run(self.source, self.root / "result")
        self.assertEqual(manifest["status"], "review-required")
        self.assertTrue(Path(manifest["outputs"]["pdf"]).exists())
        self.assertTrue(Path(manifest["outputs"]["html"]).exists())

    def test_unmarked_export_is_blocked(self):
        def export(_html, output):
            writer = PdfWriter()
            writer.add_blank_page(width=612, height=792)
            writer.write(output)
        with patch("pipeline.cli.recognize_raster_pages", return_value=self.content), patch("pipeline.cli.export_tagged_pdf", side_effect=export):
            with self.assertRaisesRegex(RuntimeError, "tagged-structure"):
                run(self.source, self.root / "unmarked")
        manifest = json.loads((self.root / "unmarked/scan.remediation.json").read_text())
        self.assertEqual(manifest["outputs"], {})

    def test_hybrid_requires_ocr_evidence(self):
        inspection = inspect_pdf(self.source)
        inspection.pages[0].route = PageRoute.HYBRID
        with self.assertRaisesRegex(RuntimeError, "hybrid"):
            build_semantic_html(self.source, inspection, self.root / "hybrid.html")

        path = build_semantic_html(
            self.source, inspection, self.root / "hybrid-draft.html", self.content
        )
        self.assertIn("OCR DRAFT", path.read_text(encoding="utf-8"))

    def test_hybrid_page_enters_ocr_and_records_its_route(self):
        inspection = inspect_pdf(self.source)
        inspection.pages[0].route = PageRoute.HYBRID
        evidence = {}
        content = recognize_raster_pages(self.source, inspection, evidence, lambda *_: None)
        self.assertIn(1, content)
        self.assertEqual(evidence["pages"][0]["sourceRoute"], PageRoute.HYBRID)
        self.assertTrue(any("Hybrid page" in reason for reason in content[1].page.review_reasons))


if __name__ == "__main__":
    unittest.main()
