from __future__ import annotations

import unittest
import tempfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from pypdf import PdfWriter
from pypdf.generic import ArrayObject, DecodedStreamObject, DictionaryObject, NameObject, NumberObject

from pipeline.images import ImageRemediation, _classify, _validate_approval, analyze_digital_images, write_image_review
from pipeline.inspect_pdf import _complex_content, _raster_coverage, _route_page
from pipeline.models import PageRoute
from pipeline.providers import BoundingBox, OcrLine
from pipeline.reading_order import lines_from_words, order_layout_items
from pipeline.rebuild_html import _image_html, _linkify, _semantic_blocks, build_semantic_html


def digital_image_fixture(path: Path) -> None:
    writer = PdfWriter()
    page = writer.add_blank_page(width=612, height=792)
    image = DecodedStreamObject()
    image.set_data(bytes([35, 90, 180]) * 100)
    image.update({
        NameObject("/Type"): NameObject("/XObject"),
        NameObject("/Subtype"): NameObject("/Image"),
        NameObject("/Width"): NumberObject(10),
        NameObject("/Height"): NumberObject(10),
        NameObject("/ColorSpace"): NameObject("/DeviceRGB"),
        NameObject("/BitsPerComponent"): NumberObject(8),
    })
    image_ref = writer._add_object(image)
    font = DictionaryObject({
        NameObject("/Type"): NameObject("/Font"),
        NameObject("/Subtype"): NameObject("/Type1"),
        NameObject("/BaseFont"): NameObject("/Helvetica"),
    })
    page[NameObject("/Resources")] = DictionaryObject({
        NameObject("/XObject"): DictionaryObject({NameObject("/Im0"): image_ref}),
        NameObject("/Font"): DictionaryObject({NameObject("/F1"): writer._add_object(font)}),
    })
    content = DecodedStreamObject()
    text = "This digital page contains enough embedded text to remain on the digital route while its image is reviewed."
    content.set_data(f"BT /F1 12 Tf 40 700 Td ({text}) Tj ET q 150 0 0 100 40 520 cm /Im0 Do Q".encode("ascii"))
    page[NameObject("/Contents")] = writer._add_object(content)
    writer.write(path)


class PageRoutingTests(unittest.TestCase):
    def test_digital_page(self) -> None:
        route, reasons = _route_page(400, 0.02)
        self.assertEqual(route, PageRoute.DIGITAL)
        self.assertIn("usable embedded text", reasons)

    def test_scanned_page(self) -> None:
        route, _ = _route_page(0, 0.95)
        self.assertEqual(route, PageRoute.SCANNED)

    def test_hybrid_page(self) -> None:
        route, _ = _route_page(400, 0.95)
        self.assertEqual(route, PageRoute.HYBRID)

    def test_uncertain_page(self) -> None:
        route, _ = _route_page(10, 0.0)
        self.assertEqual(route, PageRoute.UNCERTAIN)

    def test_overlapping_images_are_counted_once(self) -> None:
        class Page:
            width = 100
            height = 100
            images = [
                {"x0": 0, "x1": 60, "top": 0, "bottom": 100},
                {"x0": 40, "x1": 100, "top": 0, "bottom": 100},
            ]

        self.assertEqual(_raster_coverage(Page()), 1.0)

    def test_image_bounds_are_clipped_to_page(self) -> None:
        class Page:
            width = 100
            height = 100
            images = [{"x0": -50, "x1": 50, "top": -50, "bottom": 50}]

        self.assertEqual(_raster_coverage(Page()), 0.25)

    def test_complex_content_records_columns_tables_images_and_graphics(self) -> None:
        page = SimpleNamespace(
            width=612,
            images=[{}],
            lines=[{}] * 20,
            rects=[],
            curves=[],
            find_tables=lambda: [object()],
        )
        words = []
        for y in (50, 80):
            words.extend([
                {"text": "Left", "x0": 40, "x1": 100, "top": y, "bottom": y + 12},
                {"text": "Right", "x0": 330, "x1": 400, "top": y, "bottom": y + 12},
            ])
        columns, tables, drawings, signals = _complex_content(page, words)
        self.assertEqual(columns, 2)
        self.assertEqual(tables, 1)
        self.assertEqual(drawings, 20)
        self.assertEqual(
            signals,
            ["multi-column-layout", "table", "embedded-image", "dense-vector-graphics"],
        )


class SemanticHtmlTests(unittest.TestCase):
    def test_builds_headings_lists_and_links(self) -> None:
        blocks = "".join(
            _semantic_blocks(
                [
                    "Overview",
                    "- First item",
                    "- Documentation: https://example.com/docs",
                    "Closing paragraph.",
                ]
            )
        )
        self.assertIn("<h3>Overview</h3>", blocks)
        self.assertIn("<ul>", blocks)
        self.assertIn('<a href="https://example.com/docs">', blocks)
        self.assertIn("<p>Closing paragraph.</p>", blocks)

    def test_query_string_link_is_escaped_once(self) -> None:
        linked = _linkify("Open https://example.com/search?a=one&b=two.")
        self.assertIn('href="https://example.com/search?a=one&amp;b=two"', linked)
        self.assertNotIn("&amp;amp;", linked)

    def test_lettered_items_keep_ordered_list_semantics(self) -> None:
        blocks = "".join(_semantic_blocks(["A. First", "B. Second"]))
        self.assertIn('<ol type="A">', blocks)
        self.assertIn("<li>First</li>", blocks)
        self.assertIn("<li>Second</li>", blocks)
        self.assertIn("</ol>", blocks)


class ReadingOrderTests(unittest.TestCase):
    @staticmethod
    def line(text: str, x0: float, y0: float, x1: float, y1: float) -> OcrLine:
        return OcrLine(text, BoundingBox(x0, y0, x1, y1), 0.99)

    def test_two_columns_are_read_down_the_left_then_down_the_right(self) -> None:
        lines = [
            self.line("Right 1", 320, 100, 500, 120),
            self.line("Left 1", 40, 90, 220, 110),
            self.line("Right 2", 320, 150, 500, 170),
            self.line("Left 2", 40, 140, 220, 160),
        ]
        result = order_layout_items(lines, 612)
        self.assertEqual([line.text for line in result.items], ["Left 1", "Left 2", "Right 1", "Right 2"])
        self.assertEqual(result.column_count, 2)
        self.assertEqual(result.method, "geometric-two-column-bands")
        self.assertTrue(result.warnings)

    def test_full_width_heading_stays_before_columns(self) -> None:
        lines = [
            self.line("Heading", 40, 20, 570, 50),
            self.line("R1", 320, 100, 500, 120),
            self.line("L1", 40, 90, 220, 110),
            self.line("R2", 320, 150, 500, 170),
            self.line("L2", 40, 140, 220, 160),
        ]
        result = order_layout_items(lines, 612)
        self.assertEqual([line.text for line in result.items], ["Heading", "L1", "L2", "R1", "R2"])

    def test_single_column_keeps_top_to_bottom_order(self) -> None:
        lines = [
            self.line("Second", 40, 100, 260, 120),
            self.line("First", 40, 50, 260, 70),
        ]
        result = order_layout_items(lines, 612)
        self.assertEqual([line.text for line in result.items], ["First", "Second"])
        self.assertEqual(result.column_count, 1)

    def test_words_on_same_row_are_split_at_a_column_gutter(self) -> None:
        words = [
            {"text": "Left", "x0": 40, "x1": 80, "top": 50, "bottom": 65},
            {"text": "side", "x0": 86, "x1": 120, "top": 50, "bottom": 65},
            {"text": "Right", "x0": 330, "x1": 375, "top": 50, "bottom": 65},
            {"text": "side", "x0": 381, "x1": 415, "top": 50, "bottom": 65},
        ]
        self.assertEqual([line.text for line in lines_from_words(words)], ["Left side", "Right side"])


class ImageRemediationTests(unittest.TestCase):
    def test_chart_language_routes_image_as_complex(self) -> None:
        classification, detected_type = _classify(
            0.2, "Figure 2. Enrollment chart", "", ["2024", "Enrollment"]
        )
        self.assertEqual(classification, "complex")
        self.assertEqual(detected_type, "chart")

    def test_unverified_numeric_claim_is_rejected(self) -> None:
        item = ImageRemediation(
            id="p1-i1", page_number=1, image_number=1,
            bounds={"x0": 0, "top": 0, "x1": 100, "bottom": 100}, page_coverage=0.1,
            image_path="image.png", detected_type="chart", classification="complex",
            caption="Enrollment chart", nearby_text="", extracted_text=["2023", "50"],
            extracted_values=["2023", "50"], alt_text="Enrollment reached 75 in 2023.",
            long_description="The chart shows enrollment values by year and has been reviewed.",
            accessible_data_table=None, evidence=[], approved=True,
        )
        issues = _validate_approval(item)
        self.assertTrue(any("75" in issue for issue in issues))

    def test_reviewer_can_record_a_visually_verified_value(self) -> None:
        item = ImageRemediation(
            id="p1-i1", page_number=1, image_number=1,
            bounds={"x0": 0, "top": 0, "x1": 100, "bottom": 100}, page_coverage=0.1,
            image_path="image.png", detected_type="chart", classification="complex",
            caption="Enrollment chart", nearby_text="", extracted_text=["2023"], extracted_values=["2023"],
            alt_text="Enrollment chart for 2023.",
            long_description="The chart shows a value of 75 for the 2023 enrollment series.",
            accessible_data_table=None, evidence=[], reviewer_verified_values=["75"],
            reviewer_notes="Read 75 directly from the labeled bar.", approved=True,
        )
        self.assertEqual(_validate_approval(item), [])

    def test_approved_image_is_inserted_with_alt_text(self) -> None:
        item = ImageRemediation(
            id="p1-i1", page_number=1, image_number=1,
            bounds={"x0": 0, "top": 0, "x1": 100, "bottom": 100}, page_coverage=0.1,
            image_path="image.png", detected_type="unclassified-image", classification="informative",
            caption="Blue square", nearby_text="", extracted_text=[], extracted_values=[],
            alt_text='A blue square labeled "sample".', long_description=None,
            accessible_data_table=None, evidence=[], approved=True, image_data="YWJj",
        )
        markup = _image_html(item)
        self.assertIn('alt="A blue square labeled &quot;sample&quot;."', markup)

    def test_generic_informative_placeholder_cannot_be_approved(self) -> None:
        item = ImageRemediation(
            id="p1-i1", page_number=1, image_number=1,
            bounds={"x0": 0, "top": 0, "x1": 100, "bottom": 100}, page_coverage=0.1,
            image_path="image.png", detected_type="unclassified-image", classification="informative",
            caption=None, nearby_text="", extracted_text=[], extracted_values=[],
            alt_text="Informative image on page 1; description required.", long_description=None,
            accessible_data_table=None, evidence=[], approved=True,
        )
        self.assertTrue(any("placeholder" in issue for issue in _validate_approval(item)))

    def test_digital_image_creates_review_record_and_accepts_approval(self) -> None:
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as directory:
            root = Path(directory)
            source = root / "digital-image.pdf"
            digital_image_fixture(source)
            from pipeline.inspect_pdf import inspect_pdf
            inspection = inspect_pdf(source)
            self.assertEqual(inspection.pages[0].route, PageRoute.DIGITAL)
            from pipeline.providers import ProviderUnavailable
            with patch("pipeline.images.RapidOcrProvider", side_effect=ProviderUnavailable("test")):
                images, issues = analyze_digital_images(source, inspection, root / "images")
            self.assertIn(1, images)
            self.assertTrue(any("Human approval" in issue for issue in issues))
            review = write_image_review(root / "review.json", images, issues)
            payload = __import__("json").loads(review.read_text(encoding="utf-8"))
            payload["images"][0].update({
                "classification": "informative",
                "alt_text": "Solid blue square used as a sample image.",
                "approved": True,
            })
            review.write_text(__import__("json").dumps(payload), encoding="utf-8")
            with patch("pipeline.images.RapidOcrProvider", side_effect=ProviderUnavailable("test")):
                approved, approval_issues = analyze_digital_images(source, inspection, root / "approved-images", review)
            self.assertEqual(approval_issues, [])
            html_path = build_semantic_html(source, inspection, root / "output.html", image_remediations=approved)
            output_html = html_path.read_text(encoding="utf-8")
            self.assertIn("Solid blue square", output_html)
            self.assertLess(output_html.index("This digital page"), output_html.index("Solid blue square"))


if __name__ == "__main__":
    unittest.main()
