from __future__ import annotations

import unittest

from pipeline.inspect_pdf import _raster_coverage, _route_page
from pipeline.models import PageRoute
from pipeline.rebuild_html import _linkify, _semantic_blocks


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


if __name__ == "__main__":
    unittest.main()
