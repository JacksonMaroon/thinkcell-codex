import importlib.util
import unittest
from pathlib import Path

SCRIPT = Path(__file__).with_name("multi_chart_update.py")
SPEC = importlib.util.spec_from_file_location("release_multi_chart_update", SCRIPT)
module = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(module)

class AppearanceMergeTests(unittest.TestCase):
    def request(self, categories):
        return {"expected_model": {"categories": categories, "series_names": ["A", "B"]}}

    def appearance(self, series=None, points=None):
        return {"schema": "tc.datasheet-fill.v1", "series_fills": series or {}, "point_fills": points or {}}

    def test_partial_request_preserves_series_and_reorders_points_by_category(self):
        current = {
            "A": {"series": "111111", "points": {0: "AA0000", 2: "00AA00"}},
            "B": {"series": "222222", "points": {}},
        }
        actual = module.merge_enabled_appearance(
            current, ["North", "South", "West"], self.request(["West", "South", "North"]),
            self.appearance(points={"A": [None, "#0000AA", None]}),
        )
        self.assertEqual(actual["series_fills"], {"A": "#111111", "B": "#222222"})
        self.assertEqual(actual["point_fills"]["A"], ["#00AA00", "#0000AA", "#AA0000"])

    def test_series_override_clears_prior_points_then_applies_requested_points(self):
        current = {"A": {"series": "111111", "points": {0: "AA0000", 1: "00AA00"}}}
        request = {"expected_model": {"categories": ["North", "South"], "series_names": ["A"]}}
        actual = module.merge_enabled_appearance(
            current, ["North", "South"], request,
            self.appearance(series={"A": "#333333"}, points={"A": [None, "#0000AA"]}),
        )
        self.assertEqual(actual["series_fills"], {"A": "#333333"})
        self.assertEqual(actual["point_fills"]["A"], [None, "#0000AA"])

    def test_ambiguous_category_identity_is_rejected(self):
        current = {"A": {"series": "111111", "points": {0: "AA0000"}}}
        request = {"expected_model": {"categories": ["North", "North"], "series_names": ["A"]}}
        with self.assertRaisesRegex(ValueError, "ambiguous"):
            module.merge_enabled_appearance(current, ["North", "South"], request, self.appearance())

    def test_color_fingerprint_ignores_data_but_detects_style_change(self):
        xml1 = module.E.fromstring(b'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:ser><c:val>1</c:val><c:spPr><a:solidFill><a:srgbClr val="111111"/></a:solidFill></c:spPr></c:ser></c:chartSpace>')
        xml2 = module.E.fromstring(b'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:ser><c:val>2</c:val><c:spPr><a:solidFill><a:srgbClr val="111111"/></a:solidFill></c:spPr></c:ser></c:chartSpace>')
        xml3 = module.E.fromstring(b'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:ser><c:val>2</c:val><c:spPr><a:solidFill><a:srgbClr val="222222"/></a:solidFill></c:spPr></c:ser></c:chartSpace>')
        self.assertEqual(module._color_fingerprint_root(xml1), module._color_fingerprint_root(xml2))
        self.assertNotEqual(module._color_fingerprint_root(xml1), module._color_fingerprint_root(xml3))

if __name__ == "__main__":
    unittest.main()
