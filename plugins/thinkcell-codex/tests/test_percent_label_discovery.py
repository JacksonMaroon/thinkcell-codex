"""Read-only regression coverage for native percentage-label discovery."""
from __future__ import annotations

import copy
import sys
import unittest
from pathlib import Path
from unittest.mock import patch


SCRIPTS = Path(__file__).resolve().parents[1] / "skills/thinkcell-edit/scripts"
PERCENT_LABELS = SCRIPTS / "percent_labels"
sys.path[:0] = [str(SCRIPTS), str(SCRIPTS / "thinkcell_no_click/implementation"), str(PERCENT_LABELS)]

import discover_percent_semantics as discovery  # noqa: E402


ROOT = Path(__file__).resolve().parents[3] / "research/native-percent-labels"
FIXTURES = sorted(ROOT.glob("[0-9][0-9]-*.pptx"))
EXPECTED_TARGET_RATIOS = {
    "01-native-percent.pptx": 55.0 / 100.0,
    "02-controlled-percent.pptx": 50.0 / 100.0,
    "03-numerator-change.pptx": 200.0 / 300.0,
    "04-denominator-change.pptx": 200.0 / 400.0,
    "05-reopened-saved.pptx": 200.0 / 400.0,
}


class PercentLabelDiscoveryTests(unittest.TestCase):
    def test_all_six_committed_fixtures_are_classified(self):
        self.assertEqual([p.name for p in FIXTURES], [
            "00-native-absolute.pptx",
            "01-native-percent.pptx",
            "02-controlled-percent.pptx",
            "03-numerator-change.pptx",
            "04-denominator-change.pptx",
            "05-reopened-saved.pptx",
        ])

        for fixture in FIXTURES:
            with self.subTest(fixture=fixture.name):
                rows = discovery.discover(fixture)
                self.assertEqual(len(rows), 9)
                self.assertTrue(all(row["owning_chart"]["exact"] for row in rows))
                self.assertTrue(all(row["owning_chart"]["owner_type"] == "CSequenceChartSE"
                                    for row in rows))
                if fixture.name == "00-native-absolute.pptx":
                    self.assertTrue(all(row["direct_rendered"] for row in rows))
                    self.assertTrue(all(row["classification"] == "unverified" for row in rows))
                    self.assertTrue(all(not row["verified_direct_percentage"] for row in rows))
                    self.assertTrue(all(not row["native_semantics"]["verified"] for row in rows))
                else:
                    self.assertTrue(all(row["direct_rendered"] for row in rows))
                    self.assertTrue(all(row["direct_precision"]["suffix"] == "%" for row in rows))
                    self.assertTrue(all(row["direct_precision"]["decimal_digits"] == "0" for row in rows))
                    self.assertTrue(all(row["native_semantics"]["verified"] for row in rows))
                    self.assertTrue(all(row["cache_agreement"]["verified"] for row in rows))
                    self.assertTrue(all(row["classification"] == "verified_direct_percentage" for row in rows))
                    self.assertTrue(all(row["verified_direct_percentage"] for row in rows))

    def test_changed_and_reopened_fixtures_preserve_semantic_target_ratio(self):
        for filename, expected_ratio in EXPECTED_TARGET_RATIOS.items():
            with self.subTest(fixture=filename):
                rows = discovery.discover(ROOT / filename)
                target = next(row for row in rows
                              if row["category"] == 2024.0 and row["series"] == "Series 3")
                self.assertEqual(target["category_index"], 0)
                self.assertEqual(target["series_index"], 2)
                self.assertAlmostEqual(target["numerator"] / target["denominator"], expected_ratio)
                self.assertAlmostEqual(target["cache_agreement"]["expected"], expected_ratio * 100.0)
                self.assertAlmostEqual(target["cache_agreement"]["observed"], expected_ratio * 100.0)

    def test_cache_mapping_handles_reversed_series_order(self):
        rows = discovery.discover(ROOT / "02-controlled-percent.pptx")
        mappings = {tuple(row["cache_agreement"]["series_cache_to_model"])
                    for row in rows}
        self.assertEqual(mappings, {(2, 1, 0)})
        for row in rows:
            with self.subTest(category=row["category"], series=row["series"]):
                self.assertEqual(row["cache"]["series_index"], 2 - row["series_index"])
                self.assertEqual(row["cache"]["category_index"], row["category_index"])
                self.assertAlmostEqual(row["cache"]["value"], row["cache_agreement"]["expected"])

    def test_cache_mapping_accepts_native_series_order(self):
        model = {
            "series_values": [[10.0, 20.0], [30.0, 40.0]],
            "category_extents": [100.0, 100.0],
            "categories": ["A", "B"],
        }
        cache = {"verified": True, "series": [
            {"values": {0: 10.0, 1: 20.0}},
            {"values": {0: 30.0, 1: 40.0}},
        ]}
        result = discovery._cache_matrix_agreement(model, cache)
        self.assertTrue(result["verified"])
        self.assertEqual(result["series_cache_to_model"], [0, 1])
        self.assertEqual(result["category_cache_to_model"], [0, 1])

    def test_cache_mapping_rejects_ambiguous_identity(self):
        model = {
            "series_values": [[50.0], [50.0]],
            "category_extents": [100.0],
            "categories": ["A"],
        }
        cache = {"verified": True, "series": [
            {"values": {0: 50.0}},
            {"values": {0: 50.0}},
        ]}
        result = discovery._cache_matrix_agreement(model, cache)
        self.assertFalse(result["verified"])
        self.assertEqual(result["reason"], "mapping_ambiguous")
        self.assertGreater(result["candidate_count"], 1)

    def test_cache_mapping_budget_fails_closed(self):
        model = {
            "series_values": [[10.0] for _ in range(9)],
            "category_extents": [100.0],
            "categories": ["A"],
        }
        cache = {"verified": True, "series": [{"values": {0: 10.0}} for _ in range(9)]}
        result = discovery._cache_matrix_agreement(model, cache)
        self.assertFalse(result["verified"])
        self.assertEqual(result["reason"], "mapping_bound_exceeded")

    def test_suffix_and_msgraph_flag_do_not_bypass_axis_or_numeric_guards(self):
        fixture = ROOT / "01-native-percent.pptx"
        false_axis = {"verified": False, "reason": "synthetic_missing_axis"}
        with patch.object(discovery, "_native_percent_axis", return_value=false_axis):
            rows = discovery.discover(fixture)

        self.assertTrue(all(row["direct_rendered"] for row in rows))
        self.assertTrue(all(row["direct_precision"]["suffix"] == "%" for row in rows))
        self.assertTrue(all(row["direct_precision"]["decimal_digits"] == "0" for row in rows))
        self.assertTrue(all(row["native_semantics"] == false_axis for row in rows))
        self.assertTrue(all(row["cache_agreement"]["verified"] for row in rows))
        self.assertTrue(all(row["classification"] == "unverified" for row in rows))
        self.assertTrue(all(not row["verified_direct_percentage"] for row in rows))

        false_cache = {"verified": False, "reason": "synthetic_numeric_mismatch"}
        with patch.object(discovery, "_cache_matrix_agreement", return_value=false_cache):
            rows = discovery.discover(fixture)
        self.assertTrue(all(row["direct_rendered"] for row in rows))
        self.assertTrue(all(row["direct_precision"]["suffix"] == "%" for row in rows))
        self.assertTrue(all(row["native_semantics"]["verified"] for row in rows))
        self.assertTrue(all(not row["cache_agreement"]["verified"] for row in rows))
        self.assertTrue(all(row["cache_agreement"]["mapping_reason"] == false_cache["reason"]
                            for row in rows))
        self.assertTrue(all(row["classification"] == "unverified" for row in rows))
        self.assertTrue(all(not row["verified_direct_percentage"] for row in rows))

    def test_field_backed_selector_guard_remains_usable_and_fail_closed(self):
        # The public checkout has no field-backed PPTX fixture; use the
        # smallest in-memory row required to exercise the existing guard.
        source = discovery.discover(ROOT / "02-controlled-percent.pptx")[0]
        field_backed = copy.deepcopy(source)
        field_backed.update({
            "category": "synthetic",
            "series": "field-backed",
            "chart_part": "synthetic-chart",
            "chart_name": "synthetic-chart",
            "physical_shapes": [{"shape_id": "shape-1", "fields": [{"id": "field-1"}]}],
            "relative_source_id": "relative-1",
            "relative_text_variable": "text-1",
            "relative_suffix": "%",
            "relative_decimal_digits": "0",
            "direct_rendered": False,
            "verified_direct_percentage": False,
            "classification": "unverified",
        })
        selected = discovery.select([field_backed], category="synthetic", series="field-backed",
                                    chart_part="synthetic-chart", chart_name="synthetic-chart")
        self.assertIs(selected, field_backed)

        missing_text = copy.deepcopy(field_backed)
        missing_text["relative_text_variable"] = None
        with self.assertRaisesRegex(ValueError, "no active relative field"):
            discovery.select([missing_text], category="synthetic", series="field-backed",
                             chart_part="synthetic-chart", chart_name="synthetic-chart")

        direct_rows = discovery.discover(ROOT / "02-controlled-percent.pptx")
        with self.assertRaisesRegex(ValueError, "exactly one physical shape"):
            discovery.select(direct_rows, category=direct_rows[0]["category"],
                             series=direct_rows[0]["series"], chart_part=direct_rows[0]["chart_part"])


if __name__ == "__main__":
    unittest.main()
