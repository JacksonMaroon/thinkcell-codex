"""Offline tests for the native percentage before/after contract."""
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
import native_percent_contract as contract  # noqa: E402


ROOT = Path(__file__).resolve().parents[3] / "research/native-percent-labels"


def inventory_target(path: Path):
    _, charts, _ = discovery.inventory(path.read_bytes())
    return charts[0]


class NativePercentContractTests(unittest.TestCase):
    def test_snapshot_reports_verified_labels_and_ratios(self):
        source = ROOT / "02-controlled-percent.pptx"
        result = contract.snapshot(source, inventory_target(source))
        self.assertEqual(result["status"], "native_percentage")
        self.assertTrue(result["applicable"])
        self.assertEqual(result["label_count"], 9)
        self.assertEqual(result["category_count"], 3)
        self.assertEqual(result["series_count"], 3)
        target = next(row for row in result["labels"]
                      if row["category_index"] == 0 and row["series_index"] == 2)
        self.assertAlmostEqual(target["ratio"], 0.5)
        self.assertAlmostEqual(target["cache_expected"], 50.0)

    def test_verify_preserves_indices_and_precision_after_data_change(self):
        source = ROOT / "02-controlled-percent.pptx"
        before = contract.snapshot(source, inventory_target(source))
        result = contract.verify(before, ROOT / "03-numerator-change.pptx", inventory_target(source))
        self.assertEqual(result["status"], "native_percentage_preserved")
        self.assertTrue(all(result["checks"].values()))
        self.assertAlmostEqual(next(row for row in result["after"]["labels"]
                                    if row["category_index"] == 0 and row["series_index"] == 2)["ratio"],
                               200.0 / 300.0)

    def test_absolute_chart_is_not_applicable(self):
        source = ROOT / "00-native-absolute.pptx"
        result = contract.snapshot(source, inventory_target(source))
        self.assertEqual(result, {
            "status": "not_applicable",
            "applicable": False,
            "input": str(source),
            "target": result["target"],
            "labels": [],
            "label_count": 0,
            "count": 0,
        })
        self.assertEqual(contract.verify(result, source, inventory_target(source))["status"],
                         "not_applicable")

    def test_partial_direct_recognition_rejects(self):
        source = ROOT / "02-controlled-percent.pptx"
        rows = discovery.discover(source)
        rows[0]["verified_direct_percentage"] = False
        with patch.object(contract, "discover", return_value=rows), self.assertRaisesRegex(
                ValueError, "partial native direct percentage"):
            contract.snapshot(source, inventory_target(source))

    def test_target_requires_exact_preserved_frame_tag(self):
        source = ROOT / "02-controlled-percent.pptx"
        target = inventory_target(source)
        target["frames"][0]["shape_tag"] = "missing-shape-tag"
        with self.assertRaisesRegex(ValueError, "target chart was not found"):
            contract.snapshot(source, target)

    def test_verify_uses_indices_when_names_change_but_rejects_precision_change(self):
        source = ROOT / "02-controlled-percent.pptx"
        before = contract.snapshot(source, inventory_target(source))
        after_rows = discovery.discover(ROOT / "03-numerator-change.pptx")
        renamed = copy.deepcopy(after_rows)
        for row in renamed:
            row["category"] = f"renamed-{row['category']}"
            row["series"] = f"renamed-{row['series']}"
        with patch.object(contract, "discover", return_value=renamed):
            result = contract.verify(before, ROOT / "03-numerator-change.pptx", inventory_target(source))
        self.assertEqual(result["status"], "native_percentage_preserved")

        changed_precision = copy.deepcopy(after_rows)
        changed_precision[0]["direct_precision"]["decimal_digits"] = "1"
        with patch.object(contract, "discover", return_value=changed_precision), self.assertRaisesRegex(
                ValueError, "contract verification failed"):
            contract.verify(before, ROOT / "03-numerator-change.pptx", inventory_target(source))

    def test_non_sequence_target_is_not_applicable_even_with_sequence_rows(self):
        source = ROOT / "02-controlled-percent.pptx"
        target = copy.deepcopy(inventory_target(source))
        target["owner"] = {"id": target["owner"].get("id"), "tag": "CPieChartSE"}
        with patch.object(contract, "discover", side_effect=AssertionError("must return before discover")):
            result = contract.snapshot(source, target)
        self.assertEqual(result["status"], "not_applicable")
        self.assertFalse(result["applicable"])

    def test_label_free_sequence_target_is_not_applicable_when_sibling_has_rows(self):
        source = ROOT / "02-controlled-percent.pptx"
        target = inventory_target(source)
        sibling_rows = copy.deepcopy(discovery.discover(source))
        for row in sibling_rows:
            row["owning_chart"]["shape_tag"] = "sibling-shape-tag"
        with patch.object(contract, "discover", return_value=sibling_rows):
            result = contract.snapshot(source, target)
        self.assertEqual(result["status"], "not_applicable")
        self.assertEqual(result["count"], 0)


if __name__ == "__main__":
    unittest.main()
