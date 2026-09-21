"""Offline/package tests for the bounded native relative-label preparer."""
from __future__ import annotations

import hashlib
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from lxml import etree as E

SCRIPTS = Path(__file__).resolve().parents[1] / "skills/thinkcell-edit/scripts"
PERCENT = SCRIPTS / "percent_labels"
sys.path[:0] = [str(SCRIPTS), str(PERCENT), str(SCRIPTS / "thinkcell_no_click/implementation")]

import prepare_native_relative_candidate as candidate  # noqa: E402
from discover_percent_semantics import discover  # noqa: E402


ROOT = Path(__file__).resolve().parents[3] / "research/native-percent-labels"


class NativeRelativeCandidateTests(unittest.TestCase):
    def test_absolute_fixture_prepares_one_relative_physical_field(self):
        source = ROOT / "00-native-absolute.pptx"
        expected_sha = hashlib.sha256(source.read_bytes()).hexdigest().upper()
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / "candidate.pptx"
            report = Path(folder) / "candidate.json"
            result = candidate.prepare(source, output, report, expected_sha=expected_sha,
                                       category_index=0, series_index=2)
            self.assertEqual(result["status"], "PREPARED_NATIVE_RELATIVE_CANDIDATE_NATIVE_PROOF_REQUIRED")
            self.assertEqual(result["source_sha256"], expected_sha)
            self.assertTrue(result["geometry_unchanged"])
            self.assertEqual(result["source_cached_label_color"], "bg1")
            self.assertTrue(result["physical_field_id"].startswith("{"))
            self.assertEqual(len(result["changed_entries"]), 5)
            rows = [row for row in discover(output)
                    if row["category_index"] == 0 and row["series_index"] == 2]
            self.assertEqual(len(rows), 1)
            self.assertEqual(len(rows[0]["physical_shapes"]), 1)
            self.assertIsNotNone(rows[0]["relative_text_variable"])
            self.assertIsNone(rows[0]["absolute_text_variable"])
            self.assertEqual(rows[0]["physical_shapes"][0]["fields"][0]["text"], "63%")
            self.assertEqual(json.loads(report.read_text(encoding="utf-8"))["output_sha256"],
                             result["output_sha256"])

    def test_semantic_selector_resolves_same_fixed_profile(self):
        source = ROOT / "00-native-absolute.pptx"
        expected_sha = hashlib.sha256(source.read_bytes()).hexdigest().upper()
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / "candidate.pptx"
            report = Path(folder) / "candidate.json"
            result = candidate.prepare(source, output, report, expected_sha=expected_sha,
                                       category="2024.0", series="Series 3")
            self.assertEqual(result["selector"]["category_index"], 0)
            self.assertEqual(result["selector"]["series_index"], 2)

    def test_hash_and_profile_guards_fail_before_output(self):
        source = ROOT / "00-native-absolute.pptx"
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / "candidate.pptx"
            report = Path(folder) / "candidate.json"
            with self.assertRaisesRegex(ValueError, "SHA-256"):
                candidate.prepare(source, output, report, expected_sha="0" * 64,
                                  category_index=0, series_index=2)
            self.assertFalse(output.exists())
            with self.assertRaisesRegex(ValueError, "zero decimal"):
                candidate.prepare(source, output, report,
                                  expected_sha=hashlib.sha256(source.read_bytes()).hexdigest(),
                                  category_index=0, series_index=2, digits=1)
            self.assertFalse(output.exists())

    def test_candidate_package_has_dynamic_tag_part_and_bg1_text_styles(self):
        source = ROOT / "00-native-absolute.pptx"
        expected_sha = hashlib.sha256(source.read_bytes()).hexdigest().upper()
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / "candidate.pptx"
            report = Path(folder) / "candidate.json"
            candidate.prepare(source, output, report, expected_sha=expected_sha,
                              category_index=0, series_index=2)
            with zipfile.ZipFile(output) as package:
                tag_parts = [name for name in package.namelist()
                             if name.startswith("ppt/tags/nativePercent-")]
                self.assertEqual(len(tag_parts), 1)
                slide = E.fromstring(package.read("ppt/slides/slide1.xml"))
                shape = next(shape for shape in slide.xpath(".//p:sp", namespaces=candidate.NS)
                             if shape.find(".//p:cNvPr", candidate.NS).get("name") ==
                             "Native percentage label candidate")
                self.assertEqual(shape.findtext(".//a:t", namespaces=candidate.NS), "63%")
                self.assertFalse(shape.xpath('.//a:schemeClr[@val="tx1"]', namespaces=candidate.NS))
                rels = E.fromstring(package.read("ppt/slides/_rels/slide1.xml.rels"))
                self.assertTrue(all(rel.tag == "{" + rels.tag.split("}")[0].lstrip("{") + "}Relationship"
                                    for rel in rels))


if __name__ == "__main__":
    unittest.main()
