"""Portable selection guards. No installed vendor assets or Office required."""
import unittest
from unittest.mock import patch
from pathlib import Path
import sys
import argparse
import tempfile
from lxml import etree as E

SCRIPTS = Path(__file__).resolve().parents[1] / 'skills/thinkcell-edit/scripts'
sys.path.insert(0, str(SCRIPTS))
import native_library_create as library


def cache(body):
    return E.fromstring(('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">'
                         + body + '</c:chartSpace>').encode())


class NativeLibrarySelectionTests(unittest.TestCase):
    def test_type_uses_native_owner_and_cache_together(self):
        sequence = "CSequenceChartSE"
        self.assertEqual(library.classify(sequence, sequence, cache('<c:scatterChart/>')), "line")
        self.assertEqual(library.classify("CScatterChartSE", "CScatterChartSE", cache('<c:scatterChart/>')), "scatter")
        self.assertEqual(library.classify("waterfall", sequence, cache('<c:barChart/>')), "waterfall")
        self.assertEqual(library.classify(sequence, sequence, cache('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/></c:barChart>')), "clustered-column")

    def test_combination_never_masquerades_as_plain_column(self):
        self.assertEqual(library.classify("CSequenceChartSE", "CSequenceChartSE",
                                         cache('<c:barChart/><c:lineChart/>')), "combination")

    def test_unknown_cache_fails_closed(self):
        self.assertIsNone(library.classify("CSequenceChartSE", "CSequenceChartSE", cache('<c:barChart/>')))
        self.assertIsNone(library.classify("CSequenceChartSE", "CSequenceChartSE", None))

    def row(self, slide=1, eligible=True, features=None):
        return {"source": "native-library.potx", "slide_number": slide,
                "type": "clustered-column", "generation_eligible": eligible,
                "features_on_slide": features or {}}

    @patch.object(library, "compatible", return_value=("flexible_sequence", {"count_change": True}))
    def test_multi_chart_seed_is_not_silently_selected(self, compatibility):
        catalog = {"templates": [self.row(1, False), self.row(2)]}
        chosen, route, contract = library.choose_template(catalog, "clustered-column", {})
        self.assertEqual(chosen["slide_number"], 2)
        self.assertEqual(compatibility.call_count, 1)

    @patch.object(library, "compatible", side_effect=ValueError("reserved row disagreement"))
    def test_invalid_contract_rejected_before_office(self, compatibility):
        with self.assertRaisesRegex(ValueError, "reserved row disagreement"):
            library.choose_template({"templates": [self.row()]}, "clustered-column", {})

    @patch.object(library, "compatible", return_value=("flexible_sequence", {}))
    def test_required_native_feature_matches_exact_class(self, compatibility):
        catalog = {"templates": [self.row(1), self.row(2, features={"CSequenceChartDataVectorCAGR": 1})]}
        selected, _, _ = library.choose_template(catalog, "clustered-column", {}, required_features=["CSequenceChartDataVectorCAGR"])
        self.assertEqual(selected["slide_number"], 2)
        with self.assertRaisesRegex(ValueError, "No installed native template"):
            library.choose_template(catalog, "clustered-column", {}, required_features=["CFakeFeature"])

    def test_empty_library_has_no_fake_fallback(self):
        with self.assertRaisesRegex(ValueError, "No installed native template"):
            library.choose_template({"templates": []}, "clustered-column", {})

    def test_construction_plan_binds_first_slide_and_exact_tag(self):
        chart = {"doc": {"slide_number": 1}, "frames": [{"shape_tag": "ExactNativeTag"}]}
        plan = library.sequence_plan(chart, {"matrix": []})
        self.assertEqual(plan["targets"][0]["selector"], {"slide_number": 1, "shape_tag": "ExactNativeTag"})
        from prepare_thinkcell_name import choose
        candidate = {"doc": {"slide_number": 1, "slide_id": 42}, "frames": [{"shape_id": 9, "shape_tag": "ExactNativeTag"}], "tags": ["ExactNativeTag"], "exact": True}
        self.assertIs(choose([candidate], **plan["targets"][0]["selector"]), candidate)

    def test_construction_plan_rejects_unextracted_slide(self):
        chart = {"doc": {"slide_number": 3}, "frames": [{"shape_tag": "ExactNativeTag"}]}
        with self.assertRaisesRegex(ValueError, "extracted first slide"):
            library.sequence_plan(chart, {})

    def test_normalized_seed_reuses_its_existing_unique_name(self):
        chart = {"doc": {"slide_number": 1}, "frames": [{"shape_tag": "ExactNativeTag"}],
                 "owner_name": "NormalizedExistingName", "table_name": "NormalizedExistingName"}
        chart["owner"] = E.fromstring('<CSequenceChartSE><m_strName>NormalizedExistingName</m_strName></CSequenceChartSE>')
        chart["table"] = E.fromstring('<Table><m_strName>NormalizedExistingName</m_strName></Table>')
        plan = library.sequence_plan(chart, {})
        self.assertEqual(plan["targets"][0]["name"], "NormalizedExistingName")
        from prepare_thinkcell_name import name_plan
        name, reused = name_plan(chart, [{"name": "NormalizedExistingName"}], "hash", plan["targets"][0]["name"])
        self.assertEqual(name, "NormalizedExistingName")
        self.assertTrue(reused)

    @patch.object(library, "digest", return_value="EXACT_SOURCE_HASH")
    def test_old_schema_normalization_uses_unchanged_data_and_existing_native_gate(self, digest):
        chart = {"frames": [{"shape_tag": "ExactNativeTag"}]}
        arguments = library.normalization_args(Path("seed.pptx"), chart,
                                  Path("unchanged-baseline.json"), Path("artifacts"), None)
        self.assertEqual(arguments.data_json, Path("unchanged-baseline.json"))
        self.assertEqual(arguments.expected_sha256, "EXACT_SOURCE_HASH")
        self.assertEqual(arguments.slide_number, 1)
        self.assertEqual(arguments.shape_tag, "ExactNativeTag")
        self.assertTrue(arguments.execute)
        self.assertEqual(arguments.output, Path("artifacts/normalized-seed.pptx"))

    def test_request_changed_during_preflight_never_opens_office(self):
        import thinkcell
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            request_path = root / "request.json"
            request_path.write_text('{}', encoding="utf-8")
            args = argparse.Namespace(data_json=request_path, library_root=None, chart_type="clustered-column",
                     slide_number=None, template=None, output_directory=root / "new-output", execute=True)
            def selection(*unused):
                request_path.write_text('{"changed": true}', encoding="utf-8")
                return ({"source": "installed.potx"}, "flexible_sequence", {})
            with patch.object(library, "discover", return_value={}), patch.object(library, "choose_template", side_effect=selection), patch.object(thinkcell, "create") as office_create:
                with self.assertRaisesRegex(ValueError, "Requested data changed during preflight"):
                    library.construct.__wrapped__(args)
                office_create.assert_not_called()
                self.assertFalse(args.output_directory.exists())


if __name__ == "__main__":
    unittest.main()
