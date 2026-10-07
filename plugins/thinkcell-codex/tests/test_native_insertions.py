"""Native insertion guard tests use synthetic XML only, no proprietary donors."""
import copy
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

from lxml import etree as E

SCRIPTS = Path(__file__).resolve().parents[1] / "skills/thinkcell-edit/scripts"
sys.path[:0] = [str(SCRIPTS), str(SCRIPTS / "thinkcell_no_click/implementation")]
import native_insertions as adapter


class GraphTests(unittest.TestCase):
    def test_boundary_rebind_retains_sibling_and_allocates_fresh_ids(self):
        root = E.fromstring(b'<root><CChart id="1"/><CSibling id="20"><fixed val="original"/></CSibling></root>')
        sibling = E.tostring(root[1])
        donor = E.fromstring(b'<root><Feature id="8"><owner idref="3"/><label idref="9"/></Feature><Label id="9"><m_bstrShapeName>native-tag</m_bstrShapeName></Label><Chart id="3"/></root>')
        ids = {x.get("id"): x for x in donor}
        closure = adapter.dependency_closure(ids, ["8"], {"3": "1"})
        self.assertEqual(closure, {"8", "9"})
        mapping, tags, clones = adapter.clone_graph(root, ids, closure, {"3": "1"})
        self.assertEqual(E.tostring(root[1]), sibling)
        self.assertEqual(mapping["8"], "21")
        self.assertEqual(clones[0].find("owner").get("idref"), "1")
        self.assertNotEqual(tags["native-tag"], "native-tag")
        self.assertEqual(len(tags["native-tag"]), 23)
        self.assertEqual(E.tostring(donor), b'<root><Feature id="8"><owner idref="3"/><label idref="9"/></Feature><Label id="9"><m_bstrShapeName>native-tag</m_bstrShapeName></Label><Chart id="3"/></root>')

    def test_unresolved_reference_is_rejected(self):
        node = E.fromstring(b'<Feature id="8"><unresolved idref="404"/></Feature>')
        with self.assertRaisesRegex(ValueError, "Unresolved"):
            adapter.dependency_closure({"8": node}, ["8"], {})

    def test_unbound_clone_is_rejected(self):
        root = E.fromstring(b'<root><Chart id="1"/></root>')
        node = E.fromstring(b'<Feature id="8"><owner idref="404"/></Feature>')
        with self.assertRaisesRegex(ValueError, "Unbound"):
            adapter.clone_graph(root, {"8": node}, {"8"}, {})


class Guards(unittest.TestCase):
    def test_alias_rejected_before_source_or_native_acquisition(self):
        with tempfile.TemporaryDirectory() as td, patch.object(adapter, "_select") as select:
            p = Path(td) / "same.pptx"
            with self.assertRaisesRegex(ValueError, "distinct"):
                adapter.prepare(p, p, Path(td)/"donor.pptx", "legend", {}, {}, "A"*64, "B"*64)
            select.assert_not_called()

    def test_existing_output_is_preserved(self):
        with tempfile.TemporaryDirectory() as td, patch.object(adapter, "_select") as select:
            p = Path(td); output = p/"output.pptx"; output.write_bytes(b"keep")
            with self.assertRaisesRegex(ValueError, "overwrite"):
                adapter.prepare(p/"source.pptx", output, p/"donor.pptx", "legend", {}, {}, "A"*64, "B"*64)
            self.assertEqual(output.read_bytes(), b"keep")
            select.assert_not_called()

    def test_wrong_source_hash_rejected_before_inventory(self):
        with tempfile.TemporaryDirectory() as td, patch.object(adapter.naming, "inventory") as inventory:
            source = Path(td)/"source.pptx"; source.write_bytes(b"original")
            with self.assertRaisesRegex(ValueError, "hash changed"):
                adapter._select(source, {}, "A"*64)
            inventory.assert_not_called()
            self.assertEqual(source.read_bytes(), b"original")


class ScatterBinding(unittest.TestCase):
    def chart(self):
        series = E.fromstring(b'<Series id="3"><m_varsrc idref="4"/></Series>')
        source = E.fromstring(b'<Source id="4"><m_varval>Group A</m_varval></Source>')
        return {"doc": {"ids": {"3": series, "4": source}}}

    def data(self):
        return {"group_labels": ["Group A", "Group A", "Group B"], "x_values": [1,2,10], "y_values": [3,5,99]}

    def cache(self):
        return E.fromstring(b'<chart><scatterChart><ser><xVal><numRef><numCache><pt idx="0"><v>1</v></pt><pt idx="1"><v>2</v></pt></numCache></numRef></xVal><yVal><numRef><numCache><pt idx="0"><v>3</v></pt><pt idx="1"><v>5</v></pt></numCache></numRef></yVal></ser></scatterChart></chart>')

    def test_correct_model_physical_group_is_selected(self):
        cache = self.cache()
        self.assertIs(adapter.scatter_carrier(cache, self.chart(), "3", self.data()), cache.find("scatterChart/ser"))

    def test_one_point_group_rejected(self):
        data = self.data(); data["group_labels"][1] = "Group B"
        with self.assertRaisesRegex(ValueError, "two points"):
            adapter.scatter_carrier(self.cache(), self.chart(), "3", data)

    def test_wrong_index_closure_rejected(self):
        cache = self.cache(); cache.find("scatterChart/ser/yVal/numRef/numCache/pt").set("idx", "5")
        with self.assertRaisesRegex(ValueError, "index closure"):
            adapter.scatter_carrier(cache, self.chart(), "3", self.data())

    def test_duplicate_group_carriers_rejected(self):
        cache = self.cache(); cache.find("scatterChart").append(copy.deepcopy(cache.find("scatterChart/ser")))
        with self.assertRaisesRegex(ValueError, "ambiguous"):
            adapter.scatter_carrier(cache, self.chart(), "3", self.data())

    def test_duplicate_point_indices_rejected(self):
        cache = self.cache()
        points = cache.find("scatterChart/ser/xVal/numRef/numCache")
        points.append(copy.deepcopy(points[0]))
        with self.assertRaisesRegex(ValueError, "duplicated"):
            adapter.scatter_carrier(cache, self.chart(), "3", self.data())


class FullNumericCache(unittest.TestCase):
    def cache(self):
        return E.fromstring(b'<errBars><plus><numRef><numCache><ptCount val="2"/><pt idx="0"><v>-5</v></pt><pt idx="1"><v>-9</v></pt></numCache></numRef></plus></errBars>')

    def test_reordered_points_read_by_semantic_index(self):
        root = self.cache(); cache = root.find("plus/numRef/numCache")
        cache.append(cache[1])
        self.assertEqual(adapter.indexed_numeric_values(root, "plus"), [-5,-9])

    def test_duplicate_indices_rejected(self):
        root = self.cache(); root.find("plus/numRef/numCache/pt[@idx='1']").set("idx", "0")
        with self.assertRaisesRegex(ValueError, "ambiguous"):
            adapter.indexed_numeric_values(root, "plus")

    def test_stale_point_count_rejected(self):
        root = self.cache(); root.find("plus/numRef/numCache/ptCount").set("val", "3")
        with self.assertRaisesRegex(ValueError, "point count"):
            adapter.indexed_numeric_values(root, "plus")

    def test_wrong_indices_rejected(self):
        root = self.cache(); root.find("plus/numRef/numCache/pt[@idx='1']").set("idx", "9")
        with self.assertRaisesRegex(ValueError, "point count"):
            adapter.indexed_numeric_values(root, "plus")

    def test_nonfinite_extent_rejected(self):
        root = self.cache(); root.find("plus/numRef/numCache/pt/v").text = "NaN"
        with self.assertRaisesRegex(ValueError, "nonfinite"):
            adapter.indexed_numeric_values(root, "plus")


class CategoryBinding(unittest.TestCase):
    def scatter(self):
        return E.fromstring(b'<scatterChart><ser><xVal><numRef><numCache><ptCount val="2"/><pt idx="0"><v>0.5</v></pt><pt idx="1"><v>1.5</v></pt></numCache></numRef></xVal></ser></scatterChart>')

    def test_vertical_scatter_categories_have_exact_positions(self):
        adapter.verify_category_positions(self.scatter()[0], "y", ["A", "B"])

    def test_permuted_orthogonal_positions_rejected(self):
        root = self.scatter()
        points = root.findall('.//pt/v')
        points[0].text, points[1].text = points[1].text, points[0].text
        with self.assertRaisesRegex(ValueError, "orthogonal category positions"):
            adapter.verify_category_positions(root[0], "y", ["A", "B"])

    def test_line_labels_cannot_be_permuted_with_same_value_vector(self):
        root = E.fromstring(b'<lineChart><ser><cat><strRef><strCache><ptCount val="2"/><pt idx="0"><v>B</v></pt><pt idx="1"><v>A</v></pt></strCache></strRef></cat></ser></lineChart>')
        with self.assertRaisesRegex(ValueError, "category order"):
            adapter.verify_category_positions(root[0], "y", ["A", "B"])


class PowerProfile(unittest.TestCase):
    def data(self):
        return {"group_labels": ["A"]*3, "x_values": [1, 2, 4], "y_values": [3, 12, 48]}

    def test_independent_known_power_fit_and_changed_data(self):
        data = self.data()
        coefficient, exponent = adapter.power_fit(data, "A")
        self.assertAlmostEqual(coefficient, 3)
        self.assertAlmostEqual(exponent, 2)
        data["y_values"] = [v*2 for v in data["y_values"]]
        changed_coefficient, changed_exponent = adapter.power_fit(data, "A")
        self.assertAlmostEqual(changed_coefficient, 6)
        self.assertAlmostEqual(changed_exponent, exponent)

    def test_nonpositive_power_point_rejected(self):
        data = self.data(); data["x_values"][0] = 0
        with self.assertRaisesRegex(ValueError, "positive"):
            adapter.power_fit(data, "A")

    def test_constant_power_x_rejected(self):
        data = self.data(); data["x_values"] = [1, 1, 1]
        with self.assertRaisesRegex(ValueError, "distinct X"):
            adapter.power_fit(data, "A")

    def test_observed_forecast_proves_fit_dependent_clip(self):
        chart = E.fromstring(b'<chart><valAx><axPos val="b"/><scaling><min val="0"/><max val="9"/></scaling></valAx><valAx><axPos val="l"/><scaling><max val="48"/></scaling></valAx></chart>')
        trend = E.fromstring(b'<trendline><forward val="0"/></trendline>')
        result = adapter.verify_power_forecast(chart, trend, 4, 3, 2)
        self.assertEqual(result["actual_upper_x"], 4)
        self.assertTrue(result["fit_limits_visible_endpoint"])
        trend[0].set("val", "5")
        with self.assertRaisesRegex(ValueError, "contradicts"):
            adapter.verify_power_forecast(chart, trend, 4, 3, 2)


class LegendSwatch(unittest.TestCase):
    def test_native_owner_vector_color_cannot_be_replaced_by_orphan_swatch(self):
        chart = E.fromstring(b'<chart><barChart><ser><spPr><solidFill><schemeClr val="accent2"/></solidFill></spPr><val><numRef><numCache><ptCount val="2"/><pt idx="0"><v>5</v></pt><pt idx="1"><v>9</v></pt></numCache></numRef></val></ser></barChart></chart>')
        swatch = E.fromstring(b'<sp><spPr><solidFill><schemeClr val="accent2"/></solidFill></spPr></sp>')
        adapter.verify_legend_swatch(chart, swatch, [5, 9])
        swatch.find('.//schemeClr').set('val', 'accent3')
        with self.assertRaisesRegex(ValueError, "swatch color"):
            adapter.verify_legend_swatch(chart, swatch, [5, 9])


if __name__ == "__main__":
    unittest.main()
