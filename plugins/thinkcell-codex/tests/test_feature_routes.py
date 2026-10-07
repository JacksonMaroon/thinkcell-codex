"""Synthetic feature adapter contracts; no Office, donors or network required."""
from pathlib import Path
import copy
import importlib.util
import sys
import tempfile
import unittest
from unittest.mock import patch
from lxml import etree as E

SCRIPTS = Path(__file__).resolve().parents[1] / 'skills/thinkcell-edit/scripts'
sys.path[:0] = [str(SCRIPTS), str(SCRIPTS/'thinkcell_no_click/implementation')]
import feature_pipeline as pipeline
import native_label_controls as labels

def load(name, relative):
    spec = importlib.util.spec_from_file_location(name, SCRIPTS/relative)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

errorbars = load('test_errorbar_adapter', 'errorbars/reusable_errorbar_route.py')
trendline = load('test_trendline_adapter', 'trendline/portable_trendline_implementation.py')
legend = load('test_legend_adapter', 'legend_controls/legend_high_adapter.py')

def feature_plan(kind='datasheet_fill_enable', **extra):
    return {'schema': pipeline.SCHEMA, 'source_sha256': 'A'*64, 'data_plan': {},
            'features': [{'kind': kind, 'selector': {'slide_number': 1, 'shape_tag': 'tag'}, **extra}]}

class PipelineGuardTests(unittest.TestCase):
    def test_selector_and_digest_malformed_rejected(self):
        for key, value in [('slide_number', True), ('slide_number', 0), ('shape_tag', ''), ('shape_tag', 42)]:
            plan = feature_plan(); plan['features'][0]['selector'][key] = value
            with self.assertRaises(ValueError): pipeline.validate_plan(plan)
        plan = feature_plan(); plan['source_sha256'] = 'z'*64
        with self.assertRaises(ValueError): pipeline.validate_plan(plan)

    def test_relative_schema_reaches_explicit_unsupported_gate_without_side_effects(self):
        plan = feature_plan('relative_label_content', relative_field_id='field')
        pipeline.validate_plan(plan)
        with patch.object(pipeline, '_dependencies') as dependencies, tempfile.TemporaryDirectory() as td:
            folder = Path(td)
            with self.assertRaisesRegex(ValueError, 'experimental'):
                pipeline.run(folder/'input.pptx', folder/'out.pptx', folder/'report.json', plan, execute=True)
            dependencies.assert_not_called()
            self.assertEqual(list(folder.iterdir()), [])

class LabelPlanGuardTests(unittest.TestCase):
    def test_alias_and_existing_destination_rejected_before_acquisition(self):
        with tempfile.TemporaryDirectory() as td, patch.object(labels, 'make_plan') as acquire:
            folder = Path(td); source = folder/'source.pptx'; source.write_bytes(b'original')
            with self.assertRaisesRegex(ValueError, 'distinct'):
                labels.write_plan(source, folder/'target.json', folder/'controls.json', source)
            output = folder/'plan.json'; output.write_text('existing')
            with self.assertRaisesRegex(ValueError, 'already exists'):
                labels.write_plan(source, folder/'target.json', folder/'controls.json', output)
            acquire.assert_not_called()
            self.assertEqual(source.read_bytes(), b'original')
            self.assertEqual(output.read_text(), 'existing')

class ErrorbarGuardTests(unittest.TestCase):
    def setUp(self):
        self.data = {'categories': ['A', 'B'], 'series': {'Min': [1, 2], 'Max': [4, 5], 'Marker': [2, 3]}}

    def test_json_key_order_does_not_change_semantic_series_order(self):
        series = {k: self.data['series'][k] for k in ('Marker', 'Max', 'Min')}
        result = errorbars.table_payload(self.data['categories'], series)
        self.assertEqual([row[0]['string'] for row in result[1:]], ['Min', 'Max', 'Marker'])

    def test_invalid_bounds_and_nonfinite_values_rejected(self):
        for value in (True, float('nan'), float('inf'), '1'):
            bad = copy.deepcopy(self.data); bad['series']['Min'][0] = value
            with self.assertRaisesRegex(ValueError, 'finite numbers'):
                errorbars.table_payload(bad['categories'], bad['series'])
        self.data['series']['Min'][0] = 9
        with self.assertRaisesRegex(ValueError, 'Min must not exceed Max'):
            errorbars.table_payload(self.data['categories'], self.data['series'])

    def test_bad_data_rejected_before_naming_or_output_write(self):
        with tempfile.TemporaryDirectory() as td, patch.object(errorbars.naming, 'inventory') as inventory:
            folder = Path(td); source = folder/'source.pptx'; source.write_bytes(b'donor')
            bad = copy.deepcopy(self.data); bad['series']['Max'] = []
            with self.assertRaisesRegex(ValueError, 'category count'):
                errorbars.build(source, folder/'out.pptx', folder/'plan.ppttc', 'Name', bad, 1, None, 'tag')
            inventory.assert_not_called()
            self.assertEqual([p.name for p in folder.iterdir()], ['source.pptx'])

    def test_read_only_native_verifier_is_hash_bound(self):
        with tempfile.TemporaryDirectory() as td, patch.object(errorbars, 'verify_authentic_donor', return_value={'ok': True}) as verify:
            artifact = Path(td)/'native.pptx'; artifact.write_bytes(b'native')
            with self.assertRaisesRegex(ValueError, 'SHA-256 mismatch'):
                errorbars.verify_native(artifact, 'Name', 1, self.data, '0'*64)
            verify.assert_not_called()
            result = errorbars.verify_native(artifact, 'Name', 1, self.data, errorbars.sha(artifact))
            self.assertEqual(result['status'], 'ERRORBAR_MODEL_AND_CACHE_GATES_PASS')
            self.assertEqual(artifact.read_bytes(), b'native')

    def test_native_model_data_mismatch_fails_before_physical_cache_readback(self):
        selected = {'owner_name': 'Name', 'doc': {'slide_number': 1},
                    'frames': [{'native_chart_part': 'ppt/charts/chart7.xml', 'shape_tag': 'tag'}]}
        bad_actual = copy.deepcopy(self.data); bad_actual['series']['Min'][0] = 2
        with tempfile.TemporaryDirectory() as td, patch.object(errorbars.naming, 'inventory', return_value=([], [selected], [])), \
                patch.object(errorbars.naming, 'choose', side_effect=lambda candidates, **selector:
                             selected if selector.get('shape_tag') == 'tag' else self.fail('physical selector missing')), \
                patch.object(errorbars, 'donor_seed_data', return_value=bad_actual):
            path = Path(td)/'native.pptx'; path.write_bytes(b'native')
            with self.assertRaisesRegex(ValueError, 'model data differs'):
                errorbars.verify_authentic_donor(path, 'Name', 1, self.data)

    def test_numeric_cache_uses_point_indices_not_xml_order(self):
        ref = E.fromstring(b'<numRef><numCache><ptCount val="2"/><pt idx="1"><v>9</v></pt><pt idx="0"><v>4</v></pt></numCache></numRef>')
        self.assertEqual(errorbars._indexed_cache(ref), [4, 9])
        for bad in (b'<pt idx="0"><v>9</v></pt><pt idx="0"><v>4</v></pt>',
                    b'<pt idx="0"><v>9</v></pt><pt idx="2"><v>4</v></pt>',
                    b'<pt idx="0"><v>nan</v></pt><pt idx="1"><v>4</v></pt>'):
            ref = E.fromstring(b'<numRef><numCache><ptCount val="2"/>'+bad+b'</numCache></numRef>')
            with self.assertRaises(ValueError): errorbars._indexed_cache(ref)

    def test_wrong_error_bar_direction_cannot_reuse_valid_plus_cache(self):
        for bar_type in ('minus', 'both', ''):
            error = E.fromstring(('<errBars><errDir val="x"/><errValType val="cust"/>'
                                  '<errBarType val="'+bar_type+'"/><plus><numRef><f>Sheet1!A1</f>'
                                  '<numCache><ptCount val="1"/><pt idx="0"><v>3</v></pt></numCache>'
                                  '</numRef></plus></errBars>').encode())
            with self.assertRaisesRegex(ValueError, 'plus-only'):
                errorbars._error_cache(error)
        error.find('errBarType').set('val', 'plus')
        self.assertEqual(errorbars._error_cache(error), ('Sheet1!A1', [3]))
        E.SubElement(error, 'minus')
        with self.assertRaisesRegex(ValueError, 'minus error-bar extents'):
            errorbars._error_cache(error)

class RetentionGraderTests(unittest.TestCase):
    def test_duplicate_partitions_cannot_collapse_during_semantic_compare(self):
        identity = {'slide_number': 1, 'shape_tag': 'tag', 'model_type': 'CScatterChartSE',
                    'series_labels': ['A'], 'partitions': [{'series': 'A'}, {'series': 'A'}]}
        with self.assertRaisesRegex(RuntimeError, 'duplicate'):
            trendline.compare_semantics(identity, identity)

    def test_legend_bounds_use_absolute_tolerance(self):
        self.assertFalse(legend._close([1000000000.1], [1000000000], .05))

    def test_repeat_requires_geometry_retention_even_when_individual_grades_pass(self):
        first = {'status': 'PASS', 'physical_bounds': [0, 0, 1, 1],
                 'model_text': {'model_bounds': [0, 0, 1, 1], 'text_bounds': [0, 0, 1, 1]}}
        second = copy.deepcopy(first); second['model_text']['text_bounds'][0] = 1
        with patch.object(legend, 'grade', side_effect=[first, second]):
            result = legend.grade_repeat(Path('same'), Path('changed'), {})
        self.assertEqual(result['status'], 'FAIL')
        self.assertFalse(result['geometry_retained'])

if __name__ == '__main__': unittest.main()
