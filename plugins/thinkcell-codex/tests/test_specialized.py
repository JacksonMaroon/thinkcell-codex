"""Portable specialized-contract tests; synthetic data only, no Office or donors."""
from contextlib import ExitStack
from pathlib import Path
import copy
import sys
import tempfile
import unittest
from unittest.mock import patch

SCRIPTS = Path(__file__).resolve().parents[1] / 'skills/thinkcell-edit/scripts'
sys.path[:0] = [str(SCRIPTS), str(SCRIPTS / 'thinkcell_no_click/implementation')]
from lxml import etree as E
import chart_semantics
import update_thinkcell_json as updater


def request(matrix, values, widths=None):
    model = {'categories': matrix[0][1:], 'series_names': [r[0] for r in matrix[2:]],
             'series_values': values}
    if widths is not None:
        model['column_widths'] = widths
    return {'matrix': matrix, 'expected_model': model}


def check_contract(kind, source, baseline, intended):
    """Mock only source acquisition; exercise real request and slot validation."""
    chart = {'owner': E.Element('CSequenceChartSE'), 'owner_name': 'Synthetic',
             'doc': {'streams': {('store', 'think-cellXML'): b'<root><PersistentType><m_eorient val="0"/></PersistentType></root>',
                                 ('store', 'Package'): b'synthetic'}}}
    cells = [{'row': i+1, 'column': j+1, 'value': value}
             for i, row in enumerate(source) for j, value in enumerate(row) if value is not None]
    with tempfile.TemporaryDirectory() as folder, ExitStack() as stack:
        path = Path(folder) / 'source.pptx'
        path.write_bytes(b'synthetic')
        for name, value in [('inventory', ([], [chart], [])), ('link_contract', {'storage': 'store'}),
                            ('read_blob', [{'nonempty_cells': cells}]), ('model_of', baseline)]:
            stack.enter_context(patch.object(updater, name, return_value=value))
        stack.enter_context(patch.object(chart_semantics, 'kind', return_value=kind))
        stack.enter_context(patch.object(chart_semantics, 'details', return_value={'kind': kind}))
        updater.validate_request(intended, 'CSequenceChartSE')
        return updater.canonical_sequence_contract(path, intended, name='Synthetic')


class WaterfallTests(unittest.TestCase):
    def setUp(self):
        self.source = [[None, 'Start', 'Step', 'Total'], [None]*4, ['Series', 100, 20, 'e']]
        self.base = request(self.source, [[100, 20, 120]])['expected_model']
        self.good = request(copy.deepcopy(self.source), [[100, 21, 121]])
        self.good['matrix'][2][2] = 21

    def test_literal_equals_uses_numeric_expected_total(self):
        result = check_contract('waterfall', self.source, self.base, self.good)
        self.assertEqual(result['family'], 'waterfall')
        self.assertEqual(self.good['matrix'][2][3], 'e')
        self.assertEqual(self.good['expected_model']['series_values'][0][2], 121)

    def test_formula_text_rejected_in_numeric_slot(self):
        self.good['matrix'][2][1] = '=1+1'
        self.good['expected_model']['series_values'][0][0] = '=1+1'
        with self.assertRaisesRegex(ValueError, 'finite numbers or null'):
            check_contract('waterfall', self.source, self.base, self.good)

    def test_boolean_equals_expectation_rejected(self):
        self.good['expected_model']['series_values'][0][2] = True
        with self.assertRaisesRegex(ValueError, 'finite numbers or null'):
            check_contract('waterfall', self.source, self.base, self.good)

    def test_equals_cannot_be_replaced_by_fixed_total(self):
        self.good['matrix'][2][3] = 121
        with self.assertRaisesRegex(ValueError, 'equals slots must stay fixed'):
            check_contract('waterfall', self.source, self.base, self.good)

    def test_new_equals_slot_rejected(self):
        self.good['matrix'][2][2] = 'e'
        with self.assertRaisesRegex(ValueError, 'equals slots must stay fixed'):
            check_contract('waterfall', self.source, self.base, self.good)


class MekkoTests(unittest.TestCase):
    def setUp(self):
        self.source = [[None, 'A', 'B'], [None, 300, 200], ['Series', 20, 15]]
        self.base = request(self.source, [[20, 15]], [300, 200])['expected_model']
        self.good = request(copy.deepcopy(self.source), [[20, 15]], [310, 200])
        self.good['matrix'][1][1] = 310

    def test_units_width_changes_without_scaling_heights(self):
        result = check_contract('mekko-units', self.source, self.base, self.good)
        self.assertEqual(result['family'], 'mekko-units')
        self.assertEqual(self.good['expected_model']['series_values'], self.base['series_values'])

    def test_nonpositive_widths_rejected(self):
        for width in (0, -1):
            with self.subTest(width=width):
                self.good['matrix'][1][1] = width
                self.good['expected_model']['column_widths'][0] = width
                with self.assertRaisesRegex(ValueError, 'positive expected column_widths'):
                    check_contract('mekko-units', self.source, self.base, self.good)

    def test_width_matrix_and_expectation_must_agree(self):
        self.good['expected_model']['column_widths'][0] = 311
        with self.assertRaisesRegex(ValueError, 'X extent row must match'):
            check_contract('mekko-units', self.source, self.base, self.good)

    def test_negative_height_rejected(self):
        self.good['matrix'][2][1] = -20
        self.good['expected_model']['series_values'][0][0] = -20
        with self.assertRaisesRegex(ValueError, 'nonnegative absolute data'):
            check_contract('mekko-units', self.source, self.base, self.good)

    def test_percent_width_must_equal_segment_total(self):
        source = [[None, 'A', 'B'], [None]*3, ['Series', 20, 15]]
        baseline = request(source, [[20, 15]], [20, 15])['expected_model']
        intended = request(copy.deepcopy(source), [[20, 15]], [20, 15])
        self.assertEqual(check_contract('mekko-percent', source, baseline, intended)['family'], 'mekko-percent')
        intended['expected_model']['column_widths'][0] = 21
        with self.assertRaisesRegex(ValueError, 'widths must equal absolute category totals'):
            check_contract('mekko-percent', source, baseline, intended)


if __name__ == '__main__':
    unittest.main()

