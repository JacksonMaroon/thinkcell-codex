"""Portable preservation/geometry checks. No Office process is started."""
import copy
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch
import zipfile
from PIL import Image

SCRIPTS = Path(__file__).parents[1] / 'skills/thinkcell-edit/scripts'
sys.path.insert(0, str(SCRIPTS))
import chart_geometry as geometry
import insert_chart as insertion


def chart(legend=False, family='CSequenceChartSE'):
    def anchors(values, start=0):
        return [{'grid_id': str(i + start), 'side': side, 'value': value,
                 'old_string': str(value * 8)}
                for i, (side, value) in enumerate(zip(geometry.SIDES, values))]
    return {'source_sha256': 'A' * 64, 'name': 'Chart', 'family': family,
            'plot': anchors([10, 20, 110, 100]), 'bundle': [0, 0, 120, 140],
            'legends': [anchors([20, 110, 100, 130], 4)] if legend else []}


class GeometryTests(unittest.TestCase):
    def test_margins_and_legend_in_outer_frame(self):
        frame = {'left': 300, 'top': 100, 'width': 250, 'height': 200}
        with patch.object(geometry, 'inspect_geometry', return_value=chart(True)):
            plan = geometry.make_plan('unused', frame)
        self.assertEqual(plan['expected_plot'], [310, 120, 540, 260])
        legend = [e['new_value'] / 8 for e in plan['edits'][4:]]
        self.assertEqual(legend, [320, 270, 400, 290])
        self.assertEqual(len({e['grid_id'] for e in plan['edits']}), 8)

    def test_tiny_plot_rejected_even_with_positive_outer_frame(self):
        with patch.object(geometry, 'inspect_geometry', return_value=chart()):
            with self.assertRaises(ValueError):
                geometry.make_plan('unused', {'left': 0, 'top': 0, 'width': 30, 'height': 100})

    def test_pie_plot_stays_square(self):
        with patch.object(geometry, 'inspect_geometry', return_value=chart(family='CPieChartSE')):
            p = geometry.make_plan('unused', {'left': 0, 'top': 0, 'width': 300, 'height': 200})
        l, t, r, b = p['expected_plot']
        self.assertEqual(r - l, b - t)

    def test_nonfinite_frame_rejected(self):
        for value in (float('nan'), float('inf'), True):
            with self.subTest(value=value), patch.object(geometry, 'inspect_geometry', return_value=chart()):
                with self.assertRaises(ValueError):
                    geometry.make_plan('unused', {'left': 0, 'top': 0, 'width': value, 'height': 200})

    def test_modified_plan_rejected_before_any_writer(self):
        expected = {'outer_frame': {'left': 0, 'top': 0, 'width': 200, 'height': 200}, 'edits': []}
        tampered = copy.deepcopy(expected); tampered['edits'] = [{'new_value': 10}]
        with tempfile.TemporaryDirectory() as td, \
                patch.object(geometry, 'inspect_geometry', return_value={'document': {}}), \
                patch.object(geometry, 'make_plan', return_value=expected), \
                patch.object(geometry.subprocess, 'run') as writer:
            with self.assertRaises(ValueError):
                geometry.prepare(Path(td) / 'source.pptx', Path(td) / 'output.pptx', tampered)
            writer.assert_not_called()

    def test_final_bundle_overflow_rejected(self):
        g = chart(); frame = {'left': 0, 'top': 0, 'width': 120, 'height': 130}
        p = {'automation_name': 'Chart', 'expected_plot': [10, 20, 110, 100], 'outer_frame': frame}
        with patch.object(geometry, 'inspect_geometry', return_value=g):
            with self.assertRaises(ValueError): geometry.verify('unused', p)

    def test_malformed_expected_plot_rejected_before_inspection(self):
        for expected in ([], [1, 2, 3], [1, 2, 3, float('nan')], [1, 2, 3, True], '1234'):
            with self.subTest(expected=expected), patch.object(geometry, 'inspect_geometry') as reader:
                with self.assertRaises(ValueError): geometry.verify('unused', {'expected_plot': expected})
                reader.assert_not_called()

    def test_invalid_tolerance_rejected_before_inspection(self):
        for tolerance in (-1, 3, float('nan'), float('inf'), True):
            with self.subTest(tolerance=tolerance), patch.object(geometry, 'inspect_geometry') as reader:
                with self.assertRaises(ValueError): geometry.verify('unused', {}, tolerance)
                reader.assert_not_called()

    def test_horizontal_inward_reflow_reports_exact_adjustment(self):
        g = chart(); g['outer'] = copy.deepcopy(g['plot'])
        g['plot'][2]['value'] -= 3.125; g['plot'][3]['value'] -= 4
        p = {'automation_name': 'Chart', 'expected_plot': [10, 20, 110, 100],
             'outer_frame': {'left': 0, 'top': 0, 'width': 120, 'height': 140}}
        with patch.object(geometry, 'inspect_geometry', return_value=g):
            result = geometry.verify('unused', p)
        self.assertTrue(result['native_inward_reflow_accepted'])
        self.assertEqual(result['native_plot_adjustment_points'],
                         {'left': 0, 'top': 0, 'right': -3.125, 'bottom': -4})

    def test_inward_reflow_is_narrow_and_preserves_minimum_plot(self):
        p = {'automation_name': 'Chart', 'expected_plot': [10, 20, 110, 100],
             'outer_frame': {'left': 0, 'top': 0, 'width': 120, 'height': 140}}
        for name, side, delta, horizontal in [('outward', 2, 2, True), ('large', 2, -4.1, True),
                                              ('left', 0, -2, True), ('top', 1, 2, True),
                                              ('vertical', 2, -3, False)]:
            g = chart(); g['outer'] = copy.deepcopy(g['plot']) if horizontal else []
            g['plot'][side]['value'] += delta
            with self.subTest(name=name), patch.object(geometry, 'inspect_geometry', return_value=g):
                with self.assertRaises(ValueError): geometry.verify('unused', p)
        g = chart(); g['outer'] = copy.deepcopy(g['plot']); g['plot'][2]['value'] = 43
        tiny = {**p, 'expected_plot': [10, 20, 46, 100]}
        with patch.object(geometry, 'inspect_geometry', return_value=g):
            with self.assertRaises(ValueError): geometry.verify('unused', tiny)


class DonorTypeTests(unittest.TestCase):
    def test_zero_hole_string_is_pie_and_positive_hole_is_doughnut(self):
        from lxml import etree
        c = {'exact': True, 'owner': etree.Element('CPieChartSE'), 'owner_name': 'Chart'}
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / 'donor.pptx'; path.write_bytes(b'mocked fixture')
            for hole, expected in [('0', 'pie'), (0, 'pie'), (None, 'pie'), ('45', 'doughnut')]:
                with self.subTest(hole=hole), \
                        patch.object(insertion, 'inventory', return_value=([], [c], [])), \
                        patch.object(insertion, 'kind', return_value='pie'), \
                        patch.object(insertion, 'snapshot', return_value={'target_model': {'hole_percent': hole}}):
                    self.assertEqual(insertion.donor_type(path), expected)


class PixelTests(unittest.TestCase):
    def test_changes_inside_allowed_and_outside_detected(self):
        with tempfile.TemporaryDirectory() as td:
            before, after = Path(td) / 'before.png', Path(td) / 'after.png'
            source = Image.new('RGB', (200, 100), 'white'); source.save(before)
            changed = source.copy(); changed.putpixel((60, 40), (0, 0, 0)); changed.save(after)
            frame = {'left': 25, 'top': 10, 'width': 25, 'height': 25}
            self.assertTrue(insertion.outside_pixels(before, after, frame, 100))
            changed.putpixel((150, 80), (0, 0, 0)); changed.save(after)
            with self.assertRaises(ValueError): insertion.outside_pixels(before, after, frame, 100)

    def test_different_native_render_sizes_rejected(self):
        with tempfile.TemporaryDirectory() as td:
            a, b = Path(td) / 'a.png', Path(td) / 'b.png'
            Image.new('RGB', (200, 100)).save(a); Image.new('RGB', (201, 100)).save(b)
            with self.assertRaises(ValueError):
                insertion.outside_pixels(a, b, {'left': 0, 'top': 0, 'width': 10, 'height': 10}, 100)


class TargetPreflightTests(unittest.TestCase):
    def package(self, path, extra):
        p = insertion.NS['p']; a = insertion.NS['a']; r = insertion.NS['r']
        with zipfile.ZipFile(path, 'w') as z:
            z.writestr('slide.xml', f'<p:sld xmlns:p="{p}" xmlns:a="{a}" xmlns:r="{r}"><p:cSld><p:spTree/></p:cSld>{extra}</p:sld>')
            z.writestr('ppt/presentation.xml', f'<p:presentation xmlns:p="{p}"><p:sldSz cx="100" cy="100"/></p:presentation>')
            z.writestr('theme.xml', f'<a:theme xmlns:a="{a}"><a:themeElements/></a:theme>')

    def inspect(self, path):
        def rels(z, part):
            if part == 'slide.xml': return {'r1': {'type': '/slideLayout', 'resolved': 'layout.xml'}}
            if part == 'layout.xml': return {'r1': {'type': '/slideMaster', 'resolved': 'master.xml'}}
            return {'r1': {'type': '/theme', 'resolved': 'theme.xml'}}
        with patch.object(insertion, 'logical_slides', return_value=[{'part': 'slide.xml', 'id': 256}]), \
                patch.object(insertion, 'relationship_map', side_effect=rels):
            return insertion.package_info(path, target=True)

    def test_transition_and_animation_rejected(self):
        with tempfile.TemporaryDirectory() as td:
            for element in ('transition', 'timing'):
                path = Path(td) / (element + '.pptx'); self.package(path, f'<p:{element}/>')
                with self.subTest(element=element), self.assertRaises(ValueError): self.inspect(path)

    def test_plain_target_preflight_is_read_only(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / 'target.pptx'; self.package(path, '')
            before = path.read_bytes(); result = self.inspect(path)
            self.assertEqual(result['slide_id'], 256)
            self.assertEqual(before, path.read_bytes())


if __name__ == '__main__':
    unittest.main()
