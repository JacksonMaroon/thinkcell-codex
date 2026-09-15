import copy
import importlib.util
from pathlib import Path
import unittest

SCRIPT = Path(__file__).parents[1] / 'skills/thinkcell-edit/scripts/placement_plan.py'
SPEC = importlib.util.spec_from_file_location('placement_plan', SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)


class PlacementTests(unittest.TestCase):
    def setUp(self):
        self.context = {'source_sha256': 'a' * 64, 'slide_id': 256,
                        'width': 720, 'height': 405,
                        'shapes': [{'id': 2, 'left': 20, 'top': 20, 'width': 680, 'height': 40},
                                   {'id': 3, 'left': 20, 'top': 90, 'width': 240, 'height': 230}],
                        'reserved': [{'left': 0, 'top': 375, 'width': 720, 'height': 30}]}
        self.brief = {'inferred_chart_type': 'bar',
                      'rationale': 'Compare the categories in the supplied data.'}

    def plan(self, **kwargs):
        return MODULE.plan(self.context, {**self.brief, **kwargs})

    def test_no_geometry_guidance_preserves_content_and_rails(self):
        original = copy.deepcopy(self.context)
        p = self.plan()
        self.assertEqual(original, self.context)
        self.assertGreaterEqual(p['frame']['left'], 266)
        self.assertLessEqual(p['frame']['top'] + p['frame']['height'], 369)
        self.assertEqual(p['preserve_shape_ids'], [2, 3])
        self.assertFalse(p['native_execution'])

    def test_side_only(self):
        p = self.plan(side='right')
        self.assertGreaterEqual(p['frame']['left'], 360)

    def test_requested_type_overrides_inferred(self):
        p = self.plan(requested_chart_type='waterfall')
        self.assertEqual(p['chart_type'], 'waterfall')
        self.assertEqual(p['provenance']['chart_type'], 'explicit')

    def test_full_rectangle_exact(self):
        frame = {'left': 300, 'top': 90, 'width': 300, 'height': 200}
        self.assertEqual(self.plan(frame=frame)['frame'], frame)

    def test_explicit_small_dimensions_are_preserved_for_review(self):
        frame = {'left': 300, 'top': 90, 'width': 50, 'height': 30}
        self.assertEqual(self.plan(frame=frame)['frame'], frame)
        with self.assertRaises(ValueError): self.plan(frame=frame, min_width=72)

    def test_explicit_canvas_edge_overrides_default_margin(self):
        self.context['shapes'] = []; self.context['reserved'] = []
        frame = {'left': 0, 'top': 0, 'width': 720, 'height': 405}
        self.assertEqual(self.plan(frame=frame)['frame'], frame)
        with self.assertRaises(ValueError): self.plan(frame=frame, margin=12)

    def test_explicit_full_width_overrides_default_margin(self):
        self.context['shapes'] = []; self.context['reserved'] = []
        p = self.plan(frame={'width': 720})
        self.assertEqual(p['frame']['left'], 0)
        self.assertEqual(p['frame']['width'], 720)
        self.assertEqual(p['frame']['top'], 12)

    def test_explicit_frame_cannot_exceed_canvas(self):
        self.context['shapes'] = []; self.context['reserved'] = []
        with self.assertRaises(ValueError):
            self.plan(frame={'left': 700, 'top': 0, 'width': 100, 'height': 100})

    def test_partial_width_and_x(self):
        p = self.plan(frame={'left': 300, 'width': 200})
        self.assertEqual(p['frame']['left'], 300)
        self.assertEqual(p['frame']['width'], 200)
        self.assertEqual(p['frame']['top'], 66)

    def test_partial_height_and_top(self):
        p = self.plan(frame={'top': 100, 'height': 150})
        self.assertEqual(p['frame']['top'], 100)
        self.assertEqual(p['frame']['height'], 150)

    def test_collision_explicit_fails(self):
        with self.assertRaises(ValueError):
            self.plan(frame={'left': 20, 'top': 100, 'width': 400, 'height': 200})

    def test_conflicting_side_fails(self):
        with self.assertRaises(ValueError):
            self.plan(side='left', frame={'left': 400})

    def test_rail_collision_fails(self):
        with self.assertRaises(ValueError):
            self.plan(frame={'left': 300, 'top': 320, 'width': 250, 'height': 60})

    def test_invalid_geometry_fails(self):
        for value in (float('nan'), float('inf'), -1, True):
            with self.subTest(value=value), self.assertRaises(ValueError):
                self.plan(frame={'left': value})

    def test_no_space_fails(self):
        self.context['shapes'] = [{'id': 1, 'left': 0, 'top': 0, 'width': 720, 'height': 405}]
        with self.assertRaises(ValueError): self.plan()

    def test_no_semantic_guess_or_data_invention(self):
        with self.assertRaises(ValueError): MODULE.plan(self.context, {})
        p = self.plan()
        self.assertEqual(p['rationale'], self.brief['rationale'])
        self.assertNotIn('data', p)
        self.assertEqual(p['source_sha256'], 'a' * 64)

    def test_native_horizontal_rule_blocks_frame(self):
        self.context['shapes'].append({'id': 4, 'left': 270, 'top': 200, 'width': 400, 'height': 0})
        with self.assertRaises(ValueError):
            self.plan(frame={'left': 300, 'top': 150, 'width': 200, 'height': 100})

    def test_unknown_guidance_is_not_silently_dropped(self):
        with self.assertRaises(ValueError): self.plan(position='right')

    def test_unbound_plan_fails(self):
        self.context['source_sha256'] = 'unknown'
        with self.assertRaises(ValueError): self.plan()

    def test_nonobstacle_shapes_still_preserved(self):
        self.context['all_shapes'] = self.context['shapes'] + [{'id': 10, 'visible': 0}]
        self.assertEqual(self.plan()['preserve_shape_ids'], [2, 3, 10])


if __name__ == '__main__':
    unittest.main()
