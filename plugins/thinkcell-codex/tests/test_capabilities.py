"""Route discovery must not turn a missing or research adapter into success."""
from pathlib import Path
import copy
import json
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

SCRIPTS = Path(__file__).resolve().parents[1] / 'skills/thinkcell-edit/scripts'
sys.path.insert(0, str(SCRIPTS))
import capabilities
import chart_geometry_cli


class CapabilityTests(unittest.TestCase):
    def test_registry_covers_every_documented_matrix_entry(self):
        lines = (SCRIPTS.parent / 'references/capability-matrix.md').read_text().split('## Public route stages')[0].splitlines()
        titles = [line.split('|')[1].strip() for line in lines if line.startswith('| ')][1:]
        actual = {entry['title'] for entry in capabilities.load_registry()['capabilities']}
        self.assertTrue(set(titles).issubset(actual))

    def test_missing_route_and_traversal_are_rejected(self):
        for value in ('../thinkcell.py', 'C:/thinkcell.py', '/tmp/thinkcell.py', 'gantt\\run.py', 'missing.py'):
            with self.assertRaises(ValueError):
                capabilities.packaged_path(value)

    def test_symlink_outside_package_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / 'scripts'
            root.mkdir()
            outside = Path(directory) / 'outside.py'
            outside.write_text('')
            link = root / 'link.py'
            try:
                link.symlink_to(outside)
            except OSError:
                self.skipTest('Creating symlinks is unavailable')
            with self.assertRaisesRegex(ValueError, 'escapes'):
                capabilities.packaged_path('link.py', root)

    def test_research_route_cannot_execute(self):
        with patch.object(capabilities.subprocess, 'run') as run:
            with self.assertRaisesRegex(ValueError, 'no packaged execution route'):
                capabilities.dispatch('gantt-dependency-reflow', ['--execute'])
            run.assert_not_called()

    def test_dispatch_preserves_arguments_and_child_failure(self):
        with patch.object(capabilities.subprocess, 'run', return_value=subprocess.CompletedProcess([], 7)) as run:
            code = capabilities.dispatch('multi-chart-update', ['--', '--input', 'space name.pptx', '--plan', 'p.json'])
            self.assertEqual(code, 7)
            command = run.call_args.args[0]
            self.assertNotIn('--execute', command)
            self.assertEqual(command[-4:], ['--input', 'space name.pptx', '--plan', 'p.json'])

    def test_duplicate_ids_and_invalid_status_fail_registry(self):
        original = capabilities.load_registry()
        variants = [copy.deepcopy(original), copy.deepcopy(original)]
        variants[0]['capabilities'][1]['id'] = variants[0]['capabilities'][0]['id']
        variants[1]['capabilities'][0]['status'] = 'certified_everywhere'
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'registry.json'
            for variant in variants:
                path.write_text(json.dumps(variant))
                with self.assertRaises(ValueError):
                    capabilities.load_registry(path)

    def test_discovery_does_not_open_office_or_dispatch(self):
        with patch.object(capabilities.subprocess, 'run') as run:
            result = capabilities.discover('gantt-taskbar')
            self.assertFalse(result['native_certification_renewed'])
            self.assertEqual(result['capabilities'][0]['route']['stage'], 'preparation')
            run.assert_not_called()

    def test_unknown_capability_fails(self):
        with self.assertRaisesRegex(ValueError, 'Unknown capability'):
            capabilities.discover('gantt-anything')


class GeometryCliTests(unittest.TestCase):
    def test_input_and_existing_output_are_preserved(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'input.json'
            source.write_text('original')
            for path in (source, source.parent / 'sub' / '..' / 'input.json'):
                with self.assertRaises(ValueError):
                    chart_geometry_cli.fresh_output(path, [source])
            self.assertEqual(source.read_text(), 'original')

    def test_wrong_source_hash_rejects_before_plan_or_output(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / 'source.pptx'
            source.write_bytes(b'original')
            selection = root / 'selector.json'
            selection.write_text('{}')
            frame = root / 'frame.json'
            frame.write_text('{}')
            out = root / 'plan.json'
            argv = ['geometry', 'make-plan', '--input', str(source), '--expected-sha256', '0'*64,
                    '--selection-json', str(selection), '--frame-json', str(frame), '--plan-out', str(out)]
            with patch.object(sys, 'argv', argv), patch.object(chart_geometry_cli.geometry, 'make_plan') as make:
                self.assertEqual(chart_geometry_cli.main(), 1)
                make.assert_not_called()
            self.assertFalse(out.exists())
            self.assertEqual(source.read_bytes(), b'original')


if __name__ == '__main__':
    unittest.main()
