"""Ensure discovery completeness and honest failures without executing fixtures."""
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

SCRIPTS = Path(__file__).resolve().parents[1] / "skills/thinkcell-edit/scripts"
sys.path.insert(0, str(SCRIPTS))
import run_portable_tests as runner


class PortableRunnerTests(unittest.TestCase):
    def test_missing_private_fixtures_are_explicit_skips(self):
        with patch.object(runner, "_execute", return_value={"status": "PASS", "unittest_count": 2}) as execute:
            result = runner.run()
        self.assertEqual(execute.call_count, 1 + len(runner.PORTABLE_CHECKS))
        self.assertEqual(result["fixture_checks_skipped"], len(runner.FIXTURE_CHECKS))
        self.assertFalse(result["native_execution"])

    def test_subprocess_failures_and_timeouts_fail_report(self):
        with patch.object(runner.subprocess, "run", return_value=subprocess.CompletedProcess([], 1, "", "broken")):
            self.assertEqual(runner._execute("check", [], 1)["status"], "FAIL")
        with patch.object(runner.subprocess, "run", side_effect=subprocess.TimeoutExpired([], 1)):
            self.assertEqual(runner._execute("check", [], 1)["status"], "FAIL")
        with patch.object(runner, "_execute", return_value={"status": "FAIL"}):
            self.assertEqual(runner.run()["status"], "FAIL")

    def test_fixture_scripts_are_subprocesses_and_globals_are_remapped(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for _, names in runner.FIXTURE_CHECKS:
                for name in names.values():
                    (root / name).touch()
            with patch.object(runner, "_execute", return_value={"status": "PASS"}) as execute:
                result = runner.run(root)
            self.assertEqual(result["fixture_checks_skipped"], 0)
            commands = [call.args[1] for call in execute.call_args_list]
            self.assertTrue(any("--fixture" in command for command in commands))
            wrappers = [command[-1] for command in commands if "-c" in command]
            self.assertTrue(any("axis-break-ordinary-38764.pptx" in code for code in wrappers))
            self.assertTrue(all("r.testsRun" in code for code in wrappers))

    def test_unittest_summary_counts_and_assertions_enabled(self):
        process = subprocess.CompletedProcess([], 0, "", "Ran 7 tests in 0.02s\nOK (skipped=2)")
        with patch.object(runner.subprocess, "run", return_value=process):
            result = runner._execute("check", [], 1)
        self.assertEqual((result["unittest_count"], result["unittest_skipped"]), (7, 2))
        with patch.dict(runner.os.environ, {"PYTHONOPTIMIZE": "2"}):
            self.assertNotIn("PYTHONOPTIMIZE", runner._environment())


if __name__ == "__main__":
    unittest.main()
