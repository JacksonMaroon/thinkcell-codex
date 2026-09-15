"""Exercise the PowerShell state comparator using JSON fixtures, never Office."""
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

SCRIPTS = Path(__file__).parents[1] / 'skills/thinkcell-edit/scripts'
sys.path.insert(0, str(SCRIPTS))
from runtime import powershell, powershell_env


@unittest.skipUnless(os.name == 'nt', 'Windows PowerShell comparator')
class OfficeStateTests(unittest.TestCase):
    def test_inspector_empty_notes_normalization_preserves_whitespace(self):
        def quote(path): return "'" + str(path).replace("'", "''") + "'"
        # Load only this pure function's AST, never the native inspector entrypoint.
        command = (r'''$tokens=$null;$errors=$null;
$ast=[System.Management.Automation.Language.Parser]::ParseFile(''' +
                   quote(SCRIPTS / 'inspect_layout.ps1') + r''',[ref]$tokens,[ref]$errors);
$fn=$ast.Find({param($node) $node-is[System.Management.Automation.Language.FunctionDefinitionAst]-and$node.Name-eq'Normalize-NoteRuns'},$true);
if($null-eq$fn){throw 'Normalizer not found'}
. ([scriptblock]::Create($fn.Extent.Text));
$runs=@([ordered]@{text='';font='Arial';size=12});
$empty=@(Normalize-NoteRuns '' $runs);
$missing=@(Normalize-NoteRuns $null $runs);
$whitespace=@(Normalize-NoteRuns ' ' $runs);
$text=@(Normalize-NoteRuns 'Note' $runs);
[ordered]@{empty=$empty;missing=$missing;whitespace=$whitespace;text=$text}|ConvertTo-Json -Depth 7 -Compress
''')
        result = subprocess.run([powershell(), '-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-Command', command],
                                capture_output=True, text=True, env=powershell_env(), timeout=30)
        self.assertEqual(result.returncode, 0, result.stderr)
        output = json.loads(result.stdout)
        self.assertEqual(output['empty'], [])
        self.assertEqual(output['missing'], [])
        self.assertEqual(output['whitespace'], output['text'])
        self.assertEqual(len(output['text']), 1)

    def test_wait_is_immediate_when_clean_and_bounded_when_changed(self):
        def quote(path): return "'" + str(path).replace("'", "''") + "'"
        # Override sleep locally: verify requested delays without waiting or Office.
        command = ('. ' + quote(SCRIPTS / 'office_state.ps1') + r''';
function Start-Sleep { param([int]$Milliseconds) $script:sleeps++; if($Milliseconds-ne250){throw 'Unexpected delay'} }
$before='[{"path":"Saved.pptx","saved":-1,"windows":1,"slides":1}]'
$changed='[{"path":"Saved.pptx","saved":0,"windows":1,"slides":1}]'
$script:reads=0;$script:sleeps=0
$clean=Wait-OfficeState $before { $script:reads++; return $before }
$cleanReads=$script:reads;$cleanSleeps=$script:sleeps
$script:reads=0;$script:sleeps=0
$settled=Wait-OfficeState $before { $script:reads++; if($script:reads-lt3){return $changed}; return $before }
$settledReads=$script:reads;$settledSleeps=$script:sleeps
$script:reads=0;$script:sleeps=0
$unsettled=Wait-OfficeState $before { $script:reads++; return $changed }
[ordered]@{clean=$clean;cleanReads=$cleanReads;cleanSleeps=$cleanSleeps;
 settled=$settled;settledReads=$settledReads;settledSleeps=$settledSleeps;
 unsettled=$unsettled;unsettledReads=$script:reads;unsettledSleeps=$script:sleeps} | ConvertTo-Json -Compress
''')
        result = subprocess.run([powershell(), '-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-Command', command],
                                capture_output=True, text=True, env=powershell_env(), timeout=30)
        self.assertEqual(result.returncode, 0, result.stderr)
        output = json.loads(result.stdout)
        self.assertEqual((output['cleanReads'], output['cleanSleeps']), (1, 0))
        self.assertEqual((output['settledReads'], output['settledSleeps']), (3, 2))
        self.assertEqual((output['unsettledReads'], output['unsettledSleeps']), (21, 20))
        self.assertEqual(json.loads(output['clean'])[0]['saved'], -1)
        self.assertEqual(json.loads(output['settled'])[0]['saved'], -1)
        self.assertEqual(json.loads(output['unsettled'])[0]['saved'], 0)

    def test_transients_only_are_excluded(self):
        def entry(path='Presentation23', saved=-1, windows=0):
            return {'path': path, 'saved': saved, 'windows': windows, 'slides': 1}
        visible = entry('Visible.pptx', windows=1)
        cases = [
            {'name': 'saved hidden unnamed addition', 'before': [visible], 'after': [visible, entry()], 'equal': True},
            {'name': 'saved hidden unnamed removal', 'before': [visible, entry()], 'after': [visible], 'equal': True},
            {'name': 'visible unnamed retained', 'before': [], 'after': [entry(windows=1)], 'equal': False},
            {'name': 'unsaved hidden unnamed retained', 'before': [], 'after': [entry(saved=0)], 'equal': False},
            {'name': 'named hidden saved retained', 'before': [], 'after': [entry('Saved.pptx')], 'equal': False},
            {'name': 'path prefix is not unnamed', 'before': [], 'after': [entry('folder/Presentation23')], 'equal': False},
            {'name': 'non-numeric suffix retained', 'before': [], 'after': [entry('Presentation23x')], 'equal': False},
            {'name': 'visible modified slide count retained', 'before': [visible], 'after': [{**visible, 'slides': 2}], 'equal': False},
            {'name': 'empty states equal', 'before': [], 'after': [], 'equal': True},
        ]
        def quote(path): return "'" + str(path).replace("'", "''") + "'"
        with tempfile.TemporaryDirectory() as td:
            fixture = Path(td) / 'cases.json'; fixture.write_text(json.dumps(cases), encoding='utf-8')
            command = ('. ' + quote(SCRIPTS / 'office_state.ps1') + '; '
                       '$cases=Get-Content -LiteralPath ' + quote(fixture) + ' -Raw | ConvertFrom-Json; '
                       '$results=@(foreach($case in $cases){'
                       '$before=ConvertTo-Json -InputObject @($case.before) -Depth 7 -Compress; '
                       '$after=ConvertTo-Json -InputObject @($case.after) -Depth 7 -Compress; '
                       '[ordered]@{name=$case.name;equal=(Compare-OfficeState $before $after)}}); '
                       'ConvertTo-Json -InputObject $results -Depth 7 -Compress')
            result = subprocess.run([powershell(), '-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-Command', command],
                                    capture_output=True, text=True, env=powershell_env(), timeout=30)
            self.assertEqual(result.returncode, 0, result.stderr)
            actual = json.loads(result.stdout)
            self.assertEqual(len(actual), len(cases))
            for wanted, got in zip(cases, actual):
                with self.subTest(case=wanted['name']): self.assertEqual(got['equal'], wanted['equal'])


if __name__ == '__main__':
    unittest.main()
