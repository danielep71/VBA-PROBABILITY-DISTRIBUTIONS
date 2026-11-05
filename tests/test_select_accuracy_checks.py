"""Regression coverage for CI selection, reduced execution and the required gate."""
import contextlib
import io
import os
from pathlib import Path
import runpy
import subprocess
import sys
import tempfile
import textwrap
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'tools'))
import select_accuracy_checks as S  # noqa: E402


class SelectionTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.previous = Path.cwd()
        os.chdir(self.tmp.name)
        self.addCleanup(os.chdir, self.previous)
        self.git('init', '-q')
        self.git('config', 'user.name', 'Fixture')
        self.git('config', 'user.email', 'fixture@example.invalid')
        self.change('README.md', 'baseline')
        self.base = self.git('rev-parse', 'HEAD').strip()

    def git(self, *args):
        return subprocess.check_output(['git', *args], text=True, stderr=subprocess.PIPE)

    def change(self, path, value):
        p = Path(path)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(value)
        self.git('add', '-A')
        self.git('commit', '-qm', 'fixture change')

    def test_issue_metadata_skips(self):
        self.change('.github/labels.json', '{}')
        self.change('.github/ISSUE_TEMPLATE/bug.yml', 'name: bug')
        for event in ('push', 'pull_request'):
            self.assertFalse(S.select(event, self.base)[0])

    def test_unknown_and_numerical_inputs_run(self):
        for path in ('README.md', 'src/A.bas', 'benchmark/requirements-ci.txt',
                     'benchmark/positive_ratio_subnormal/gamma_holdout.csv',
                     '.github/workflows/accuracy-gate.yml', 'tools/new_helper.py',
                     '.gitattributes', 'docs/numerical_contract.md'):
            with self.subTest(path=path):
                self.assertFalse(S.metadata_only([path]))

    def test_complete_history_not_last_commit(self):
        self.change('benchmark/helper.py', 'changed')
        self.change('.github/labels.json', '{}')
        self.assertTrue(S.select('pull_request', self.base)[0])

    def test_reverted_relevant_change_still_runs(self):
        self.change('README.md', 'changed')
        self.change('README.md', 'baseline')
        self.change('.github/labels.json', '{}')
        self.assertTrue(S.select('push', self.base)[0])

    def test_rename_out_of_relevant_path_runs(self):
        self.git('mv', 'README.md', '.github-labels.tmp')
        self.git('commit', '-qm', 'rename')
        self.assertTrue(S.select('pull_request', self.base)[0])

    def test_deleted_relevant_path_runs(self):
        self.git('rm', 'README.md')
        self.git('commit', '-qm', 'delete')
        self.assertTrue(S.select('push', self.base)[0])

    def test_manual_schedule_empty_or_missing_base_run(self):
        for event, base in [('workflow_dispatch', self.base), ('schedule', self.base),
                            ('push', ''), ('push', '0' * 40), ('push', 'f' * 40),
                            ('pull_request', self.base)]:
            self.assertTrue(S.select(event, base)[0])

    def test_comparison_errors_run(self):
        with patch.object(S, 'git', side_effect=OSError('unavailable')):
            self.assertTrue(S.select('push', self.base)[0])

    def test_non_ancestral_comparison_runs(self):
        self.change('README.md', 'future')
        future = self.git('rev-parse', 'HEAD').strip()
        self.assertTrue(S.select('push', future, self.base)[0])


class ExecutionTests(unittest.TestCase):
    def execute_shim(self, args, status=0):
        calls = []
        def run(command, **kwargs):
            calls.append(command[1:])
            return subprocess.CompletedProcess(command, status)
        with patch.object(sys, 'argv', ['test_evidence_tools.py', *args]), \
             patch('subprocess.run', side_effect=run), contextlib.redirect_stdout(io.StringIO()):
            runpy.run_path(str(ROOT / 'benchmark/test_evidence_tools.py'), run_name='__main__')
        return calls

    def test_single_list_full_and_live_modes(self):
        full = self.execute_shim([])
        live = self.execute_shim(['--checks-only'])
        self.assertEqual(len(full), 45)
        expected = [c for c in full if Path(c[0]).name.startswith('check_')
                    and '--self-test' not in c]
        self.assertEqual(live, expected)
        self.assertIn(['check_manifest_provenance.py'], live)
        self.assertIn(['check_verification_depth.py'], live)
        self.assertIn(['check_module_lists.py'], live)
        self.assertLess(len(live), len(full))

    def test_live_failure_propagates(self):
        with self.assertRaises(SystemExit) as error:
            self.execute_shim(['--checks-only'], status=1)
        self.assertEqual(error.exception.code, 1)

    def test_final_gate_rejects_missing_failed_or_cancelled_jobs(self):
        workflow = (ROOT / '.github/workflows/accuracy-gate.yml').read_text()
        script = textwrap.dedent(workflow.rsplit('        run: |\n', 1)[1])
        for full in ('true', 'false', ''):
            for selection in ('success', 'failure', 'cancelled', 'skipped'):
                for core in ('success', 'failure', 'cancelled', 'skipped'):
                    for study in ('success', 'failure', 'cancelled', 'skipped'):
                        env = dict(os.environ, FULL_PROOFS=full, SELECTION=selection,
                                   CORE=core, STUDY=study)
                        result = subprocess.run(['bash', '-e', '-c', script], env=env)
                        expected = selection == core == 'success' and (
                            (full == 'true' and study == 'success') or
                            (full == 'false' and study == 'skipped'))
                        self.assertEqual(result.returncode == 0, expected,
                                         (full, selection, core, study))


if __name__ == '__main__':
    unittest.main()
