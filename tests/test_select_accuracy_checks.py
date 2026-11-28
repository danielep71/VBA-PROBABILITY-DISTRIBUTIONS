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


class GitFixture(unittest.TestCase):
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


class SelectionTests(GitFixture):
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
        self.assertEqual(len(full), 140)
        self.assertIn(["gamma_route_study/test_pair_candidate.py"], full)
        self.assertIn(["gamma_route_study/test_scalar_candidate.py"], full)
        self.assertIn(["gamma_cache_study/test_cache_vba.py"], full)
        self.assertIn(["gamma_recurrence_study/test_prepare.py"], full)
        self.assertIn(['gamma_route_study/test_binary64_map.py'], full)
        self.assertIn(['test_harness_studies.py'], full)
        self.assertIn(['test_harness_timing.py'], full)
        self.assertIn(['gamma_lower_q/test_promote_x05_ladder.py'], full)
        self.assertIn(['gamma_lower_q/test_gamma_lower_q.py'], full)
        self.assertIn(['gamma_cf_study/test_promote_x2_ladder.py'], full)
        for name in ('test_precision_probe.py', 'test_precision_v2.py',
                     'test_research_checks.py', 'test_precision_arithmetic_audit.py',
                     'test_cf_enclosure.py', 'test_scaled_product.py',
                     'test_scaled_vba_port.py', 'test_scaled_dd.py', 'test_scaled_dd_vba_port.py', 'test_scaled_gamma_vba_port.py', 'test_extended_gamma.py', 'test_wide_gamma.py', 'test_public_integration.py', 'test_public_fresh.py'):
            command = ['gamma_cf_study/coverage/' + name]
            self.assertIn(command, full)
            self.assertNotIn(command, live)
        self.assertIn(["gamma_cf_study/replay_baseline.py", "isolation"], full)
        self.assertIn(["gamma_cf_study/replay_baseline.py", "candidate"], full)
        self.assertIn(["positive_ratio_subnormal/check_post23_cutoffs.py"], full)
        expected = [c for c in full if Path(c[0]).name.startswith('check_')
                    and '--self-test' not in c
                    or c == ['gamma_cf_study/replay_baseline.py', 'candidate']]
        self.assertEqual(live, expected)
        self.assertIn(['gamma_cf_study/replay_baseline.py', 'candidate'], live)
        self.assertIn(['check_manifest_provenance.py'], live)
        self.assertIn(['check_verification_depth.py'], live)
        self.assertIn(['check_module_lists.py'], live)
        self.assertIn(['holdout_older/check_holdout_older_summary.py'], live)
        self.assertLess(len(live), len(full))

    def test_live_failure_propagates(self):
        with self.assertRaises(SystemExit) as error:
            self.execute_shim(['--checks-only'], status=1)
        self.assertEqual(error.exception.code, 1)

    def test_final_gate_rejects_missing_failed_or_cancelled_jobs(self):
        workflow = (ROOT / '.github/workflows/accuracy-gate.yml').read_text()
        script = workflow.rsplit('        run: |\n', 1)[1].split('\n      - name:', 1)[0]
        script = textwrap.dedent(script + '\n')
        self.assertIn('test "$CORE" = success', script)
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


TREE = 'a' * 40
REPO = 'owner/project'


def attestation(run_id=7, **changes):
    item = {'name': S.ATTESTATION_PREFIX + TREE, 'expired': False,
            'workflow_run': {'id': run_id, 'repository_id': 1, 'head_repository_id': 1}}
    item.update(changes)
    return item


def gate_run(run_id=7, **changes):
    run = {'id': run_id, 'path': S.WORKFLOW_PATH, 'event': 'pull_request',
           'status': 'completed', 'conclusion': 'success',
           'repository': {'full_name': REPO}, 'head_repository': {'full_name': REPO}}
    run.update(changes)
    return run


def service(artifacts, runs):
    """A stand-in for the two API calls; records what was asked."""
    asked = []

    def fetch(url, token):
        asked.append(url)
        if '/actions/artifacts?' in url:
            return {'artifacts': artifacts}
        return runs[int(url.rsplit('/', 1)[1])]
    fetch.asked = asked
    return fetch


class VerifiedTreeTests(unittest.TestCase):
    """#386: full proofs are not repeated on an attested tree, and only then."""

    def verified(self, artifacts, runs, tree=TREE, repo=REPO):
        return S.verified_tree(tree, repo, fetch=service(artifacts, runs))[0]

    def test_valid_attestation_is_accepted(self):
        for event in ('pull_request', 'push'):
            self.assertTrue(self.verified([attestation()], {7: gate_run(event=event)}))

    def test_lookup_names_the_exact_tree(self):
        fetch = service([attestation()], {7: gate_run()})
        S.verified_tree(TREE, REPO, fetch=fetch)
        self.assertIn('name=' + S.ATTESTATION_PREFIX + TREE, fetch.asked[0])
        self.assertIn('/repos/' + REPO + '/', fetch.asked[0])

    def test_no_attestation_runs_full(self):
        self.assertFalse(self.verified([], {}))

    def test_other_tree_or_expired_attestation_runs_full(self):
        other = attestation(name=S.ATTESTATION_PREFIX + 'b' * 40)
        self.assertFalse(self.verified([other], {7: gate_run()}))
        self.assertFalse(self.verified([attestation(expired=True)], {7: gate_run()}))

    def test_run_that_did_not_succeed_runs_full(self):
        for change in ({'conclusion': 'failure'}, {'conclusion': 'cancelled'},
                       {'conclusion': None, 'status': 'in_progress'},
                       {'status': 'queued'}):
            self.assertFalse(self.verified([attestation()], {7: gate_run(**change)}), change)

    def test_other_workflow_or_event_runs_full(self):
        for change in ({'path': '.github/workflows/static-checks.yml'},
                       {'event': 'workflow_dispatch'}, {'event': 'schedule'},
                       {'event': 'pull_request_target'}, {'event': 'workflow_run'}):
            self.assertFalse(self.verified([attestation()], {7: gate_run(**change)}), change)

    def test_fork_attestation_runs_full(self):
        fork = attestation(workflow_run={'id': 7, 'repository_id': 1, 'head_repository_id': 2})
        self.assertFalse(self.verified([fork], {7: gate_run()}))
        foreign = gate_run(head_repository={'full_name': 'someone/fork'})
        self.assertFalse(self.verified([attestation()], {7: foreign}))
        self.assertFalse(self.verified([attestation()], {7: gate_run(head_repository=None)}))

    def test_run_identity_must_match(self):
        self.assertFalse(self.verified([attestation()], {7: gate_run(run_id=8)}))
        missing = attestation(workflow_run=None)
        self.assertFalse(self.verified([missing], {7: gate_run()}))

    def test_one_valid_attestation_among_invalid_ones_is_enough(self):
        artifacts = [attestation(run_id=5), attestation(run_id=7)]
        runs = {5: gate_run(run_id=5, conclusion='failure'), 7: gate_run()}
        self.assertTrue(self.verified(artifacts, runs))

    def test_malformed_input_runs_full_without_asking(self):
        for tree, repo in (('', REPO), ('HEAD', REPO), (TREE + '&per_page=1', REPO),
                           (TREE, ''), (TREE, 'owner/project/../other')):
            fetch = service([attestation()], {7: gate_run()})
            self.assertFalse(S.verified_tree(tree, repo, fetch=fetch)[0], (tree, repo))
            self.assertEqual(fetch.asked, [])

    def test_lookup_errors_run_full(self):
        for error in (OSError('network'), ValueError('json'), KeyError('artifacts'),
                      S.urllib.error.HTTPError('u', 403, 'rate limit', None, None)):
            def fetch(url, token, error=error):
                raise error
            self.assertFalse(S.verified_tree(TREE, REPO, fetch=fetch)[0], error)
        self.assertFalse(S.verified_tree(TREE, REPO, fetch=lambda u, t: {})[0])
        self.assertFalse(S.verified_tree(TREE, REPO, fetch=lambda u, t: [])[0])


class MainSelectionTests(GitFixture):
    """The command line: which events may use an attestation, and the outputs."""

    def run_main(self, event, verified):
        out, summary = Path('out.txt'), Path('summary.md')
        for path in (out, summary):
            path.unlink(missing_ok=True)
        calls = []

        def lookup(tree, repo, token, api):
            calls.append(tree)
            return verified, 'Tree was verified in full by run 7' if verified else 'No valid attestation for this tree'
        argv = ['select', '--event', event, '--base', self.base,
                '--output', str(out), '--summary', str(summary)]
        with patch.object(S, 'verified_tree', lookup), patch.object(sys, 'argv', argv), \
                patch.dict(os.environ, {'GITHUB_REPOSITORY': REPO}), \
                contextlib.redirect_stdout(io.StringIO()):
            S.main()
        values = dict(line.split('=', 1) for line in out.read_text().splitlines())
        return values, calls, summary.read_text()

    def test_attested_tree_selects_the_live_checks_only(self):
        self.change('benchmark/tool.py', 'relevant')
        tree = self.git('rev-parse', 'HEAD^{tree}').strip()
        for event in ('push', 'pull_request'):
            values, calls, summary = self.run_main(event, True)
            self.assertEqual(values, {'full': 'false', 'tree': tree})
            self.assertEqual(calls, [tree])
            self.assertIn('verified in full by run 7', summary)

    def test_unattested_tree_runs_full_and_reports_its_tree(self):
        self.change('benchmark/tool.py', 'relevant')
        tree = self.git('rev-parse', 'HEAD^{tree}').strip()
        values, calls, summary = self.run_main('push', False)
        self.assertEqual(values, {'full': 'true', 'tree': tree})
        self.assertIn('no valid attestation', summary)

    def test_manual_and_scheduled_runs_never_use_an_attestation(self):
        self.change('benchmark/tool.py', 'relevant')
        for event in ('workflow_dispatch', 'schedule', 'unknown'):
            values, calls, _ = self.run_main(event, True)
            self.assertEqual(values['full'], 'true', event)
            self.assertEqual(calls, [], event)

    def test_metadata_only_change_does_not_ask(self):
        self.change('.github/labels.json', '[]')
        values, calls, _ = self.run_main('push', True)
        self.assertEqual(values['full'], 'false')
        self.assertEqual(calls, [])


class AttestationWorkflowTests(unittest.TestCase):
    """The workflow writes an attestation only for a full, successful run."""

    def setUp(self):
        self.workflow = (ROOT / '.github/workflows/accuracy-gate.yml').read_text()
        self.final = self.workflow.split('\n  accuracy-gate:\n', 1)[1]

    def test_attestation_follows_the_final_requirement(self):
        requirement = self.final.index('Require every applicable check to succeed')
        record = self.final.index('Record the verified tree')
        upload = self.final.index('actions/upload-artifact@')
        self.assertLess(requirement, record)
        self.assertLess(record, upload)
        self.assertNotIn('upload-artifact', self.workflow.split('\n  accuracy-gate:\n', 1)[0])

    def test_attestation_steps_require_full_proofs(self):
        steps = self.final.split('\n      - name: ')[2:]
        self.assertEqual(len(steps), 2)
        for step in steps:
            self.assertIn("if: ${{ needs.select.outputs.full == 'true' "
                          "&& needs.select.outputs.tree != '' }}", step)
            self.assertNotIn('always()', step)
            self.assertNotIn('continue-on-error', step)
        self.assertIn('name: ' + S.ATTESTATION_PREFIX + '${{ needs.select.outputs.tree }}', steps[1])

    def test_only_the_select_job_reads_actions(self):
        self.assertEqual(self.workflow.count('actions: read'), 1)
        select = self.workflow.split('\n  select:\n', 1)[1].split('\n  accuracy-core:\n', 1)[0]
        self.assertIn('actions: read', select)
        self.assertIn('GH_TOKEN: ${{ github.token }}', select)
        self.assertNotIn(': write', self.workflow)
        self.assertEqual(S.WORKFLOW_PATH, '.github/workflows/accuracy-gate.yml')


if __name__ == '__main__':
    unittest.main()
