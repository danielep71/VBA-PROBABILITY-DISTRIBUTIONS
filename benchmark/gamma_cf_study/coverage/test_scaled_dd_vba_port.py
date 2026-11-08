"""Scaled arithmetic fixture identity, independent error screen and source isolation."""
import base64
import hashlib
import re
import json
import math
import subprocess
import unittest
from fractions import Fraction as F
from install_scaled_dd_fidelity import HERE, ROOT, BEGIN, END, BASELINE, dataset, render
from install_scaled_fidelity import power


def exact(words):
    return (F(float.fromhex(words[0]))+F(float.fromhex(words[1])))*power(words[2])


class ScaledDDVbaTests(unittest.TestCase):
    def test_retained_excel_identity_and_failure_details(self):
        artifact = json.loads((HERE / 'excel_scaled_dd_artifact.json').read_text())
        decoded = {}
        for name, entry in artifact['members'].items():
            raw = base64.b64decode(entry['base64'], validate=True)
            self.assertEqual(hashlib.sha256(raw).hexdigest(), entry['sha256'])
            decoded[name] = raw
        record = json.loads(decoded['excel-certification.json'].decode('utf-8-sig'))
        sha = '86a365f5f7caa57ba0deff166549dfd357a244d5'
        self.assertEqual(record['candidate_sha'], sha)
        self.assertEqual(record['runner']['workflow']['sha'], sha)
        self.assertEqual(record['runner']['workflow']['run_id'], 42494438289)
        self.assertEqual(record['harness'], dict(entry_point='Test_STATS_PROBDIST_RunAll_CI',
                                                assertions=7715, passed=7674, failed=41))
        self.assertEqual(record['stages']['compile']['status'], 'PASS')
        self.assertEqual(record['stages']['regression']['status'], 'FAIL')
        self.assertEqual(record['stages']['cleanup']['status'], 'PASS')
        self.assertTrue(all(not grid['exported'] for grid in record['grids'].values()))
        self.assertEqual(record['log']['sha256'],
                         'sha256:'+hashlib.sha256(decoded['test-result.txt']).hexdigest())
        policy = json.loads(subprocess.check_output(
            ['git', 'show', sha+':.github/excel-evidence-policy.json'], cwd=ROOT))
        self.assertEqual({s['path'] for s in record['sources']}, set(policy['regression_sources']))
        for source in record['sources']:
            blob = subprocess.check_output(['git', 'show', sha+':'+source['path']], cwd=ROOT)
            self.assertEqual('sha256:'+hashlib.sha256(blob).hexdigest(), source['sha256'])
        old = json.loads((HERE / 'excel_scaled_product_artifact.json').read_text())
        old_log = base64.b64decode(old['members']['test-result.txt']['base64'])
        def failures(raw):
            text = raw.decode('utf-8-sig').replace('\r\n', '\n')
            return re.findall(r'^\[[^\]]+\]   (.*)$', text, re.M)
        previous, current = failures(old_log), failures(decoded['test-result.txt'])
        self.assertEqual(len(current), 41)
        self.assertEqual(current, previous)
        self.assertFalse(any(label.startswith('SD ') for label in current))

    def test_dataset_reproduces(self):
        saved = json.loads((HERE/'scaled_dd_fidelity.json').read_text())
        self.assertEqual(saved, dataset())
        self.assertEqual(saved['additional_assertions'], 611)
        self.assertEqual(len(saved['rows']), 611)

    def test_model_outputs_against_independent_rationals(self):
        for row in dataset()['rows']:
            if not row['ok']:
                continue
            a, b = exact(row['a']), exact(row['b'])
            op = row['operation']
            reference = {1: lambda: a+b, 2: lambda: a-b,
                         3: lambda: a*b, 4: lambda: a/b}[op]()
            error = abs(exact(row['result'])-reference)
            self.assertLessEqual(error, abs(reference)*F(1, 2**98), row)

    def test_generated_source_and_policy(self):
        text = (ROOT/'tests/M_STATS_PROBDIST_TEST.bas').read_text()
        self.assertEqual(text.count(BEGIN), 1)
        block = BEGIN+text.split(BEGIN)[1].split(END)[0]+END+'\n\n'
        self.assertEqual(block, render(dataset()))
        self.assertEqual(text.count('    Test_ScaledDD_Fidelity\n'), 1)
        policy = json.loads((ROOT/'.github/excel-evidence-policy.json').read_text())
        extension = json.loads((HERE/'scaled_gamma_fidelity.json').read_text())
        extended = json.loads((HERE/'extended_gamma_fidelity.json').read_text())
        wide = json.loads((HERE/'wide_gamma_fidelity.json').read_text())
        self.assertEqual(policy['expected_assertions'], BASELINE+611+extension['additional_assertions']+extended['additional_assertions']+wide['additional_assertions'])

    def test_old_vba_suite_is_unchanged(self):
        baseline = subprocess.check_output(['git', 'show',
            '4c9f965f2bf8489828cb0367b400941dcbfcc82d:tests/M_STATS_PROBDIST_TEST.bas'], cwd=ROOT).decode()
        text = (ROOT/'tests/M_STATS_PROBDIST_TEST.bas').read_text()
        before, rest = text.split(BEGIN)
        _, after = rest.split(END+'\n\n')
        restored = (before+after).replace('    Test_ScaledDD_Fidelity\n', '')
        from install_scaled_gamma_fidelity import BEGIN as GB, END as GE
        prefix, remainder = restored.split(GB)
        _, suffix = remainder.split(GE+'\n\n')
        restored = (prefix+suffix).replace('    Test_ScaledGamma_Fidelity\n', '')
        from install_extended_gamma_fidelity import strip
        self.assertEqual(strip(restored), baseline.replace('\r\n', '\n'))

    def test_rejection_and_cancellation_coverage(self):
        data = dataset()['rows']
        self.assertEqual(sum(not row['ok'] for row in data), 23)
        self.assertTrue(any(row['ok'] and exact(row['result']) != 0
                            and abs(exact(row['result'])) < F(1, 2**2100)
                            for row in data))
        self.assertTrue(any(not math.isfinite(float.fromhex(row['a'][0]))
                            and not row['ok'] for row in data))


if __name__ == '__main__':
    unittest.main()
