"""Scaled arithmetic fixture identity, independent error screen and source isolation."""
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
        self.assertEqual(policy['expected_assertions'], BASELINE+611)

    def test_old_vba_suite_is_unchanged(self):
        baseline = subprocess.check_output(['git', 'show',
            '4c9f965f2bf8489828cb0367b400941dcbfcc82d:tests/M_STATS_PROBDIST_TEST.bas'], cwd=ROOT).decode()
        text = (ROOT/'tests/M_STATS_PROBDIST_TEST.bas').read_text()
        before, rest = text.split(BEGIN)
        _, after = rest.split(END+'\n\n')
        restored = (before+after).replace('    Test_ScaledDD_Fidelity\n', '')
        self.assertEqual(restored, baseline.replace('\r\n', '\n'))

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
