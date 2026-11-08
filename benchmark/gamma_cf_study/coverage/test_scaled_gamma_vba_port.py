"""Private Gamma generation, model fidelity and preservation of earlier VBA."""
import json
from fractions import Fraction
import math
import subprocess
import unittest
from install_scaled_gamma_fidelity import HERE, ROOT, BEGIN, END, BASELINE, dataset, render


class ScaledGammaVbaTests(unittest.TestCase):
    def test_dataset_reproduces(self):
        saved = json.loads((HERE/'scaled_gamma_fidelity.json').read_text())
        self.assertEqual(saved, dataset())
        self.assertEqual(len(saved['rows']),129)
        self.assertEqual(saved['additional_assertions'],475)
        self.assertEqual(len(saved['normalization']),48)
        self.assertEqual(len(saved['invalid']),8)
        self.assertEqual(len(saved['seams']),8)

    def test_generated_source_and_policy(self):
        text = (ROOT/'tests/M_STATS_PROBDIST_TEST.bas').read_text()
        self.assertEqual(text.count(BEGIN),1)
        block = BEGIN+text.split(BEGIN)[1].split(END)[0]+END+'\n\n'
        self.assertEqual(block,render(dataset()))
        self.assertNotIn('@COEFFICIENTS@',block)
        self.assertNotIn('@LN2@',block)
        self.assertEqual(text.count('    Test_ScaledGamma_Fidelity\n'),1)
        policy = json.loads((ROOT/'.github/excel-evidence-policy.json').read_text())
        self.assertEqual(policy['expected_assertions'],BASELINE+475)

    def test_normalization_oracles_preserve_exact_dyadic_values(self):
        for row in dataset()['normalization']:
            h,l,eh,el = [float.fromhex(row[k]) for k in ('h','l','expected_h','expected_l')]
            exact = (Fraction(h)+Fraction(l))*2**row['e']
            scale = Fraction(2)**row['expected_e']
            self.assertEqual(exact,(Fraction(eh)+Fraction(el))*scale)
            self.assertEqual(eh,math.ldexp(float(exact),-row['expected_e']))
            self.assertGreaterEqual(abs(eh),.5)
            self.assertLess(abs(eh),1.)
            self.assertEqual(float(Fraction(eh)+Fraction(el)),eh)

    def test_earlier_vba_suite_and_public_sources_are_unchanged(self):
        sha = 'd94b70e7a98c78de9c0dd6fcfa2be997b3bf96af'
        baseline = subprocess.check_output(['git','show',sha+':tests/M_STATS_PROBDIST_TEST.bas'],cwd=ROOT).decode()
        text = (ROOT/'tests/M_STATS_PROBDIST_TEST.bas').read_text()
        before,rest = text.split(BEGIN)
        _,after = rest.split(END+'\n\n')
        restored = (before+after).replace('    Test_ScaledGamma_Fidelity\n','')
        self.assertEqual(restored,baseline.replace('\r\n','\n'))
        for path in (ROOT/'src').glob('*.bas'):
            original = subprocess.check_output(['git','show',sha+':'+path.relative_to(ROOT).as_posix()],cwd=ROOT).decode()
            self.assertEqual(path.read_text(),original.replace('\r\n','\n'))

    def test_revealed_boundaries_are_present(self):
        rows = dataset()['rows']
        for row in json.loads((HERE/'references.json').read_text()):
            a,x = float.fromhex(row['a_hex']),float.fromhex(row['x_hex'])
            if 0 < a <= 1 and 0 < x <= 32:
                self.assertTrue(any(r['a']==a.hex() and r['x']==x.hex() for r in rows))
        self.assertTrue(any(r['a']==(2.**-1074).hex() for r in rows))
        self.assertTrue(any(r['x']==(2.**-1074).hex() for r in rows))


if __name__ == '__main__':
    unittest.main()
