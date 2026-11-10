"""Private Gamma generation, model fidelity and preservation of earlier VBA."""
import json
import base64
import hashlib
import re
from fractions import Fraction
import math
import subprocess
import unittest
from public_integration import restore,EXTRA
from install_scaled_gamma_fidelity import HERE, ROOT, BEGIN, END, BASELINE, dataset, render
from install_extended_gamma_fidelity import strip


class ScaledGammaVbaTests(unittest.TestCase):
    def test_retained_excel_attempts_and_exact_source_binding(self):
        attempts = [
            ('initial', 'da14d50d648f806a6f85d7a964f7fe7a255bd949', 42511055298, 13629850422, 8142, 139,
             'cf4c98e4941c66b1e1ead91f9e9099336d38ad9a9c0e45111bc807377f30a4dc'),
            ('diagnostic', '1a3854badf088e4ed9a11d4fbac78d00c1358c51', 42511468913, 13629187789, 8142, 139,
             'e3bb3a08a885a10007396f9209be359944a57200fcfe406281800ea9656f2647'),
            ('bound', '974c6616a70947bc9f33b28938b48cc6f9e80473', 42511903287, 13630235299, 8142, 139,
             '7deb15ee7e51bfc24dffc05928086156f36e0b8aec5d07740c47892b5b4b32b1'),
            ('bits', '6754fe77b455c1ae92ec4a122b995092b16f9a70', 42512236419, 13630171461, 8142, 139,
             'bb2ecd88d88a3c7b5998666540f9a6f584a02fbbb6642a25b0d7038bb283007b'),
            ('', 'f0e8ef0e70f77e161f193be969b8ec3625580e10', 42512900145, 13629714744, 8190, 41,
             'b83fe11337dd961a73ec2b4c240625dc87c09ec008f111b9df061bb87efd91d4'),
        ]
        old = json.loads((HERE/'excel_scaled_dd_artifact.json').read_text())
        def failures(raw):
            text = raw.decode('utf-8-sig').replace('\r\n', '\n')
            return re.findall(r'^\[[^\]]+\]   (.*)$', text, re.M)
        baseline = failures(base64.b64decode(old['members']['test-result.txt']['base64']))
        self.assertEqual(len(baseline), 41)
        for label, sha, run, artifact_id, count, failed, zip_digest in attempts:
            with self.subTest(attempt=label):
                name = 'excel_scaled_gamma'+('_'+label if label else '')+'_artifact.json'
                artifact = json.loads((HERE/name).read_text())
                self.assertEqual(artifact['run_id'], run)
                self.assertEqual(artifact['artifact_id'], artifact_id)
                self.assertEqual(artifact['artifact_zip_sha256'], zip_digest)
                decoded = {}
                for member, entry in artifact['members'].items():
                    raw = base64.b64decode(entry['base64'], validate=True)
                    self.assertEqual(hashlib.sha256(raw).hexdigest(), entry['sha256'])
                    decoded[member] = raw
                record = json.loads(decoded['excel-certification.json'].decode('utf-8-sig'))
                self.assertEqual(record['candidate_sha'], sha)
                self.assertEqual(record['runner']['workflow']['sha'], sha)
                self.assertEqual(record['runner']['workflow']['run_id'], run)
                self.assertEqual(record['harness'], dict(entry_point='Test_STATS_PROBDIST_RunAll_CI',
                    assertions=count, passed=count-failed, failed=failed))
                for stage in ('import', 'compile', 'cleanup'):
                    self.assertEqual(record['stages'][stage]['status'], 'PASS')
                self.assertEqual(record['stages']['regression']['status'], 'FAIL')
                self.assertTrue(all(not grid['exported'] for grid in record['grids'].values()))
                self.assertEqual(record['log']['sha256'],
                    'sha256:'+hashlib.sha256(decoded['test-result.txt']).hexdigest())
                policy = json.loads(subprocess.check_output(
                    ['git', 'show', sha+':.github/excel-evidence-policy.json'], cwd=ROOT))
                self.assertEqual(policy['expected_assertions'], count)
                self.assertEqual({s['path'] for s in record['sources']}, set(policy['regression_sources']))
                for source in record['sources']:
                    blob = subprocess.check_output(['git','show',sha+':'+source['path']],cwd=ROOT)
                    self.assertEqual('sha256:'+hashlib.sha256(blob).hexdigest(),source['sha256'])
                    if not label:
                        self.assertEqual(strip((ROOT/source['path']).read_text()),blob.decode().replace('\r\n','\n'))
                details = failures(decoded['test-result.txt'])
                self.assertEqual(len(details), failed)
                self.assertEqual([s for s in details if not s.startswith('SG ')], baseline)
                self.assertEqual(sum(s.startswith('SG ') for s in details), failed-41)

    def test_dataset_reproduces(self):
        saved = json.loads((HERE/'scaled_gamma_fidelity.json').read_text())
        self.assertEqual(saved, dataset())
        self.assertEqual(len(saved['rows']),129)
        self.assertEqual(saved['additional_assertions'],475)
        self.assertEqual(len(saved['normalization']),48)
        self.assertEqual(len(saved['invalid']),8)
        self.assertEqual(len(saved['seams']),8)

    def test_generated_source_and_policy(self):
        text = restore((ROOT/'tests/M_STATS_PROBDIST_TEST.bas').read_text())
        self.assertEqual(text.count(BEGIN),1)
        block = BEGIN+text.split(BEGIN)[1].split(END)[0]+END+'\n\n'
        self.assertEqual(block,render(dataset()))
        self.assertNotIn('@COEFFICIENTS@',block)
        self.assertNotIn('@LN2@',block)
        self.assertEqual(text.count('    Test_ScaledGamma_Fidelity\n'),1)
        policy = json.loads((ROOT/'.github/excel-evidence-policy.json').read_text())
        extended = json.loads((HERE/'extended_gamma_fidelity.json').read_text())
        wide = json.loads((HERE/'wide_gamma_fidelity.json').read_text())
        self.assertEqual(policy['expected_assertions'],BASELINE+475+extended['additional_assertions']+wide['additional_assertions']+EXTRA)

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
        text = restore((ROOT/'tests/M_STATS_PROBDIST_TEST.bas').read_text())
        before,rest = text.split(BEGIN)
        _,after = rest.split(END+'\n\n')
        restored = (before+after).replace('    Test_ScaledGamma_Fidelity\n','')
        self.assertEqual(strip(restored),baseline.replace('\r\n','\n'))
        for path in (ROOT/'src').glob('*.bas'):
            original = subprocess.check_output(['git','show',sha+':'+path.relative_to(ROOT).as_posix()],cwd=ROOT).decode()
            self.assertEqual(restore(path.read_text()),original.replace('\r\n','\n'))

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
