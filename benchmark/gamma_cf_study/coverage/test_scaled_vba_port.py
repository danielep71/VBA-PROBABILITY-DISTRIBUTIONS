"""Reproducible primitive oracle and private VBA source/fixture binding."""
import base64
import hashlib
import json
import re
import subprocess
import math
import struct
import unittest
from public_integration import restore,EXTRA,live_total
from frozen_tree import FROZEN
from fractions import Fraction as F
from install_scaled_fidelity import HERE, ROOT, BEGIN, END, TYPES, dataset, render, power, literal
from scaled_product import product, rounded


class ScaledVbaPortTests(unittest.TestCase):
    def test_oracle_artifact_reproduces(self):
        data = json.loads((HERE / 'scaled_product_fidelity.json').read_text())
        self.assertEqual(data, dataset())
        self.assertEqual(len(data['products']), 216)
        self.assertEqual(data['additional_assertions'], 4657)
        for row in data['products']:
            a, b, h, l = (float.fromhex(row[k]) for k in ['a', 'b', 'hi', 'lo'])
            self.assertEqual((F(h)+F(l))*power(row['exponent']), F(a)*F(b))
            got = product(a, b)
            self.assertEqual(got, (h, l, row['exponent']))
            if row['rounded_ok']:
                self.assertEqual(rounded(*got).hex(), row['rounded'])
            else:
                with self.assertRaises(OverflowError):
                    rounded(*got)

    def test_generated_block_and_policy_match(self):
        data = dataset()
        text = restore((FROZEN/'tests/M_STATS_PROBDIST_TEST.bas').read_text())
        self.assertEqual(text.count(BEGIN), 1)
        self.assertEqual(text.count(END), 1)
        block = BEGIN+text.split(BEGIN)[1].split(END)[0]+END+'\n\n'
        self.assertEqual(block, render(data))
        self.assertEqual(text.count(TYPES), 1)
        self.assertEqual(text.count('    Test_ScaledProduct_Fidelity\n'), 1)
        extension = json.loads((HERE / 'scaled_dd_fidelity.json').read_text())
        gamma = json.loads((HERE / 'scaled_gamma_fidelity.json').read_text())
        extended = json.loads((HERE/'extended_gamma_fidelity.json').read_text())
        wide = json.loads((HERE/'wide_gamma_fidelity.json').read_text())
        self.assertEqual(live_total(), 7104+extension['additional_assertions']+gamma['additional_assertions']+extended['additional_assertions']+wide['additional_assertions']+EXTRA)

    def test_retained_excel_identity_and_failure_details(self):
        artifact = json.loads((HERE / 'excel_scaled_product_artifact.json').read_text())
        decoded = {}
        for name, entry in artifact['members'].items():
            raw = base64.b64decode(entry['base64'], validate=True)
            self.assertEqual(hashlib.sha256(raw).hexdigest(), entry['sha256'])
            decoded[name] = raw
        record = json.loads(decoded['excel-certification.json'].decode('utf-8-sig'))
        sha = '86a7e47be621b32350fb343aa8afaa212eddddb9'
        self.assertEqual(record['candidate_sha'], sha)
        self.assertEqual(record['runner']['workflow']['sha'], sha)
        self.assertEqual(record['runner']['workflow']['run_id'], 42476278844)
        self.assertEqual(record['harness'], dict(entry_point='Test_STATS_PROBDIST_RunAll_CI',
                                                assertions=7104, passed=7063, failed=41))
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
        old = json.loads((HERE / 'excel_precision_v2_artifact.json').read_text())
        old_log = base64.b64decode(old['members']['test-result.txt']['base64'])
        def failures(raw):
            text = raw.decode('utf-8-sig').replace('\r\n', '\n')
            return re.findall(r'^\[[^\]]+\]   (.*)$', text, re.M)
        previous, current = failures(old_log), failures(decoded['test-result.txt'])
        self.assertEqual(len(current), 41)
        self.assertEqual(current, previous)
        self.assertFalse(any(label.startswith('SP ') for label in current))

    def test_bit_literals_roundtrip_including_negative_zero(self):
        for x in [0., -0., 1., -1., 2.**-1074, 2.**-1022,
                  float.fromhex('0x1.fffffffffffffp+1023'), math.nextafter(.5, 1.)]:
            text = literal(x.hex())
            words = [int(t.split('&H')[1].split('&')[0], 16)
                     for t in text.split(',')]
            restored = struct.unpack('>d', struct.pack('>II', *words))[0]
            self.assertEqual(restored.hex(), x.hex())
            # Force Long literals: &HFFFF alone could be a signed 16-bit Integer.
            self.assertEqual(text.count('&H'), 2)
            self.assertEqual(text.count('&'), 4)

    def test_rounding_cases_have_exact_rational_references(self):
        for row in dataset()['rounding']:
            h, l = float.fromhex(row['hi']), float.fromhex(row['lo'])
            exact = (F(h)+F(l))*power(row['exponent'])
            value = float(exact)
            if value == 0:
                value = math.copysign(0., h)
            self.assertEqual(value.hex(), row['rounded'])
            self.assertEqual(rounded(h, l, row['exponent']).hex(), row['rounded'])


if __name__ == '__main__':
    unittest.main()
