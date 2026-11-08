"""Reproducible primitive oracle and private VBA source/fixture binding."""
import json
import math
import struct
import unittest
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
        text = (ROOT / 'tests/M_STATS_PROBDIST_TEST.bas').read_text()
        self.assertEqual(text.count(BEGIN), 1)
        self.assertEqual(text.count(END), 1)
        block = BEGIN+text.split(BEGIN)[1].split(END)[0]+END+'\n\n'
        self.assertEqual(block, render(data))
        self.assertEqual(text.count(TYPES), 1)
        self.assertEqual(text.count('    Test_ScaledProduct_Fidelity\n'), 1)
        policy = json.loads((ROOT / '.github/excel-evidence-policy.json').read_text())
        self.assertEqual(policy['expected_assertions'], 7104)

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
