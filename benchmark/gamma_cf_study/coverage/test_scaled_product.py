"""Exact-rational exponent/rounding regression; not Gamma holdout evidence."""
import math
import random
import struct
import sys
import unittest
from fractions import Fraction as F
from scaled_product import product, rounded


def pow2(e):
    return F(2**e) if e >= 0 else F(1, 2**-e)


def exact(pair):
    hi, lo, e = pair
    return (F(hi) + F(lo)) * pow2(e)


class ScaledProductTests(unittest.TestCase):
    def check_pair(self, a, b):
        pair = product(a, b)
        reference = F(a) * F(b)
        self.assertEqual(exact(pair), reference, (a.hex(), b.hex()))
        if a != 0 and b != 0:
            self.assertTrue(.5 <= abs(pair[0]) < 1)
            self.assertEqual(pair[0] + pair[1], pair[0])
        try:
            expected = float(reference)
        except OverflowError:
            with self.assertRaises(OverflowError):
                rounded(*pair)
        else:
            got = rounded(*pair)
            self.assertEqual(got, expected, (a.hex(), b.hex()))
            if got == 0:
                self.assertEqual(math.copysign(1, got),
                                 math.copysign(1, a) * math.copysign(1, b))

    def test_retained_unscaled_counterexamples(self):
        self.check_pair(2.**1000, 2.**-1000)
        self.check_pair(2.**-1022, math.nextafter(.5, 1.))
        pair = product(2.**-1022, math.nextafter(.5, 1.))
        self.assertNotEqual(exact(pair), F(rounded(*pair)))

    def test_every_input_exponent(self):
        partners = [.5, math.nextafter(.5, 1.), 1., math.nextafter(1., 2.),
                    -1.5, 2.**-1022, 2.**-1074, sys.float_info.max]
        for e in range(-1074, 1024):
            a = math.ldexp(1., e)
            for b in partners:
                self.check_pair(a, b)

    def test_random_finite_bit_patterns(self):
        rng = random.Random(24220261108)
        def operand():
            while True:
                x = struct.unpack('>d', rng.getrandbits(64).to_bytes(8, 'big'))[0]
                if math.isfinite(x):
                    return x
        for _ in range(8192):
            self.check_pair(operand(), operand())

    def test_subnormal_halfway_residual_and_parity(self):
        # 2.5 subnormal units: high-only rounding gives 2; +lo must give 3.
        for hi, exponent in [(.625, -1072), (.875, -1072), (.5, -1074)]:
            for lo in [0., 2.**-56, -2.**-56]:
                for sign in [-1., 1.]:
                    pair = sign * hi, sign * lo, exponent
                    self.assertEqual(rounded(*pair), float(exact(pair)))
        self.assertEqual(rounded(.625, 2.**-56, -1072), 3 * 2.**-1074)
        self.assertEqual(math.ldexp(.625, -1072), 2 * 2.**-1074)

    def test_normal_boundary_and_overflow(self):
        for a in [sys.float_info.max, math.nextafter(sys.float_info.max, 0.),
                  2.**-1022, math.nextafter(2.**-1022, 0.)]:
            for b in [math.nextafter(1., 0.), 1., math.nextafter(1., 2.), 2.]:
                self.check_pair(a, b)
        self.check_pair(sys.float_info.max, sys.float_info.max)
        self.check_pair(2.**-1074, 2.**-1074)

    def test_zero_and_underflow_signs(self):
        for a in [0., -0., 2.**-1074, -2.**-1074]:
            for b in [0., -0., .25, -.25, sys.float_info.max]:
                self.check_pair(a, b)

    def test_fail_closed_invalid_inputs(self):
        for v in [math.nan, math.inf, -math.inf]:
            for a, b in [(v, 1.), (1., v), (v, 0.)]:
                with self.assertRaises(ValueError):
                    product(a, b)
            with self.assertRaises(ValueError):
                rounded(v, 0., 0)
            with self.assertRaises(ValueError):
                rounded(.5, v, 0)
        for pair in [(0., 1., 0), (1., 0., 0), (.5, .1, 0),
                     (.5, 0., 1.5), (.5, 0., True)]:
            with self.assertRaises(ValueError):
                rounded(*pair)


if __name__ == '__main__':
    unittest.main()
