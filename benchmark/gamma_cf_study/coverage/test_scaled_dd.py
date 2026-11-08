"""Exact-rational arithmetic screen; empirical guard, not a Gamma holdout."""
import math
import random
import sys
import unittest
from fractions import Fraction as F
from scaled_dd import ScaledDD as S, MAX_EXPONENT, ALIGNMENT_LIMIT


def power(e):
    return F(2**e) if e >= 0 else F(1, 2**-e)


def exact(x):
    return (F(x.hi)+F(x.lo))*power(x.exponent)


class ScaledDDTests(unittest.TestCase):
    def compare(self, got, reference):
        self.assertEqual(got.hi+got.lo, got.hi)
        if reference == 0:
            self.assertEqual(exact(got), 0)
        else:
            self.assertLessEqual(abs(exact(got)-reference), abs(reference)*F(1, 2**98))

    def test_rational_operations_across_exponents(self):
        rng = random.Random(24232767)
        for i in range(4096):
            def operand():
                h = rng.uniform(.5, 1.)
                l = math.ulp(h)*rng.uniform(-.49, .49)
                return S(h, l, rng.randrange(-2100, 2101))
            a, b = operand(), operand()
            if i % 2 == 0:
                b = S(b.hi, b.lo, a.exponent+rng.randrange(-120, 121))
            if rng.randrange(2):
                b = -b
            aa, bb = exact(a), exact(b)
            for got, reference in [(a+b, aa+bb), (a-b, aa-bb),
                                   (a*b, aa*bb), (a/b, aa/bb)]:
                self.compare(got, reference)

    def test_cancellation_and_tiny_low_components(self):
        for e in [-2100, -1022, 0, 1024, 2100]:
            for k in [54, 55, 106, 500, 1022, 1074]:
                for sign in [-1, 1]:
                    a, b = S(.75, sign*2.**-k, e), S(.75, 0., e)
                    self.assertEqual(exact(a-b), exact(a)-exact(b))
                    self.assertEqual(exact(a+-a), 0)
                    self.compare(a/b, exact(a)/exact(b))
            a = S(math.nextafter(1., 0.), 2.**-55, e)
            b = S(.5, -2.**-55, e+1)
            self.compare(a-b, exact(a)-exact(b))

    def test_alignment_boundary_and_omission_bound(self):
        for gap in [0, 1, 53, 54, 106, 119, 120, 121, 122, 1024, 4200]:
            for sign in [-1, 1]:
                a, b = S(.5, 2.**-55, 2100), S(sign*.75, sign*2.**-55, 2100-gap)
                reference = exact(a)+exact(b)
                self.compare(a+b, reference)
                if gap > ALIGNMENT_LIMIT:
                    self.assertEqual(a+b, a)
                    self.assertLess(abs(exact(a+b)-reference), abs(reference)*F(1, 2**119))

    def test_accumulated_operations(self):
        # A revealed arithmetic stress test, not an error theorem or holdout.
        for exponent in [-2100, 0, 2100]:
            total = S(.75, 2.**-56, exponent)
            reference = exact(total)
            factor = S(.5, 2.**-56, 1)
            for i in range(256):
                term = S((-1 if i % 2 else 1)*.625, 0., exponent-(i % 121))
                total = total*factor+term
                reference = reference*exact(factor)+exact(term)
                self.compare(total, reference)

    def test_extreme_products_and_quotients_stay_scaled(self):
        maximum = S.from_float(sys.float_info.max)
        minimum = S.from_float(2.**-1074)
        for a, b in [(maximum, maximum), (minimum, minimum), (maximum, minimum)]:
            self.compare(a*b, exact(a)*exact(b))
            self.compare(a/b, exact(a)/exact(b))
        self.assertEqual((maximum/maximum).rounded(), 1.)
        self.assertEqual((minimum/minimum).rounded(), 1.)
        self.assertEqual((minimum*minimum).rounded(), 0.)
        with self.assertRaises(OverflowError):
            (maximum*maximum).rounded()

    def test_zero_signs_and_division_by_zero(self):
        for z in [0., -0.]:
            for x in [2., -2.]:
                self.assertEqual((S.from_float(z)*x).rounded().hex(), (z*x).hex())
                self.assertEqual((S.from_float(z)/x).rounded().hex(), (z/x).hex())
            with self.assertRaises(ZeroDivisionError):
                S.from_float(1)/S.from_float(z)
        self.assertEqual((S.from_float(-0.)+S.from_float(-0.)).rounded().hex(), '-0x0.0p+0')

    def test_invalid_inputs_and_exponent_overflow_fail_closed(self):
        for args in [(math.inf, 0., 0), (.5, math.nan, 0), (1., 0., 0),
                     (.5, .1, 0), (0., 0., 1), (.5, 0., True),
                     (.5, 0., MAX_EXPONENT+1)]:
            with self.assertRaises(ValueError):
                S(*args)
        with self.assertRaises(ValueError):
            S(.75, 0., MAX_EXPONENT)*S.from_float(4.)
        with self.assertRaises(ValueError):
            S(.75, 0., -MAX_EXPONENT)/S.from_float(4.)
        with self.assertRaises(ValueError):
            S.from_float(math.nan)


if __name__ == '__main__':
    unittest.main()
