"""Exact-rational audit of frozen v2 primitives; empirical bounds, not proofs."""
import math
import random
import unittest
from fractions import Fraction as F
from precision_probe import DD


def exact(v):
    return F(v.hi) + F(v.lo)


class ArithmeticAudit(unittest.TestCase):
    def test_sum_eft_including_gradual_underflow(self):
        tiny = math.ldexp(1., -1074)
        cases = [(1., 2**-53), (1., -1.), (tiny, tiny),
                 (math.ldexp(1., -1022), -tiny), (2**400, 2**-400)]
        for a, b in cases:
            for x, y in [(a, b), (b, a)]:
                s, e = DD.sum(x, y)
                self.assertEqual(F(s) + F(e), F(x) + F(y))

    def test_rational_error_and_normalization(self):
        rng = random.Random(273241242)
        for _ in range(2048):
            a = DD(math.ldexp(rng.uniform(1, 2), rng.randrange(-400, 401)))
            b = DD(math.ldexp(rng.uniform(1, 2), rng.randrange(-400, 401)))
            a.lo = math.ulp(a.hi) * rng.uniform(-.49, .49)
            b.lo = math.ulp(b.hi) * rng.uniform(-.49, .49)
            if rng.randrange(2):
                b = -b
            aa, bb = exact(a), exact(b)
            for got, ref in [(a+b, aa+bb), (a*b, aa*bb), (a/b, aa/bb)]:
                self.assertTrue(math.isfinite(got.hi) and math.isfinite(got.lo))
                self.assertEqual(float(exact(got)), got.hi)
                # Deliberately conservative empirical guard, not R19's bound.
                self.assertLessEqual(abs(exact(got)-ref), abs(ref)*F(1, 2**98))

    def test_cancellation_keeps_low_components(self):
        a, b = DD(1., 2**-80), DD(-1., 2**-90)
        self.assertEqual(exact(a+b), F(2**-80)+F(2**-90))
        self.assertEqual(exact(a + -a), 0)

    def test_exponent_limits_are_real_counterexamples(self):
        # Finite exact product, but the unscaled Dekker splitter overflows.
        p, e = DD.product(math.ldexp(1., 1000), math.ldexp(1., -1000))
        self.assertEqual(p, 1.)
        self.assertFalse(math.isfinite(e))
        # Exact residual is below binary64: two Double cannot retain it.
        a, b = math.ldexp(1., -1022), math.nextafter(.5, 1.)
        p, e = DD.product(a, b)
        self.assertNotEqual(F(p)+F(e), F(a)*F(b))
        self.assertLessEqual(abs(F(p)+F(e)-F(a)*F(b)), F(1, 2**1075))


if __name__ == '__main__':
    unittest.main()
