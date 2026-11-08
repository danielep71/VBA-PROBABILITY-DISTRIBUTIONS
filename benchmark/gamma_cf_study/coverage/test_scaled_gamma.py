"""Revealed integration screen, not a disjoint acceptance set."""
import math
import unittest
import mpmath as mp
from scaled_gamma import public


class ScaledGammaTests(unittest.TestCase):
    def test_small_shape_and_subnormal_arguments(self):
        with mp.workdps(160):
            for a in [2.**-1074, 2.**-1022, 2.**-400, .01, .5, 1.]:
                for x in [2.**-1074, 2.**-1022, .25, 1., a+1, 32.]:
                    got = public(a, x)
                    aa, xx = mp.mpf(a), mp.mpf(x)
                    expected = (float(mp.gammainc(aa, 0, xx, regularized=True)),
                                float(mp.gammainc(aa, xx, mp.inf, regularized=True)))
                    for value, ref in zip(got, expected):
                        self.assertTrue(0 <= value <= 1)
                        self.assertLessEqual(abs(value-ref), 2*math.ulp(ref), (a,x,got,expected))

    def test_domain_is_explicit(self):
        for a,x in [(0,1),(-1,1),(2,1),(1,0),(1,33),(math.inf,1),(1,math.nan)]:
            with self.assertRaises(ValueError):
                public(a,x)


if __name__ == '__main__':
    unittest.main()
