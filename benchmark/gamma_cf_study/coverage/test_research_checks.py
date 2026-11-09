"""Domain, normalization and negative fixtures for supplementary bounds."""
import math
import unittest
from fractions import Fraction
import mpmath as mp
from research_checks import pinelis, diagnostic_contains, truncation_bound


class ResearchChecksTests(unittest.TestCase):
    def test_domains_fail_closed(self):
        for a, x in [(0, 1), (-1, 1), (1.01, 1), (.5, 0), (.5, math.inf)]:
            with self.assertRaises(ValueError):
                pinelis(a, x)

    def test_normalization_and_endpoint(self):
        with mp.workdps(100):
            for a, x in [(.01, .3), (.39, 1.39), (.5, 2), (.99, 20)]:
                lo, hi = pinelis(a, x)
                q = mp.gammainc(a, x, mp.inf) / mp.gamma(a)
                self.assertLess(lo, q)
                self.assertLess(q, hi)
            lo, hi = pinelis(1, 2)
            self.assertEqual(lo, mp.exp(-2))
            self.assertEqual(hi, lo)

    def test_bounds_reject_large_error_but_not_known_one_ulp_defect(self):
        with mp.workdps(100):
            for xh, qh in [('0x1.63d70a3d70a41p+0', '0x1.1897def153c27p-4'),
                           ('0x1.63d70a3d70a42p+0', '0x1.1897def153c28p-4')]:
                bounds = pinelis(.39, float.fromhex(xh))
                self.assertTrue(diagnostic_contains(float.fromhex(qh), bounds))
                self.assertFalse(diagnostic_contains(float(bounds[1]) * 2, bounds))

    def test_exact_truncation_bound(self):
        self.assertGreater(truncation_bound(), 0)
        self.assertLess(truncation_bound(), Fraction(1949, 10**34))


if __name__ == '__main__':
    unittest.main()
