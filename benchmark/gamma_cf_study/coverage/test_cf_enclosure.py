"""Domain, exact-rational arithmetic and independent-oracle enclosure checks."""
import math
import unittest
from fractions import Fraction
import mpmath as mp
from cf_enclosure import PositiveIntervals, enclose


class EnclosureTests(unittest.TestCase):
    def test_directed_interval_arithmetic(self):
        ar = PositiveIntervals(17)
        for a, b in [(.1, .3), (1e-100, 1.39), (math.nextafter(1., 0.), .5)]:
            for interval, ref in [(ar.add(ar.point(a), ar.point(b)), Fraction(a)+Fraction(b)),
                                  (ar.mul(ar.point(a), ar.point(b)), Fraction(a)*Fraction(b)),
                                  (ar.div(ar.point(a), ar.point(b)), Fraction(a)/Fraction(b))]:
                self.assertLessEqual(Fraction(interval[0]), ref)
                self.assertLessEqual(ref, Fraction(interval[1]))

    def test_fraction_contains_reference(self):
        with mp.workdps(150):
            for a, x in [(.39, 1.39), (.5, 1.5), (1e-100, 1.), (.9, 32.), (1., 2.)]:
                lo, hi, depth = enclose(a, x)
                aa, xx = mp.mpf(a), mp.mpf(x)
                ref = mp.exp(xx) * xx**(-aa) * mp.gammainc(aa, xx, mp.inf)
                self.assertLessEqual(mp.mpf(str(lo)), ref)
                self.assertLessEqual(ref, mp.mpf(str(hi)))
                self.assertLessEqual((mp.mpf(str(hi))-mp.mpf(str(lo)))/ref, mp.mpf(2)**-105)
                self.assertLessEqual(depth, 4096)

    def test_domain_and_cap_fail_closed(self):
        for a, x in [(0., 1.), (1.01, 2.), (.5, 0.), (.5, math.inf), (math.nan, 1.)]:
            with self.assertRaises(ValueError):
                enclose(a, x)
        with self.assertRaises(RuntimeError):
            enclose(.39, 1.39, cap=1)


if __name__ == '__main__':
    unittest.main()
