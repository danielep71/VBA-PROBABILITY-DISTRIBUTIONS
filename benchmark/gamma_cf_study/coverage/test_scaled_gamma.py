"""Revealed integration screen, not a disjoint acceptance set."""
import json
from pathlib import Path
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

    def test_revealed_b6_points(self):
        rows = json.loads(Path(__file__).with_name('references.json').read_text())
        tested = 0
        for row in rows:
            a, x = float.fromhex(row['a_hex']), float.fromhex(row['x_hex'])
            if not 0 < a <= 1 or not 0 < x <= 32:
                continue
            got = public(a, x)
            for value, key in zip(got, ['p', 'q']):
                ref = float(row[key])
                self.assertLessEqual(abs(value-ref), 2*math.ulp(ref), (a,x,key,value,ref))
            tested += 1
        self.assertGreater(tested, 0)

    def test_routing_neighbours_and_oracle_precision_stability(self):
        for a in [2.**-1074, 1e-100, 1e-6, .0001, .01, .25, .5, 1.]:
            center = a+1
            points = [math.nextafter(center, 0.), center, math.nextafter(center, math.inf)]
            results = []
            for x in points:
                got = public(a,x)
                refs = []
                for precision in [160, 320]:
                    with mp.workdps(precision):
                        aa, xx = mp.mpf(a), mp.mpf(x)
                        refs.append((float(mp.gammainc(aa,0,xx,regularized=True)),
                                     float(mp.gammainc(aa,xx,mp.inf,regularized=True))))
                self.assertEqual(refs[0], refs[1], (a,x))
                for value, ref in zip(got, refs[1]):
                    self.assertLessEqual(abs(value-ref), 2*math.ulp(ref), (a,x,got,refs))
                results.append(got)
            for left,right in zip(results, results[1:]):
                self.assertLessEqual(left[0], right[0], (a,results))
                self.assertGreaterEqual(left[1], right[1], (a,results))

    def test_domain_is_explicit(self):
        for a,x in [(0,1),(-1,1),(2,1),(1,0),(1,33),(math.inf,1),(1,math.nan)]:
            with self.assertRaises(ValueError):
                public(a,x)


if __name__ == '__main__':
    unittest.main()
