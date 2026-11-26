"""#498: tests of the pair model of the deviance."""
import math
import os
import random
import sys
import unittest

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import pair_model as P  # noqa: E402

POINTS = (("poisson", P.poisson_pmf, (5600.0, 4000.0)),
          ("gamma", P.gamma_density, (7500.0, 5000.0)),
          ("gamma", P.gamma_density, (3500.0, 5000.0)),
          ("binomial", P.binomial_pmf, (4076.0, 8000.0, 0.4)),
          ("binomial", P.binomial_pmf, (4466.0, 300000.0, 0.02)))


class Pairs(unittest.TestCase):
    def test_error_free_transforms(self):
        rnd = random.Random(1)
        for _ in range(200):
            a, b = rnd.uniform(-1e6, 1e6), rnd.uniform(-1e-3, 1e3)
            s, e = P.two_sum(a, b)
            self.assertEqual(mp.mpf(s) + mp.mpf(e), mp.mpf(a) + mp.mpf(b))
            p, e = P.two_prod(a, b)
            with mp.workdps(60):
                self.assertEqual(mp.mpf(p) + mp.mpf(e), mp.mpf(a) * mp.mpf(b))

    def test_log_two_is_a_pair_of_log_two(self):
        with mp.workdps(40):
            self.assertLess(abs(mp.mpf(P.LOG_TWO[0]) + mp.mpf(P.LOG_TWO[1]) - mp.log(2)), 1e-32)

    def test_deviance_pair_is_far_below_one_unit_of_its_high_part(self):
        rnd = random.Random(7)
        worst = 0.0
        for _ in range(400):
            x = 10 ** rnd.uniform(0, 6)
            v = rnd.choice((10 ** rnd.uniform(-6, -1), rnd.uniform(0.1, 0.99))) * rnd.choice((-1, 1))
            mean = x * (1 - v) / (1 + v)
            ref = P.reference("deviance", x, mean)
            if not 1e-12 < ref < 700:
                continue
            h, l = P.deviance(x, mean)
            with mp.workdps(60):
                worst = max(worst, float(abs((mp.mpf(h) + mp.mpf(l) - ref) / ref)))
        self.assertLess(worst, 1e-24)

    def test_a_mean_given_as_a_pair_is_used_whole(self):
        n, p = 8000.0, 0.4
        pair = P.two_prod(n, p)
        self.assertNotEqual(pair[1], 0.0)
        with mp.workdps(60):
            ref = P.reference("deviance", 4076.0, mp.mpf(n) * mp.mpf(p))
            h, l = P.deviance(4076.0, pair)
            self.assertLess(abs(mp.mpf(h) + mp.mpf(l) - ref), 1e-25)

    def test_the_points_of_the_issue(self):
        for fn, f, args in POINTS:
            ref = P.reference(fn, *args)
            self.assertGreater(P.rel(f(*args, rule=False), ref), 1e-13, (fn, args))
            self.assertLess(P.rel(f(*args), ref), 5e-15, (fn, args))

    def test_today_is_the_port_of_the_source(self):
        # Poisson_PMF(5600, 4000) measured 4.4E-13 in Excel (run 45789721558).
        e = P.rel(P.poisson_pmf(5600.0, 4000.0, rule=False), P.reference("poisson", 5600.0, 4000.0))
        self.assertTrue(3e-13 < e < 6e-13, e)

    def test_underflow_returns_zero(self):
        self.assertEqual(P.poisson_pmf(5000.0, 100.0), 0.0)
        self.assertTrue(math.isfinite(P.gamma_density(3.0, 2.0)))


if __name__ == "__main__":
    unittest.main()
