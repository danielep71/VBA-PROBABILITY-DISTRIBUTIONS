"""#498: tests of the pair model of the deviance and of the study's rules, on verdicts made up here."""
import math
import os
import random
import sys
import unittest

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import pair_model as P  # noqa: E402
import pair_study as S  # noqa: E402

POINTS = (("poisson", P.poisson_pmf, (5600.0, 4000.0)),
          ("gamma", P.gamma_density, (7500.0, 5000.0, 1.0)),
          ("gamma", P.gamma_density, (3500.0, 5000.0, 1.0)),
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

    def test_at_the_mean_the_deviance_is_zero(self):
        self.assertEqual(P.deviance(7.0, 7.0), (0.0, 0.0))
        ref = P.reference("poisson", 7.0, 7.0)
        self.assertLess(P.rel(P.poisson_pmf(7.0, 7.0), ref), 5e-15)

    def test_underflow_returns_zero(self):
        self.assertEqual(P.poisson_pmf(5000.0, 100.0), 0.0)
        self.assertTrue(math.isfinite(P.gamma_density(3.0, 2.0)))


def verdict(kind, state="not inferior", library=1e-15, excel=1e-16, call=None):
    return {"kind": kind, "verdict": state, "library_rel": library, "excel_rel": excel,
            "kstats_call": call or "K(" + kind + ")"}


def case(holdout_error=1e-15, holdout_state="not inferior", defect=1e-15, new_error=None,
         more=False, library_error=False):
    """Verdicts made up here: a holdout of three points, and the three other grids."""
    before = {"holdout": {}, "elsewhere": {}}
    holdout = {}
    for i in range(3):
        k = f"pm:P{i:03d}:Poisson_PMF"
        before["holdout"][k] = {"verdict": "inferior", "library_rel": 3e-13, "excel_rel": 1e-14}
        holdout[k] = verdict(k, holdout_state if i == 0 else "not inferior",
                             holdout_error if i == 0 else 1e-15)
    if library_error:
        holdout["pm:P001:Poisson_PMF"] = {"kind": "pm:P001:Poisson_PMF", "verdict": "library error",
                                          "kstats_call": "K"}
    elsewhere = {}
    for name, _ in S.ELSEWHERE:
        prior = {"x:L1:A": {"verdict": "inferior", "library_rel": 2e-13, "excel_rel": 1e-15},
                 "x:L2:B": {"verdict": "not inferior", "library_rel": 1e-15, "excel_rel": 1e-15}}
        after = {"x:L1:A": verdict("x:L1:A", "inferior", 2e-13),
                 "x:L2:B": verdict("x:L2:B")}
        if name == "second holdout of #175":
            for n, call in enumerate(S.DEFECT_CALLS):
                k = f"dv:P{n:03d}:D"
                prior[k] = {"verdict": "inferior", "library_rel": 3e-13, "excel_rel": 5e-14}
                after[k] = verdict(k, library=defect, call=call)
        if name == "floor grid" and new_error is not None:
            after["x:L2:B"] = verdict("x:L2:B", "inferior", new_error)
            if not more:
                after["x:L1:A"] = verdict("x:L1:A")
        before["elsewhere"][name] = prior
        elsewhere[name] = after
    return holdout, elsewhere, before


class Study(unittest.TestCase):
    def rules(self, **kw):
        return sorted({f.split()[0] for f in S.judge(*case(**kw))[0]})

    def test_design_is_fresh_and_the_grid_matches_it(self):
        self.assertEqual(S.overlap(), [])
        self.assertEqual(S.main(["--build", "--check"]), 0)
        classes = {c for c, _, _ in S.design()}
        self.assertEqual(classes, {"pm", "bn", "gd", "cs"})

    def test_the_holdout_reaches_large_deviances_on_both_sides(self):
        below = [a for c, fn, a in S.design() if fn == "Poisson_PMF" and a[0] < a[1]]
        above = [a for c, fn, a in S.design() if fn == "Poisson_PMF" and a[0] > a[1]]
        for side in (below, above):
            drops = [float(-P.mp.log(P.reference("poisson", *a))) for a in side]
            self.assertGreater(max(drops), 500)

    def test_accepted_when_every_rule_holds(self):
        self.assertEqual(self.rules(), [])

    def test_j1_a_holdout_point_above_the_floor(self):
        self.assertEqual(self.rules(holdout_error=2e-14), ["J1"])

    def test_j2_an_inferior_holdout_point_or_a_library_error(self):
        self.assertEqual(self.rules(holdout_error=2e-14, holdout_state="inferior"), ["J1", "J2"])
        self.assertEqual(self.rules(library_error=True), ["J2"])

    def test_j3_a_point_of_the_issue_above_the_floor(self):
        self.assertEqual(self.rules(defect=2e-14), ["J3"])

    def test_j4_a_loss_elsewhere_or_more_inferior_points(self):
        self.assertEqual(self.rules(new_error=2e-14), [])
        self.assertEqual(self.rules(new_error=5e-14), ["J4"])
        self.assertEqual(self.rules(new_error=2e-14, more=True), ["J4"])

    def test_the_five_points_are_rows_of_the_second_holdout_of_175(self):
        calls = {r["kstats_call"] for r in S.S1.read(S.S2.GRID)}
        for call in S.DEFECT_CALLS:
            self.assertIn(call, calls)


if __name__ == "__main__":
    unittest.main()
