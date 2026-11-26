"""#501: tests of the density pair model and of the study's rules, on verdicts made up here."""
import os
import sys
import unittest

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import density_model as D  # noqa: E402
import density_study as S  # noqa: E402

P = D.P
ISSUE_POINTS = (("F_Density", D.f_density, (6.0, 300.0, 900.0)),
                ("F_Density", D.f_density, (5.0, 500.0, 700.0)),
                ("StudentT_Density", D.student_t_density, (20.0, 900.0)),
                ("StudentT_Density", D.student_t_density, (16.0, 700.0)),
                ("Beta_Density", D.beta_density, (0.812111, 150.0, 90.0)),
                ("Beta_Density", D.beta_density, (0.362112, 480.0, 500.0)))


def ref(fn, args):
    with mp.workdps(60):
        return S.S1.REFERENCE[fn](*[mp.mpf(a) for a in args])


def verdict(kind, state="not inferior", library=1e-15, excel=1e-16, call=None):
    return {"kind": kind, "verdict": state, "library_rel": library, "excel_rel": excel,
            "kstats_call": call or "K(" + kind + ")"}


def case(holdout_error=1e-15, holdout_state="not inferior", defect=1e-15, new_error=None,
         more=False, library_error=False):
    before = {"holdout": {}, "elsewhere": {}}
    holdout = {}
    for i in range(3):
        k = f"fd:P{i:03d}:F_Density"
        before["holdout"][k] = {"verdict": "inferior", "library_rel": 3e-13, "excel_rel": 1e-14}
        holdout[k] = verdict(k, holdout_state if i == 0 else "not inferior",
                             holdout_error if i == 0 else 1e-15)
    if library_error:
        holdout["fd:P001:F_Density"] = {"kind": "fd:P001:F_Density", "verdict": "library error",
                                        "kstats_call": "K"}
    elsewhere = {}
    for name, _ in S.ELSEWHERE:
        prior = {"x:L1:A": {"verdict": "inferior", "library_rel": 2e-13, "excel_rel": 1e-15},
                 "x:L2:B": {"verdict": "not inferior", "library_rel": 1e-15, "excel_rel": 1e-15}}
        after = {"x:L1:A": verdict("x:L1:A", "inferior", 2e-13), "x:L2:B": verdict("x:L2:B")}
        if name == "first holdout of #175":
            for n, call in enumerate(S.DEFECT_CALLS):
                k = f"dv:P{n:03d}:D"
                prior[k] = {"verdict": "inferior", "library_rel": 3e-14, "excel_rel": 5e-15}
                after[k] = verdict(k, library=defect, call=call)
        if name == "holdout of #498" and new_error is not None:
            after["x:L2:B"] = verdict("x:L2:B", "inferior", new_error)
            if not more:
                after["x:L1:A"] = verdict("x:L1:A")
        before["elsewhere"][name] = prior
        elsewhere[name] = after
    return holdout, elsewhere, before


class Model(unittest.TestCase):
    def test_the_points_of_the_issue(self):
        for fn, f, args in ISSUE_POINTS:
            self.assertLess(P.rel(f(*args), ref(fn, args)), 5e-15, (fn, args))

    def test_today_is_the_port_of_the_source_at_the_beta_point(self):
        # Beta_Density(0.812111, 150, 90) measured 1.26E-14 in Excel (run 45822509492).
        args = (0.812111, 150.0, 90.0)
        e = P.rel(D.beta_density(*args, rule=False), ref("Beta_Density", args))
        self.assertTrue(1.2e-14 < e < 1.35e-14, e)

    def test_student_t_at_zero_and_at_small_df(self):
        for args in ((0.0, 5.0), (1.5, 1.0), (3.0, 2.0), (40.0, 3.0)):
            self.assertLess(P.rel(D.student_t_density(*args), ref("StudentT_Density", args)),
                            5e-15, args)

    def test_f_density_on_both_sides_of_ratio_one(self):
        for args in ((0.05, 250.0, 1100.0), (9.0, 8.0, 30.0)):
            self.assertLess(P.rel(D.f_density(*args), ref("F_Density", args)), 5e-15, args)


class Study(unittest.TestCase):
    def rules(self, **kw):
        return sorted({f.split()[0] for f in S.judge(*case(**kw))[0]})

    def test_design_is_fresh_and_the_grid_matches_it(self):
        self.assertEqual(S.overlap(), [])
        self.assertEqual(S.main(["--build", "--check"]), 0)
        self.assertEqual({c for c, _, _ in S.design()}, {"bd", "fd", "td"})

    def test_the_ten_points_are_rows_of_the_holdouts_of_175(self):
        calls = {r["kstats_call"] for g in (S.S1.GRID, S.S2.GRID) for r in S.S1.read(g)}
        for call in S.DEFECT_CALLS:
            self.assertIn(call, calls)

    def test_accepted_when_every_rule_holds(self):
        self.assertEqual(self.rules(), [])

    def test_l1_and_l2(self):
        self.assertEqual(self.rules(holdout_error=2e-14), ["L1"])
        self.assertEqual(self.rules(holdout_error=2e-14, holdout_state="inferior"), ["L1", "L2"])
        self.assertEqual(self.rules(library_error=True), ["L2"])

    def test_l3_a_point_of_the_issue_above_the_floor(self):
        self.assertEqual(self.rules(defect=2e-14), ["L3"])

    def test_l4_a_loss_elsewhere_or_more_inferior_points(self):
        self.assertEqual(self.rules(new_error=2e-14), [])
        self.assertEqual(self.rules(new_error=5e-14), ["L4"])
        self.assertEqual(self.rules(new_error=2e-14, more=True), ["L4"])


class NotComparedWithExcel(unittest.TestCase):
    """A point Excel did not answer is still judged on the library alone."""

    def failures(self, **point):
        holdout, elsewhere, before = case()
        holdout["fd:P000:F_Density"] = dict(kind="fd:P000:F_Density", verdict="not compared",
                                 kstats_call="K", **point)
        return sorted({f.split()[0] for f in S.judge(holdout, elsewhere, before)[0]})

    def test_an_inaccurate_library_value_fails_the_accuracy_rule(self):
        self.assertEqual(self.failures(library_rel=5e-13), ["L1"])
        self.assertEqual(self.failures(library_rel=1e-15), [])

    def test_a_library_error_fails_whatever_excel_returned(self):
        self.assertEqual(self.failures(library_missing=True), ["L2"])


if __name__ == "__main__":
    unittest.main()
