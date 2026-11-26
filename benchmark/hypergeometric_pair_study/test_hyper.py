"""#496: tests of the Hypergeometric pair model and of the study's rules, on verdicts made up here."""
import os
import sys
import unittest

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import hyper_model as H  # noqa: E402
import hyper_study as S  # noqa: E402

P = H.P
ISSUE_POINT = (100.0, 200.0, 2000.0, 50000.0)


def ref(args):
    with mp.workdps(60):
        return S.S1.REFERENCE["Hypergeometric_PMF"](*[mp.mpf(a) for a in args])


def verdict(kind, state="not inferior", library=1e-15, excel=1e-16, call=None):
    return {"kind": kind, "verdict": state, "library_rel": library, "excel_rel": excel,
            "kstats_call": call or "K(" + kind + ")"}


def case(hg_error=1e-15, hg_state="not inferior", nb_error=1e-15, nb_state="not inferior",
         defect=1e-15, new_error=None, more=False, library_error=False):
    before = {"holdout": {}, "elsewhere": {}}
    holdout = {}
    for i in range(3):
        k = f"hg:P{i:03d}:Hypergeometric_PMF"
        before["holdout"][k] = {"verdict": "inferior", "library_rel": 3e-13, "excel_rel": 1e-14}
        holdout[k] = verdict(k, hg_state if i == 0 else "not inferior",
                             hg_error if i == 0 else 1e-15)
    k = "nb:P009:NegativeBinomial_PMF"
    before["holdout"][k] = {"verdict": "inferior", "library_rel": 3e-13, "excel_rel": 1e-14}
    holdout[k] = verdict(k, nb_state, nb_error)
    if library_error:
        holdout["hg:P001:Hypergeometric_PMF"] = {"kind": "hg:P001:Hypergeometric_PMF",
                                                "verdict": "library error", "kstats_call": "K"}
    elsewhere = {}
    for name, _ in S.ELSEWHERE:
        prior = {"x:L1:A": {"verdict": "inferior", "library_rel": 2e-13, "excel_rel": 1e-15},
                 "x:L2:B": {"verdict": "not inferior", "library_rel": 1e-15, "excel_rel": 1e-15}}
        after = {"x:L1:A": verdict("x:L1:A", "inferior", 2e-13), "x:L2:B": verdict("x:L2:B")}
        if name == "floor grid":
            prior["m:L9:D"] = {"verdict": "inferior", "library_rel": 3.5e-14, "excel_rel": 1.9e-14}
            after["m:L9:D"] = verdict("m:L9:D", library=defect, call=S.DEFECT_CALLS[0])
        if name == "holdout of #498" and new_error is not None:
            after["x:L2:B"] = verdict("x:L2:B", "inferior", new_error)
            if not more:
                after["x:L1:A"] = verdict("x:L1:A")
        before["elsewhere"][name] = prior
        elsewhere[name] = after
    return holdout, elsewhere, before


class Model(unittest.TestCase):
    def test_the_point_of_the_issue(self):
        r = ref(ISSUE_POINT)
        # 3.49E-14 measured in Excel (run 45848441378); the port returns it.
        today = P.rel(H.hypergeometric_pmf(*ISSUE_POINT, rule=False), r)
        self.assertTrue(3.3e-14 < today < 3.7e-14, today)
        self.assertLess(P.rel(H.hypergeometric_pmf(*ISSUE_POINT), r), 5e-15)

    def test_a_ratio_that_is_not_a_double(self):
        args = (15609.0, 70001.0, 250000.0, 999983.0)
        self.assertLess(P.rel(H.hypergeometric_pmf(*args), ref(args)), 5e-15)

    def test_edges_are_not_modelled(self):
        for args in ((0.0, 10.0, 5.0, 50.0), (5.0, 10.0, 5.0, 50.0), (10.0, 10.0, 20.0, 50.0)):
            self.assertFalse(H.interior(*args))
            with self.assertRaises(ValueError):
                H.hypergeometric_pmf(*args)


class Study(unittest.TestCase):
    def rules(self, **kw):
        return sorted({f.split()[0] for f in S.judge(*case(**kw))[0]})

    def test_design_is_fresh_and_the_grid_matches_it(self):
        self.assertEqual(S.overlap(), [])
        self.assertEqual(S.main(["--build", "--check"]), 0)
        self.assertEqual({c for c, _, _ in S.design()}, {"hg", "nb"})
        for cls, fn, args in S.design():
            if cls == "hg":
                self.assertTrue(H.interior(*args), args)

    def test_the_point_of_the_issue_is_a_row_of_the_floor_grid(self):
        calls = {r["kstats_call"] for r in S.S1.read(S.S1.FB.GRID)}
        self.assertIn(S.DEFECT_CALLS[0], calls)

    def test_accepted_when_every_rule_holds(self):
        self.assertEqual(self.rules(), [])

    def test_m1_and_m2_judge_the_hypergeometric_points(self):
        self.assertEqual(self.rules(hg_error=2e-14), ["M1"])
        self.assertEqual(self.rules(hg_error=2e-14, hg_state="inferior"), ["M1", "M2"])
        self.assertEqual(self.rules(library_error=True), ["M2"])

    def test_the_negative_binomial_points_do_not_gate(self):
        self.assertEqual(self.rules(nb_error=5e-13, nb_state="inferior"), [])

    def test_m3_the_point_of_the_issue(self):
        self.assertEqual(self.rules(defect=2e-14), ["M3"])

    def test_m4_a_loss_elsewhere_or_more_inferior_points(self):
        self.assertEqual(self.rules(new_error=2e-14), [])
        self.assertEqual(self.rules(new_error=5e-14), ["M4"])
        self.assertEqual(self.rules(new_error=2e-14, more=True), ["M4"])


class NotComparedWithExcel(unittest.TestCase):
    """A point Excel did not answer is still judged on the library alone."""

    def failures(self, **point):
        holdout, elsewhere, before = case()
        holdout["hg:P000:Hypergeometric_PMF"] = dict(kind="hg:P000:Hypergeometric_PMF", verdict="not compared",
                                 kstats_call="K", **point)
        return sorted({f.split()[0] for f in S.judge(holdout, elsewhere, before)[0]})

    def test_an_inaccurate_library_value_fails_the_accuracy_rule(self):
        self.assertEqual(self.failures(library_rel=5e-13), ["M1"])
        self.assertEqual(self.failures(library_rel=1e-15), [])

    def test_a_library_error_fails_whatever_excel_returned(self):
        self.assertEqual(self.failures(library_missing=True), ["M2"])


if __name__ == "__main__":
    unittest.main()
