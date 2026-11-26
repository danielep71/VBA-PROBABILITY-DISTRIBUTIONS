"""#504: tests of the Negative Binomial and Geometric pair model and of the study's rules,
on verdicts made up here."""
import os
import sys
import unittest

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import nb_model as NB  # noqa: E402
import nb_study as S  # noqa: E402

P = NB.P
WORST = (21622.0, 150000.0, 0.9)


def ref(fn, args):
    with mp.workdps(60):
        return S.S1.REFERENCE[fn](*[mp.mpf(a) for a in args])


def verdict(kind, state="not inferior", library=1e-15, excel=1e-16, call=None):
    return {"kind": kind, "verdict": state, "library_rel": library, "excel_rel": excel,
            "kstats_call": call or "K(" + kind + ")"}


def case(nb_error=1e-15, nb_state="not inferior", ge_before="not inferior", ge_error=4e-14,
         ge_state="not inferior", defect=1e-15, new_error=None, more=False,
         library_error=False):
    before = {"holdout": {}, "elsewhere": {}}
    holdout = {}
    for i in range(3):
        k = f"nb:P{i:03d}:NegativeBinomial_PMF"
        before["holdout"][k] = {"verdict": "inferior", "library_rel": 3e-13, "excel_rel": 1e-14}
        holdout[k] = verdict(k, nb_state if i == 0 else "not inferior",
                             nb_error if i == 0 else 1e-15)
    k = "ge:P009:Geometric_PMF"
    before["holdout"][k] = {"verdict": ge_before, "library_rel": 4e-14, "excel_rel": 1e-15}
    holdout[k] = verdict(k, ge_state, ge_error)
    if library_error:
        holdout["nb:P001:NegativeBinomial_PMF"] = {"kind": "nb:P001:NegativeBinomial_PMF",
                                                  "verdict": "library error", "kstats_call": "K"}
    elsewhere = {}
    for name, _ in S.ELSEWHERE:
        prior = {"x:L1:A": {"verdict": "inferior", "library_rel": 2e-13, "excel_rel": 1e-15},
                 "x:L2:B": {"verdict": "not inferior", "library_rel": 1e-15, "excel_rel": 1e-15}}
        after = {"x:L1:A": verdict("x:L1:A", "inferior", 2e-13), "x:L2:B": verdict("x:L2:B")}
        if name == "holdout of #496":
            for i, call in enumerate(S.DEFECT_CALLS):
                k = f"nb:D{i}:NegativeBinomial_PMF"
                prior[k] = {"verdict": "inferior", "library_rel": 3e-13, "excel_rel": 1e-14}
                after[k] = verdict(k, library=defect if i == 0 else 1e-15, call=call)
        if name == "holdout of #498" and new_error is not None:
            after["x:L2:B"] = verdict("x:L2:B", "inferior", new_error)
            if not more:
                after["x:L1:A"] = verdict("x:L1:A")
        before["elsewhere"][name] = prior
        elsewhere[name] = after
    return holdout, elsewhere, before


class Model(unittest.TestCase):
    def test_the_worst_point_of_the_issue(self):
        r = ref("NegativeBinomial_PMF", WORST)
        # 6.11E-13 measured in Excel (the export applied for #496); the port is near it.
        today = P.rel(NB.negative_binomial_pmf(*WORST, rule=False), r)
        self.assertTrue(4e-13 < today < 9e-13, today)
        self.assertLess(P.rel(NB.negative_binomial_pmf(*WORST), r), 5e-15)

    def test_the_eight_points_of_the_issue_under_the_rule(self):
        for call in S.DEFECT_CALLS:
            args = tuple(float(a) for a in call.split("(")[1][:-1].split(";"))
            self.assertLess(P.rel(NB.negative_binomial_pmf(*args),
                                  ref("NegativeBinomial_PMF", args)), 5e-15, call)

    def test_the_geometric_mass(self):
        for args in ((9990.0, 0.002), (50000.0, 1e-05), (7.0, 0.93), (1.0, 0.5)):
            self.assertLess(P.rel(NB.geometric_pmf(*args), ref("Geometric_PMF", args)),
                            5e-15, args)


class Study(unittest.TestCase):
    def rules(self, **kw):
        return sorted({f.split()[0] for f in S.judge(*case(**kw))[0]})

    def test_design_is_fresh_and_the_grid_matches_it(self):
        self.assertEqual(S.overlap(), [])
        self.assertEqual(S.main(["--build", "--check"]), 0)
        self.assertEqual({c for c, _, _ in S.design()}, {"nb", "ge"})
        for _, _, args in S.design():
            self.assertGreaterEqual(args[0], 1.0, args)

    def test_the_points_of_the_issue_are_rows_of_the_holdout_of_496(self):
        calls = {r["kstats_call"] for r in S.S1.read(S.S6.GRID)}
        self.assertEqual([c for c in S.DEFECT_CALLS if c not in calls], [])

    def test_accepted_when_every_rule_holds(self):
        self.assertEqual(self.rules(), [])

    def test_n1_and_n2_judge_the_negative_binomial_points(self):
        self.assertEqual(self.rules(nb_error=2e-14), ["N1"])
        self.assertEqual(self.rules(nb_error=2e-14, nb_state="inferior"), ["N1", "N2"])
        self.assertEqual(self.rules(library_error=True), ["N2"])

    def test_n3_a_point_of_the_issue(self):
        self.assertEqual(self.rules(defect=2e-14), ["N3"])

    def test_n4_a_loss_elsewhere_or_more_inferior_points(self):
        self.assertEqual(self.rules(new_error=2e-14), [])
        self.assertEqual(self.rules(new_error=5e-14), ["N4"])
        self.assertEqual(self.rules(new_error=2e-14, more=True), ["N4"])

    def test_the_geometric_rule_applies_only_where_excel_was_closer_before(self):
        self.assertFalse(S.geometric_rule(case()[2]))
        self.assertTrue(S.geometric_rule(case(ge_before="inferior")[2]))

    def test_n5_without_the_geometric_rule_no_observation_changes(self):
        self.assertEqual(self.rules(), [])                       # 4E-14, as before
        self.assertEqual(self.rules(ge_error=1e-15), ["N5"])
        self.assertEqual(self.rules(ge_error=5e-14), ["N5"])

    def test_with_the_geometric_rule_its_points_are_judged(self):
        self.assertEqual(self.rules(ge_before="inferior", ge_error=1e-15), [])
        self.assertEqual(self.rules(ge_before="inferior"), ["N1"])
        self.assertEqual(self.rules(ge_before="inferior", ge_error=2e-14, ge_state="inferior"),
                         ["N1", "N2"])


class NotComparedWithExcel(unittest.TestCase):
    """A point Excel did not answer is still judged on the library alone."""

    def failures(self, **point):
        holdout, elsewhere, before = case()
        k = "nb:P000:NegativeBinomial_PMF"
        holdout[k] = dict(kind=k, verdict="not compared", kstats_call="K", **point)
        return sorted({f.split()[0] for f in S.judge(holdout, elsewhere, before)[0]})

    def test_an_inaccurate_library_value_fails_the_accuracy_rule(self):
        self.assertEqual(self.failures(library_rel=5e-13), ["N1"])
        self.assertEqual(self.failures(library_rel=1e-15), [])

    def test_a_library_error_fails_whatever_excel_returned(self):
        self.assertEqual(self.failures(library_missing=True), ["N2"])


if __name__ == "__main__":
    unittest.main()
