"""#365: tests of the Student t far-tail pair model and of the study's rules, on verdicts
made up here."""
import os
import sys
import unittest

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import tt_model as TT  # noqa: E402
import tt_study as S  # noqa: E402

P = S.P
WORST = (40.0, 1999.0)               # 2.82E-13 in Excel, the worst comparison row of the issue


def ref(t, df):
    with mp.workdps(60):
        return S._survival(mp.mpf(t), mp.mpf(df))


def verdict(kind, state="not inferior", library=1e-15, excel=1e-16):
    return {"kind": kind, "verdict": state, "library_rel": library, "excel_rel": excel,
            "kstats_call": "K(" + kind + ")"}


def case(sv_error=1e-15, sv_state="not inferior", in_error=1e-15, in_state="not inferior",
         in_before="not inferior", ed_error=2e-14, defect=1e-15, new_error=None, more=False,
         library_error=False):
    before = {"holdout": {}, "defect": {"17;99999": 1.5e-13, "40;1999": 2.8e-13},
              "elsewhere": {}}
    holdout = {}
    for i, cls in enumerate(("sv", "cu")):
        k = f"{cls}:P{i:03d}:StudentT"
        before["holdout"][k] = {"verdict": "inferior", "library_rel": 3e-13, "excel_rel": 1e-14}
        holdout[k] = verdict(k, sv_state if i == 0 else "not inferior",
                             sv_error if i == 0 else 1e-15)
    k = "in:P005:StudentT_InverseCumulative"
    before["holdout"][k] = {"verdict": in_before, "library_rel": 1e-15, "excel_rel": 1e-15}
    holdout[k] = verdict(k, in_state, in_error)
    k = "ed:P006:StudentT_Survival"
    before["holdout"][k] = {"verdict": "not inferior", "library_rel": 2e-14, "excel_rel": 3e-14}
    holdout[k] = verdict(k, library=ed_error)
    if library_error:
        holdout["sv:P000:StudentT"] = {"kind": "sv:P000:StudentT", "verdict": "library error",
                                       "kstats_call": "K"}
    elsewhere = {}
    for name, _ in S.ELSEWHERE:
        prior = {"x:L1:A": {"verdict": "inferior", "library_rel": 2e-13, "excel_rel": 1e-15},
                 "x:L2:B": {"verdict": "not inferior", "library_rel": 1e-15, "excel_rel": 1e-15}}
        after = {"x:L1:A": verdict("x:L1:A", "inferior", 2e-13), "x:L2:B": verdict("x:L2:B")}
        if name == "holdout of #498" and new_error is not None:
            after["x:L2:B"] = verdict("x:L2:B", "inferior", new_error)
            if not more:
                after["x:L1:A"] = verdict("x:L1:A")
        before["elsewhere"][name] = prior
        elsewhere[name] = after
    return holdout, {"17;99999": 1e-15, "40;1999": defect}, elsewhere, before


class Model(unittest.TestCase):
    def test_the_worst_point_of_the_issue(self):
        r = ref(*WORST)
        today = P.rel(TT.tail(*WORST, rule=False), r)
        self.assertTrue(1e-13 < today < 6e-13, today)
        self.assertLess(P.rel(TT.tail(*WORST), r), 3e-15)

    def test_the_rows_of_the_issue_under_the_rule(self):
        for key in S.defect_points():
            t, df = (float(v) for v in key.split(";"))
            self.assertLess(P.rel(TT.tail(t, df), ref(t, df)), 3e-15, key)

    def test_where_the_rule_is_not_tried(self):
        for t, df in ((16.0, 5000.0), (8.0, 300.0), (17.5, 1e5), (40.0, 1e6)):
            self.assertFalse(TT.served(t, df))
        self.assertEqual(TT.tail(16.0, 5000.0), TT.tail(16.0, 5000.0, rule=False))
        self.assertTrue(TT.served(16.03125, 99998.0))

    def test_where_the_pair_route_declines(self):
        self.assertIsNone(TT.pair_tail(1e120, 5.0))
        self.assertIsNone(TT.pair_tail(1e60, 1e-30))

    def test_small_df_and_a_large_argument(self):
        for t, df in ((2500.0, 5.0), (1e30, 5.0), (1e50, 5.0), (400.0, 23.0)):
            self.assertLess(P.rel(TT.tail(t, df), ref(t, df)), 3e-15, (t, df))


class Study(unittest.TestCase):
    def rules(self, **kw):
        return sorted({f.split()[0] for f in S.judge(*case(**kw))[0]})

    def test_design_is_fresh_and_the_grid_matches_it(self):
        self.assertEqual(S.overlap(), [])
        self.assertEqual(S.main(["--build", "--check"]), 0)
        self.assertEqual({c for c, _, _ in S.design()}, {"sv", "cu", "in", "ed"})
        for cls, _, args in S.design():
            if cls in ("sv", "cu"):
                self.assertTrue(TT.served(abs(args[0]), args[1]), args)
            if cls == "ed":
                self.assertFalse(TT.served(*args), args)
            self.assertEqual(args[1], int(args[1]), args)     # Excel truncates df

    def test_no_point_is_a_comparison_row_of_the_issue(self):
        rows = set(S.defect_points())
        self.assertEqual(len(rows), 32)
        for cls, _, args in S.design():
            self.assertNotIn(f"{abs(args[0]):g};{args[1]:g}", rows)

    def test_accepted_when_every_rule_holds(self):
        self.assertEqual(self.rules(), [])

    def test_t1_and_t2_judge_the_tail_points(self):
        self.assertEqual(self.rules(sv_error=2e-14), ["T1"])
        self.assertEqual(self.rules(sv_error=2e-14, sv_state="inferior"), ["T1", "T2"])
        self.assertEqual(self.rules(sv_error=5e-15, sv_state="inferior"), ["T2"])
        self.assertEqual(self.rules(library_error=True), ["T2"])

    def test_t3_a_row_of_the_issue(self):
        self.assertEqual(self.rules(defect=2e-14), ["T3"])
        self.assertEqual(self.rules(defect=None), ["T3"])

    def test_t4_the_inverse(self):
        self.assertEqual(self.rules(in_error=5e-13), [])
        self.assertEqual(self.rules(in_error=2e-12), ["T4"])
        self.assertEqual(self.rules(in_state="inferior"), ["T4"])
        self.assertEqual(self.rules(in_state="inferior", in_before="inferior"), [])

    def test_t5_an_edge_does_not_change(self):
        self.assertEqual(self.rules(ed_error=1e-15), ["T5"])
        self.assertEqual(self.rules(ed_error=3e-14), ["T5"])

    def test_t6_a_loss_elsewhere_or_more_inferior_points(self):
        self.assertEqual(self.rules(new_error=2e-14), [])
        self.assertEqual(self.rules(new_error=5e-14), ["T6"])
        self.assertEqual(self.rules(new_error=2e-14, more=True), ["T6"])


class NotComparedWithExcel(unittest.TestCase):
    """A point Excel did not answer is still judged on the library alone."""

    def failures(self, **point):
        holdout, defect, elsewhere, before = case()
        k = "sv:P000:StudentT"
        holdout[k] = dict(kind=k, verdict="not compared", kstats_call="K", **point)
        return sorted({f.split()[0] for f in S.judge(holdout, defect, elsewhere, before)[0]})

    def test_an_inaccurate_library_value_fails_the_accuracy_rule(self):
        self.assertEqual(self.failures(library_rel=5e-13), ["T1"])
        self.assertEqual(self.failures(library_rel=1e-15), [])

    def test_a_library_error_fails_whatever_excel_returned(self):
        self.assertEqual(self.failures(library_missing=True), ["T2"])


if __name__ == "__main__":
    unittest.main()
