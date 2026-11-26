"""#175, second registration: tests of its tools. The rules are tested on verdicts made up here."""
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import prefactor_model as M  # noqa: E402
import prefactor_study as S1  # noqa: E402
import prefactor_study2 as S2  # noqa: E402


def point(kind, verdict="not inferior", library=1e-15, excel=1e-16):
    return {"kind": kind, "verdict": verdict, "library_rel": library, "excel_rel": excel,
            "kstats_call": "K(" + kind + ")"}


def case(n=100, inferior_before=40, inferior_after=10, new=0, new_error=2e-14, ib_after=1e-15,
         floor_new_error=None, floor_after=1, defect_after=1e-15, library_error=False):
    """A holdout of n ib points; the first `inferior_before` were inferior."""
    before = {"holdout": {}, "floor": {}, "defect_points": ["m:L1:A"]}
    holdout = {}
    for i in range(n):
        k = f"ib:P{i:03d}:X"
        was = i < inferior_before
        before["holdout"][k] = {"verdict": "inferior" if was else "not inferior",
                                "library_rel": 2e-13 if was else 1e-15, "excel_rel": 1e-16}
        now = i < inferior_after or (inferior_before <= i < inferior_before + new)
        holdout[k] = point(k, "inferior" if now else "not inferior",
                           new_error if (now and not was) else ib_after)
    if library_error:
        holdout["ib:P000:X"]["verdict"] = "library error"
    before["floor"] = {"m:L1:A": {"verdict": "inferior", "library_rel": 2e-13, "excel_rel": 1e-16},
                       "m:L2:B": {"verdict": "inferior", "library_rel": 5e-14, "excel_rel": 1e-16},
                       "m:L3:C": {"verdict": "not inferior", "library_rel": 1e-15,
                                  "excel_rel": 1e-16}}
    floor = {"m:L1:A": point("m:L1:A", "not inferior", defect_after),
             "m:L2:B": point("m:L2:B", "inferior" if floor_after >= 1 else "not inferior", 5e-14),
             "m:L3:C": point("m:L3:C", "inferior" if floor_new_error else "not inferior",
                             floor_new_error or 1e-15)}
    if floor_after >= 2:
        floor["m:L1:A"]["verdict"] = "inferior"
    return holdout, floor, before


class Design(unittest.TestCase):
    def test_grid_matches_the_generator(self):
        self.assertEqual(S2.build(check=True), 0)

    def test_no_point_is_in_the_first_holdout(self):
        self.assertEqual(S2.overlap(), [])

    def test_every_ib_point_is_in_the_rule_region(self):
        for cls, fn, args in S2.design():
            if cls == "ib":
                _, _, a, b = S1.beta_call(fn, args)
                self.assertTrue(M.rule_applies(a, b), (fn, args))


class Rules(unittest.TestCase):
    def failures(self, **kw):
        return S2.judge(*case(**kw))[0]

    def test_clean_case_passes(self):
        self.assertEqual(self.failures(), [])

    def test_k1_forward(self):
        self.assertTrue(any(f.startswith("K1") for f in self.failures(ib_after=3e-13)))

    def test_k2_a_point_crossing_the_floor_by_rounding_is_allowed(self):
        self.assertEqual(self.failures(new=2, new_error=2e-14), [])

    def test_k2_a_real_loss_fails(self):
        self.assertTrue(any(f.startswith("K2") for f in self.failures(new=1, new_error=5e-14)))

    def test_k2_too_many_newly_inferior_points_fail(self):
        self.assertTrue(any(f.startswith("K2") for f in self.failures(new=3, new_error=2e-14)))

    def test_k2_inferior_points_must_halve(self):
        self.assertTrue(any(f.startswith("K2") for f in self.failures(inferior_after=21)))
        self.assertEqual(self.failures(inferior_after=20), [])

    def test_k3_library_error(self):
        self.assertTrue(any(f.startswith("K3") for f in self.failures(library_error=True)))

    def test_k4_defect_floor_loss_and_count(self):
        self.assertTrue(any(f.startswith("K4") for f in self.failures(defect_after=5e-14)))
        self.assertTrue(any(f.startswith("K4") for f in self.failures(floor_new_error=5e-14)))
        self.assertTrue(any(f.startswith("K4") for f in self.failures(floor_after=2)))


if __name__ == "__main__":
    unittest.main()
