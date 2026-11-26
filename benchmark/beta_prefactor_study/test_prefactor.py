"""#175: tests of the prefactor study's tools. The rules are tested on verdicts made up here."""
import os
import sys
import unittest

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import prefactor_model as M  # noqa: E402
import prefactor_study as S  # noqa: E402


def point(kind, verdict="not inferior", library=1e-15, excel=1e-15):
    return {"kind": kind, "verdict": verdict, "library_rel": library, "excel_rel": excel,
            "kstats_call": "K(" + kind + ")"}


def case(ib_after=1e-15, dv_verdict="not inferior", floor_new=False, defect_after=1e-15,
         library_error=False):
    before = {"holdout": {"ib:P001:X": {"verdict": "not inferior", "library_rel": 3e-13,
                                        "excel_rel": 1e-15},
                          "dv:P002:Y": {"verdict": "not inferior", "library_rel": 1e-12,
                                        "excel_rel": 1e-12}},
              "floor": {"m:L1:A": {"verdict": "inferior", "library_rel": 2e-13, "excel_rel": 1e-15},
                        "m:L2:B": {"verdict": "not inferior", "library_rel": 1e-15,
                                   "excel_rel": 1e-15}},
              "defect_points": ["m:L1:A"]}
    holdout = {"ib:P001:X": point("ib:P001:X", "library error" if library_error else "not inferior",
                                  ib_after),
               "dv:P002:Y": point("dv:P002:Y", dv_verdict, 1e-13, 1e-14)}
    floor = {"m:L1:A": point("m:L1:A", "not inferior", defect_after),
             "m:L2:B": point("m:L2:B", "inferior" if floor_new else "not inferior", 5e-14)}
    return holdout, floor, before


class Model(unittest.TestCase):
    def test_rule_region(self):
        self.assertTrue(M.rule_applies(50.0, 100.0))
        self.assertTrue(M.rule_applies(999.0, 0.5))
        self.assertFalse(M.rule_applies(0.49, 100.0))       # a shape below 1/2 keeps the literal form
        self.assertFalse(M.rule_applies(600.0, 400.0))      # Loader already

    def test_rule_repairs_the_registered_defect(self):
        x, a, b = 0.333333, 50.0, 100.0
        with mp.workdps(40):
            ref = mp.betainc(a, b, 0, mp.mpf(x), regularized=True)
            now = float(abs((mp.mpf(M.current(x, 1 - x, a, b)) - ref) / ref))
            rule = float(abs((mp.mpf(M.candidate(x, 1 - x, a, b)) - ref) / ref))
        self.assertGreater(now, 1e-13)
        self.assertLess(rule, 1e-14)

    def test_far_deviance_matches_its_definition(self):
        for x, mean in ((400.0, 274.8), (0.5, 285.0), (9000.0, 30.0)):
            with mp.workdps(40):
                exact = mp.mpf(x) * mp.log(mp.mpf(x) / mp.mpf(mean)) + mp.mpf(mean) - mp.mpf(x)
                got = M.candidate_deviance(x, mean)
                self.assertLess(float(abs(mp.mpf(got) - exact)), 3e-16 * max(x, mean) + 1e-15)


class Design(unittest.TestCase):
    def test_grid_matches_the_generator(self):
        self.assertEqual(S.build(check=True), 0)

    def test_every_ib_point_is_in_the_rule_region(self):
        for cls, fn, args in S.design():
            if cls == "ib":
                _, _, a, b = S.beta_call(fn, args)
                self.assertTrue(M.rule_applies(a, b), (fn, args))


class Rules(unittest.TestCase):
    def failures(self, **kw):
        return S.judge(*case(**kw))[0]

    def test_clean_case_passes(self):
        self.assertEqual(self.failures(), [])

    def test_h1_forward(self):
        self.assertTrue(any(f.startswith("H1") for f in self.failures(ib_after=3e-13)))

    def test_h2_newly_inferior_holdout_point(self):
        self.assertTrue(any(f.startswith("H2") for f in self.failures(dv_verdict="inferior")))

    def test_h3_library_error(self):
        self.assertTrue(any(f.startswith("H3") for f in self.failures(library_error=True)))

    def test_h4_defect_and_floor(self):
        self.assertTrue(any(f.startswith("H4") for f in self.failures(defect_after=5e-14)))
        self.assertTrue(any(f.startswith("H4") for f in self.failures(floor_new=True)))


class LibraryAlone(unittest.TestCase):
    """A point Excel did not answer still carries the library's own error."""

    def rows(self, library):
        return [{"kind": "hg:P001:Hypergeometric_PMF", "excel_formula": "F", "kstats_call": "K",
                 "reference": "0.25", "observed_excel": "ERROR", "observed_kstats": library}]

    def point(self, library):
        import struct
        rows = self.rows("ERROR" if library is None
                         else struct.pack(">d", library).hex().upper())
        points, _ = S.FC.judge(rows)
        return S.with_library(points, rows)[0]

    def test_not_compared_keeps_the_library_error(self):
        p = self.point(0.25 * (1 + 1e-12))
        self.assertEqual(p["verdict"], "not compared")
        self.assertAlmostEqual(p["library_rel"], 1e-12, delta=1e-15)

    def test_not_compared_and_no_library_number(self):
        p = self.point(None)
        self.assertEqual(p["verdict"], "not compared")
        self.assertTrue(p["library_missing"])
        self.assertNotIn("library_rel", p)


if __name__ == "__main__":
    unittest.main()
