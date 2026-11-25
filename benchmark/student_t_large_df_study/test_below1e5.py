"""#365: tests of the below-1E5 study's tools. The rules are tested on rows made up here."""
import math
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import below1e5_model as M  # noqa: E402
import below1e5_study as B  # noqa: E402
import seam_study as seam  # noqa: E402


def row(fn, a1, df, ref, value, regime):
    return {"function": fn, "arg1": repr(a1), "arg2": repr(df), "reference": repr(ref),
            "observed_vba": "" if value is None else f"{value!r};0E+000", "regime": regime}


def exported(served_error=0.0, unserved_shift=0.0, inverse_error=0.0, seam_rise=0.0,
             blank=False):
    """Rows that satisfy every rule, then perturbed."""
    maps = [row("StudentT_Survival", 4.0, 5e4, 3e-5, 3e-5 * (1.0 + served_error), B.REGIME),
            row("StudentT_Survival", 1.0, 5e4, 0.158, 0.158, B.REGIME)]
    before = {B.key(maps[1]): f"{0.158 * (1.0 + unserved_shift)!r};0E+000"}
    holdout = [row("StudentT_Survival", 6.2, 9999.0, 3e-10, None if blank else 3e-10, B.HOLDOUT),
               row("StudentT_InverseCumulative", 3e-10, 9999.0, -6.2,
                   -6.2 * (1.0 + inverse_error), B.HOLDOUT)]
    probes = [row("StudentT_Survival", t, df, 1e-50, 1e-50, B.SEAM)
              for df in B.T_SEAM_DF for t in B.t_neighbours()]
    probes += [row("StudentT_Survival", t, df, 1e-5, 1e-5, B.SEAM)
               for t in B.DF_EDGE_T for df in B.df_neighbours()]
    seam210 = []
    below, at, above = seam.df_neighbours()
    for t in B.DF_SEAM_T:
        for x in (t, -t):
            for fn in ("StudentT_Survival", "StudentT_Cumulative"):
                small = (fn == "StudentT_Survival") == (x > 0)
                value = 1e-3 if small else 1.0 - 1e-3
                rising = value * (1.0 + seam_rise) if (small and t == 2.0) else value
                seam210 += [row(fn, x, below, value, value, "seam"),
                            row(fn, x, at, value, rising, "seam"),
                            row(fn, x, above, value, rising, "seam")]
    return maps, holdout, probes, seam210, before


class Region(unittest.TestCase):
    def test_served_region(self):
        self.assertTrue(M.served(2.0, 5e4))
        self.assertTrue(M.served(16.0, 1999.0))
        self.assertFalse(M.served(math.nextafter(16.0, math.inf), 5e4))   # above the cap
        self.assertFalse(M.served(2.0, 1e5))                              # the expansion's df
        self.assertFalse(M.served(2.0, math.nextafter(1999.0, 0.0)))      # a + b below 1000
        self.assertFalse(M.served(1.73, 5e4))                             # above the mean
        self.assertTrue(M.served(1.74, 1999.0))

    def test_every_holdout_row_is_served(self):
        for df in B.HOLDOUT_DF:
            for t in B.HOLDOUT_T:
                self.assertTrue(M.served(t, df), (t, df))

    def test_row_counts(self):
        self.assertEqual(len(B.holdout_rows()), 144 + 36 + 12)
        self.assertEqual(len(B.seam_rows()), 9 + 6)

    def test_rule_improves_the_worst_map_point(self):
        df = math.nextafter(1e5, 0.0)
        with seam.mp.workdps(60):
            ref = seam.upper_tail(2.0, df)
            rule = float(abs((seam.mp.mpf(M.rule_tail(2.0, df)) - ref) / ref))
        self.assertLess(rule, 1e-14)


class Rules(unittest.TestCase):
    def failures(self, **kw):
        return B.judge(*exported(**kw))[0]

    def test_clean_rows_pass(self):
        self.assertEqual(self.failures(), [])

    def test_g1_served_forward(self):
        self.assertTrue(any(f.startswith("G1") for f in self.failures(served_error=2e-13)))
        self.assertEqual(self.failures(served_error=5e-14), [])

    def test_g2_unserved_unchanged(self):
        self.assertTrue(any(f.startswith("G2") for f in self.failures(unserved_shift=1e-15)))

    def test_g3_df_seam(self):
        self.assertTrue(any(f.startswith("G3") for f in self.failures(seam_rise=2e-13)))
        self.assertEqual(self.failures(seam_rise=5e-14), [])

    def test_g4_inverse(self):
        self.assertTrue(any(f.startswith("G4") for f in self.failures(inverse_error=2e-12)))

    def test_g5_numbers(self):
        self.assertTrue(any(f.startswith("G5") for f in self.failures(blank=True)))


if __name__ == "__main__":
    unittest.main()
