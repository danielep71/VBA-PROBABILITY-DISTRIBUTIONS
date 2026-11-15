"""#213: tests of the above-8 study's tools. The rules are tested on rows made up here."""
import math
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import above8_model as M  # noqa: E402
import above8_study as A  # noqa: E402
import seam_study as seam  # noqa: E402

ABOVE = math.nextafter(8.0, math.inf)


def row(fn, a1, df, ref, value):
    return {"function": fn, "arg1": repr(a1), "arg2": repr(df), "reference": repr(ref),
            "observed_vba": "" if value is None else f"{value!r};0E+000", "regime": "x"}


def seam_rows(shift=0.0, inverse_error=0.0):
    """Seam rows that satisfy every rule, then perturbed."""
    out = []
    for df in seam.T_SEAM_DF:
        for t in seam.t_neighbours():
            tail = 6e-16 * (1.0 - 1e-14 * (t - 8.0) / (ABOVE - 8.0))
            wrong = tail * (1.0 + shift) if t == ABOVE else tail
            out.append(row("StudentT_Survival", t, df, tail, wrong))
            out.append(row("StudentT_Cumulative", t, df, 1.0 - tail, 1.0 - tail))
            out.append(row("StudentT_Survival", -t, df, 1.0 - tail, 1.0 - tail))
            out.append(row("StudentT_Cumulative", -t, df, tail, wrong))
        out.append(row("StudentT_InverseCumulative", 6e-16, df, -8.0, -8.0 * (1.0 + inverse_error)))
    return out


def probes(error=0.0, t=10.0, blank=False):
    return [row("StudentT_Survival", t, 1e7, 1e-24, None if blank else 1e-24 * (1.0 + error)),
            row("StudentT_InverseCumulative", 1e-24, 1e7, -10.0, -10.0)]


def rules(seams, probe_rows):
    return sorted({text.split()[0] for text in A.judge(seams, probe_rows)[0]})


class Rules(unittest.TestCase):
    def test_clean(self):
        self.assertEqual(rules(seam_rows(), probes()), [])

    def test_forward_at_the_seam(self):
        self.assertEqual(rules(seam_rows(shift=-5e-13), probes()), ["F1"])

    def test_monotone(self):
        self.assertEqual(rules(seam_rows(shift=5e-14), probes()), ["F2"])
        self.assertEqual(rules(seam_rows(shift=5e-13), probes()), ["F1", "F2"])

    def test_inverse(self):
        self.assertEqual(rules(seam_rows(inverse_error=2e-12), probes()), ["F3"])
        self.assertEqual(rules(seam_rows(inverse_error=2e-14), probes()), [])

    def test_probe_gate_stops_at_14(self):
        self.assertEqual(rules(seam_rows(), probes(error=5e-13)), ["F1"])
        self.assertEqual(rules(seam_rows(), probes(error=5e-13, t=20.0)), [])

    def test_blank(self):
        self.assertEqual(rules(seam_rows(), probes(blank=True)), ["F4"])


class Design(unittest.TestCase):
    def test_rows(self):
        rows = A.grid_rows(A.REGIME)
        self.assertEqual(len(rows), 32)
        self.assertEqual(sum(r["function"] == "StudentT_Survival" for r in rows), 24)
        self.assertTrue(all(abs(float(r["arg1"])) > 8.0 or r["function"] != "StudentT_Survival"
                            for r in rows))


class Model(unittest.TestCase):
    def test_candidate_serves_above_8_and_declines_elsewhere(self):
        self.assertIsNotNone(M.candidate_tail(ABOVE, 1e7))
        self.assertIsNone(M.candidate_tail(400.0, 1e5))     # y >= 1/2
        self.assertEqual(M.tail(8.0, 1e7), M.ldf.route_tail(8.0, 1e7))

    def test_candidate_is_accurate_and_monotone_at_the_seam(self):
        for df in (1e5, 1e7, 1e8):
            reference = seam.upper_tail(ABOVE, df)
            self.assertLess(abs(M.tail(ABOVE, df) - reference) / reference, 2e-14)
            self.assertLessEqual(M.tail(ABOVE, df), M.tail(8.0, df))

    def test_the_source_of_today_is_not(self):
        reference = seam.upper_tail(ABOVE, 1e8)
        self.assertGreater(abs(M.current_tail(ABOVE, 1e8) - reference) / reference, 1e-11)


if __name__ == "__main__":
    unittest.main()
