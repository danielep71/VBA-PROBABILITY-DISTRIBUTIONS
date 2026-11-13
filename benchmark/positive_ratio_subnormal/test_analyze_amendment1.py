#!/usr/bin/env python3
"""
Fixtures for analyze_amendment1.py (#13, Amendment 1).

Synthetic points only: each is built here from exact rationals, so a failure
indicts the analyzer and not an export. They prove that the class of a point
comes from its operands, that a mislabelled or off-design point is refused,
and that each claim both passes and fails when it should.

Run: python3 test_analyze_amendment1.py   (exit 0 = pass)
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import analyze_amendment1 as M  # noqa: E402
import analyze_positive_ratio_subnormal as A  # noqa: E402

MINSUB = 5e-324
SHAPE = 0.57735026918963


def point(family, bits, construction, numerator, scale, cur_err, cand_err, shape=SHAPE,
          point_id=None, surface="cumulative"):
    """X = numerator * 2^-1074 over `scale`; current and candidate are the
    reference moved by the given relative errors."""
    x = numerator * MINSUB
    stored = x / scale
    p = A.Point(family=family, surface=surface,
                point_id=point_id or f"{construction}_k{bits}", construction=construction,
                shape_id="s", bucket_bits=bits, x=x, shape_input=shape, scale=float(scale),
                stored=stored, log_stored=None, current_status="OK", current=None,
                candidate_status="OK", candidate=None)
    ref = A.reference(p)
    p.current = float(ref * (1 + cur_err))
    p.candidate = float(ref * (1 + cand_err))
    return p


def gamma_set(rounded_low=(1e-9, 1e-13), exact=(1e-16, 5e-14), rounded_high=(2e-13, 1e-13),
              scaled_cand=5e-14):
    n39, n43 = 2 ** 38, 2 ** 42
    return [
        point("gamma", 39, "rounded", 5 * n39 + 2, 5, *rounded_low),
        point("gamma", 39, "exact_unit", n39, 1, *exact),
        point("gamma", 39, "exact_scaled", 4 * n39, 4, exact[0], scaled_cand),
        point("gamma", 43, "rounded", 7 * n43 + 3, 7, *rounded_high),
        point("gamma", 43, "exact_unit", n43, 1, *exact),
    ]


def run(points):
    lines = []
    return M.test_claims(points, out=lines.append), "\n".join(lines)


class Amendment1(unittest.TestCase):
    def test_class_comes_from_the_operands(self):
        pts = gamma_set()
        self.assertEqual([M.group(p) for p in pts],
                         ["rounded", "exact_unit", "exact_scaled", "rounded", "exact_unit"])
        self.assertEqual(M.verify(pts), [])
        tie = point("chisquare", 39, "rounded", 2 * 2 ** 38 + 1, 2, 1e-9, 1e-13)
        even = point("chisquare", 39, "exact_scaled", 2 * 2 ** 38, 2, 1e-16, 1e-13)
        self.assertEqual([M.group(tie), M.group(even)], ["rounded", "exact_scaled"])
        self.assertEqual(M.verify([tie, even]), [])

    def test_a_mislabelled_or_off_design_point_is_refused(self):
        wrong_label = point("gamma", 39, "exact_unit", 5 * 2 ** 38 + 2, 5, 1e-9, 1e-13)
        self.assertTrue(any("labelled exact_unit" in m for m in M.verify([wrong_label])))
        off_bucket = point("gamma", 40, "exact_unit", 2 ** 39, 1, 1e-16, 1e-13)
        self.assertTrue(any("not a registered bucket" in m for m in M.verify([off_bucket])))
        off_offset = point("gamma", 39, "rounded", 3 * 2 ** 38 + 2, 3, 1e-9, 1e-13)
        self.assertTrue(any("not a registered offset" in m for m in M.verify([off_offset])))
        density = point("gamma", 39, "exact_unit", 2 ** 38, 1, 1e-16, 1e-13, surface="density")
        self.assertTrue(any("only the cumulative" in m for m in M.verify([density])))

    def test_the_wired_classifier_follows_the_amendment(self):
        n39, n43 = 2 ** 38, 2 ** 42
        cases = [("gamma", 39, "exact_unit", n39, 1, "direct"),
                 ("gamma", 39, "exact_scaled", 4 * n39, 4, "direct"),
                 ("gamma", 39, "exact_scaled", 2 * n39, 2, "direct"),
                 ("gamma", 39, "rounded", 5 * n39 + 2, 5, "recovery"),
                 ("gamma", 43, "rounded", 7 * n43 + 3, 7, "direct"),
                 ("chisquare", 39, "exact_scaled", 2 * n39, 2, "direct"),
                 ("chisquare", 39, "rounded", 2 * n39 + 1, 2, "recovery"),
                 ("chisquare", 43, "rounded", 2 * n43 + 1, 2, "direct")]
        for family, bits, construction, numerator, scale, path in cases:
            p = point(family, bits, construction, numerator, scale, 1e-16, 1e-13)
            self.assertEqual(M.wired_path(p), path, (family, bits, construction, scale))
            self.assertEqual(M.expected_path(p), path)
        # A subnormal X with a fractional scale can multiply back to X from a
        # rounded quotient (3.2 ULP stored as 3, and 3 * 2.5 rounds to 8), so
        # it is never taken as exact.
        tricky = point("gamma", 39, "rounded", 8, 2.5, 1e-9, 1e-13)
        self.assertEqual(tricky.stored * 2.5, tricky.x)
        self.assertFalse(M.is_exact(tricky))
        self.assertEqual(M.wired_path(tricky), "recovery")
        # A normal X over a large scale, the ordinary route to the defect.
        wide = point("gamma", 39, "rounded", (3 * n39 + 1) * 2 ** 300, 3 * 2.0 ** 300, 1e-9, 1e-13)
        self.assertEqual(M.wired_path(wide), "recovery")

    def test_all_claims_hold_on_the_expected_pattern(self):
        ok, text = run(gamma_set())
        self.assertTrue(ok, text)
        for name in ("C1", "C2", "C3", "C4"):
            self.assertIn(f"-- {name}: PASS", text)

    def test_each_claim_can_fail(self):
        cases = {"C1": dict(rounded_low=(1e-14, 1e-13)),
                 "C2": dict(exact=(1e-12, 5e-14)),
                 "C3": dict(rounded_high=(6e-13, 1e-13)),
                 "C4": dict(scaled_cand=4e-13)}
        for name, change in cases.items():
            ok, text = run(gamma_set(**change))
            self.assertFalse(ok, name)
            self.assertIn(f"-- {name}: FAIL", text)

    def test_a_missing_value_is_pending_not_a_pass(self):
        pts = gamma_set()
        pts[0].candidate_status, pts[0].candidate = "ERROR", None
        ok, text = run(pts)
        self.assertFalse(ok)
        self.assertIn("PENDING", text)

    def test_chi_square_has_no_c4(self):
        n39, n43 = 2 ** 38, 2 ** 42
        pts = [point("chisquare", 39, "rounded", 2 * n39 + 1, 2, 1e-9, 1e-13),
               point("chisquare", 39, "exact_scaled", 2 * n39, 2, 1e-16, 5e-14),
               point("chisquare", 43, "rounded", 2 * n43 + 1, 2, 2e-13, 1e-13)]
        ok, text = run(pts)
        self.assertTrue(ok, text)
        self.assertNotIn("C4", text)

    def test_the_committed_baselines_match_the_design(self):
        """Key columns only, until the export: buckets, classes and counts."""
        import csv
        here = Path(__file__).resolve().parent
        for name, per_bucket, shapes in (("gamma_holdout2.csv", 6, 8),
                                         ("chisquare_holdout2.csv", 2, 8)):
            with (here / name).open(newline="") as handle:
                rows = list(csv.DictReader(handle))
            self.assertEqual(len(rows), per_bucket * shapes * len(M.BUCKETS), name)
            self.assertEqual({int(r["bucket_bits"]) for r in rows}, set(M.BUCKETS))
            self.assertEqual({r["surface"] for r in rows}, {"cumulative"})


if __name__ == "__main__":
    unittest.main()
