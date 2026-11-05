"""Convergence checks for the large-shape CDF generator (#167).

The generator certifies each reference by evaluating it at two precisions and
requiring agreement. Two precisions can truncate an iterative evaluation at the
same term and still agree, so every series and continued fraction must report
reaching its cap, and the generator must drop such a point instead of comparing
the two truncated results.

Run from this directory or from benchmark/:

    python test_generate_cdf_large_shape.py
"""

from __future__ import annotations

import os
import sys
import unittest

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import generate_cdf_large_shape as g  # noqa: E402


class CappedEvaluations(unittest.TestCase):
    """Each iterative routine raises instead of returning a truncated value."""

    def test_series_cap_raises(self) -> None:
        with mp.workdps(50):
            with self.assertRaises(g.NotConverged):
                g.gamma_P_series(mp.mpf(1e4), mp.mpf(1e4), cap=10)

    def test_gamma_fraction_cap_raises(self) -> None:
        with mp.workdps(50):
            with self.assertRaises(g.NotConverged):
                g.gamma_Q_cf(mp.mpf(1e4), mp.mpf(1.01e4), cap=3)

    def test_beta_fraction_cap_raises(self) -> None:
        with mp.workdps(50):
            with self.assertRaises(g.NotConverged):
                g.beta_cf(mp.mpf(1e4), mp.mpf(1e4), mp.mpf("0.49"), cap=3)

    def test_converged_values_match_mpmath(self) -> None:
        # Well inside the caps, the routines still return the right values.
        with mp.workdps(50):
            a, x = mp.mpf(100), mp.mpf(95)
            p = g.gamma_P_series(a, x)
            q = g.gamma_Q_cf(a, mp.mpf(105))
            self.assertLess(
                abs(p - mp.gammainc(a, 0, x, regularized=True)) / p, mp.mpf(10) ** -40
            )
            q_ref = mp.gammainc(a, mp.mpf(105), mp.inf, regularized=True)
            self.assertLess(abs(q - q_ref) / q_ref, mp.mpf(10) ** -40)
            i = g.beta_I(mp.mpf(50), mp.mpf(50), mp.mpf("0.45"))
            i_ref = mp.betainc(50, 50, 0, mp.mpf("0.45"), regularized=True)
            self.assertLess(abs(i - i_ref) / i_ref, mp.mpf(10) ** -40)


class ReportedCase(unittest.TestCase):
    """The case in the review finding: A = X = 1E12."""

    def test_series_at_1e12_is_rejected_not_truncated(self) -> None:
        # With the real cap, the series stops after 3,000,000 terms near
        # 0.49865 while the true value is about 0.5000001. Both precisions
        # agreed on that truncated sum; now the evaluation raises.
        A = mp.mpf(1e12)
        with mp.workdps(g.BASE_DPS + 12):
            with self.assertRaises(g.NotConverged):
                g.gamma_PQ(A, A)


if __name__ == "__main__":
    unittest.main(verbosity=2)
