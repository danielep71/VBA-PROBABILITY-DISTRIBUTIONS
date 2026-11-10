"""Numerical primitive and range fixtures for the exploratory precision path."""
import math
import unittest
import mpmath as mp
from precision_probe import DD, log_dd, log1p_dd, expm1_dd, lower


class PrecisionProbeTests(unittest.TestCase):
    def test_transcendentals_against_high_precision(self):
        cases = [
            (log_dd, mp.log, [2**-400, .00390625, .7, 1.973]),
            (lambda x: log1p_dd(DD(x)), mp.log1p, [-1e-120, -1e-12, -.1, -.5, -.8]),
            (lambda x: expm1_dd(DD(x)), mp.expm1, [-1e-120, -1e-12, -.1, -1., -8.]),
        ]
        with mp.workdps(100):
            for fn, oracle, xs in cases:
                for x in xs:
                    with self.subTest(x=x, oracle=oracle.__name__):
                        actual = fn(x)
                        target = oracle(mp.mpf(x))
                        error = abs(mp.mpf(actual.hi) + mp.mpf(actual.lo) - target)
                        self.assertLessEqual(error, abs(target) * mp.mpf(2)**-96)

    def test_rejected_v1_log_one_counterexample_is_preserved(self):
        # This is a negative fixture for the frozen rejected prototype.
        # The accepted identity is tested independently against v2.
        self.assertEqual(log_dd(1.).value(), 2**-108)
        self.assertNotEqual(log_dd(1.).value(), 0.)

    def test_cancellation_and_residuals(self):
        self.assertEqual(((DD(1) + 2**-80) - 1).value(), 2**-80)
        with mp.workdps(100):
            for a, b in [(.1, .7), (1e-120, .0043), (1.973, .973), (2**-400, 1.5)]:
                for actual, target in [(DD(a) * b, mp.mpf(a) * mp.mpf(b)),
                                       (DD(a) / b, mp.mpf(a) / mp.mpf(b))]:
                    error = abs(mp.mpf(actual.hi) + mp.mpf(actual.lo) - target)
                    self.assertLessEqual(error, abs(target) * mp.mpf(2)**-98)

    def test_outside_exploratory_range_rejected(self):
        for x in [0., -1., math.inf, math.nan, 2**-501]:
            with self.assertRaises(ValueError):
                log_dd(x)
        for x in [.1, -.81, math.nan]:
            with self.assertRaises(ValueError):
                log1p_dd(DD(x))
        for x in [.1, -8.1, math.inf]:
            with self.assertRaises(ValueError):
                expm1_dd(DD(x))
        for a, x in [(0, 1), (1.1, 1), (.5, 1.5), (1e-200, .1)]:
            with self.assertRaises(ValueError):
                lower(a, x, 0.)


if __name__ == '__main__':
    unittest.main()
