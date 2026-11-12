"""Independent arithmetic and isolation checks for the P1 screening model."""
import unittest
from unittest.mock import patch
import run_tight_cf as tight
import run_extended_fit as extended
import mpmath as mp
import pair_cf_screen as candidate
import vba_order_model as model
from _ibeta import betacf


class PairCFTests(unittest.TestCase):
    def test_complement_retains_small_member(self):
        x=2.0**-54
        y=candidate.DD(1.)-candidate.DD(x)
        self.assertEqual(y.hi,1.)
        self.assertEqual(y.lo,-x)

    def test_fraction_against_reference(self):
        with mp.workdps(80):
            x=candidate.DD(.25)
            value,_=candidate.cf_pair(2.,3.,x)
            ref=betacf(mp.mpf(2),mp.mpf(3),mp.mpf('.25'))
            self.assertLess(float(abs(value-ref)/ref),1e-14)

    def test_state_restored(self):
        before=model.forward.beta_regularized
        candidate.evaluate(.9,100.,1e10)
        self.assertIs(model.forward.beta_regularized,before)

    def test_out_of_range_rejected(self):
        with self.assertRaises(ValueError):
            candidate.checked(candidate.DD(1e101))


class TightCFTests(unittest.TestCase):
    def test_stop_is_scoped_and_restored_on_error(self):
        original=model.forward.C['PROB_NUM_EPS']
        def fail(*args):
            self.assertEqual(model.forward.C['PROB_NUM_EPS'],2.0**-100)
            raise ValueError('deliberate failure')
        with patch.object(tight.p1,'evaluate',side_effect=fail):
            with self.assertRaises(ValueError):
                tight.evaluate(.9,100.,1e10)
        self.assertEqual(model.forward.C['PROB_NUM_EPS'],original)

    def test_new_rows_are_unique_and_disjoint(self):
        rows={(p,*dfs) for p,dfs in tight.NEW_ROWS}
        retained={(p,a,b) for p,a,b,_ in tight.baseline.AFFECTED+tight.baseline.CONTROLS}
        self.assertEqual(len(rows),8)
        self.assertFalse(rows & retained)


class CheckedOracleTests(unittest.TestCase):
    def test_uniform_beta_fraction_has_analytic_value(self):
        with mp.workdps(80):
            x = mp.mpf('.25')
            actual = extended.checked_betacf(mp.mpf(1), mp.mpf(1), x)
            self.assertLess(abs(actual - 1/(1-x)), mp.mpf('1e-70'))

    def test_exhaustion_does_not_return_partial_fraction(self):
        with patch.object(extended, 'range', return_value=range(1, 2), create=True):
            with mp.workdps(80), self.assertRaisesRegex(ArithmeticError, 'exhausted'):
                extended.checked_betacf(mp.mpf('2.5'), mp.mpf('3.5'), mp.mpf('.25'))


if __name__=='__main__':
    unittest.main()
