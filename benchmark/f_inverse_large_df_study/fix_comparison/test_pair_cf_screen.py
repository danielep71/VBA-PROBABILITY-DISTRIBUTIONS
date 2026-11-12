"""Independent arithmetic and isolation checks for the P1 screening model."""
import unittest
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


if __name__=='__main__':
    unittest.main()
