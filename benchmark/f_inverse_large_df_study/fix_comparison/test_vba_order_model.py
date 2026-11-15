"""Focused controls independent of the retained F observations."""
import math
import unittest

import mpmath as mp
import vba_order_model as model


class ModelTests(unittest.TestCase):
    def test_delta_independent_reference(self):
        with mp.workdps(80):
            for z, s in [(5e9, 50.), (5e7, 5.), (5e5, 1.5), (10., .75)]:
                ref = mp.loggamma(mp.mpf(z)+s)-mp.loggamma(z)
                self.assertLess(float(abs(model.log_gamma_delta(z,s)-ref)/abs(ref)), 1e-13)

    def test_uniform_quantile(self):
        for p in [.1, .4, .6, .9]:
            x, y, trace = model.inverse_pair(p, 1-p, 1., 1.)
            self.assertLess(abs(x-p), 1e-14)
            self.assertEqual(x+y, 1.)
            self.assertTrue(trace[-1]['settled'])

    def test_supplied_complement_boundary(self):
        self.assertEqual(model.inverse_pair(.9, 0., 2., 3.), (1.,0.,[]))
        self.assertEqual(model.inverse_pair(0., .9, 2., 3.), (0.,1.,[]))

    def test_exp_contract(self):
        self.assertIsNone(model.try_exp(math.inf))
        self.assertIsNone(model.try_exp(1000.))
        self.assertEqual(model.try_exp(-1000.), 0.)

    def test_model_isolation(self):
        other = model.load_model('independent_beta', model.ROOT/'benchmark/beta_balanced_survival/beta_route_model.py')
        with self.assertRaises(other.NotModelled):
            other.log_beta(50.,5e9)
        self.assertTrue(math.isfinite(model.log_beta(50.,5e9)))


if __name__ == '__main__':
    unittest.main()
