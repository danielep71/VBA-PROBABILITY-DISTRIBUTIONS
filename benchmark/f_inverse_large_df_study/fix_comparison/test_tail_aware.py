"""P4: the two tails, the tail-aware residual and the screen's design."""
import math
import unittest

import pair_cf_screen as pair
import run_tail_aware as runner
import stagnation_model as p3
import tail_aware_model as p4


class TailAwareTests(unittest.TestCase):
    def test_the_two_tails_sum_to_one_and_the_small_one_is_direct(self):
        lower, upper = p4.tails(1e-9, 2.5, 1e6)
        self.assertEqual(upper, 1.0 - lower)
        lower, upper = p4.tails(5e-5, 2.5, 1e6)
        self.assertEqual(lower, 1.0 - upper)
        self.assertLess(upper, 1e-15)
        self.assertGreater(upper, 0.0)

    def test_the_model_state_is_restored(self):
        saved_model, saved_eps = pair.model, p3.forward.C['PROB_NUM_EPS']
        p4.tails(0.3, 4.0, 9.0)
        self.assertIs(pair.model, saved_model)
        self.assertEqual(p3.forward.C['PROB_NUM_EPS'], saved_eps)

    def test_the_far_upper_tail_keeps_its_accuracy(self):
        # P3 returns 14.0105 here, wrong in the fifth digit; 60-digit value 14.010326047225420...
        value = p4.evaluate(1 - 1e-13, 5.0, 2e6)[0]
        self.assertAlmostEqual(value / 14.01032604722542, 1.0, delta=1e-14)
        self.assertGreater(abs(p3.evaluate(1 - 1e-13, 5.0, 2e6)[0] / 14.01032604722542 - 1.0), 1e-6)

    def test_the_affected_rows_of_the_issue_stay_repaired(self):
        self.assertEqual(p4.evaluate(0.9, 100.0, 1e10)[0], p3.evaluate(0.9, 100.0, 1e10)[0])

    def test_edges(self):
        self.assertEqual(p4.inverse_pair(0.0, 1.0, 2.0, 3.0)[:2], (0.0, 1.0))
        self.assertEqual(p4.inverse_pair(1.0, 0.0, 2.0, 3.0)[:2], (1.0, 0.0))

    def test_stage_two_is_new(self):
        rows = runner.new_rows()
        self.assertEqual(len(rows), 47)
        keys = {(p, d1, d2) for _, p, d1, d2 in rows}
        self.assertEqual(len(keys), 47)
        self.assertEqual(keys & set(runner.PROBE_ROWS), set())
        self.assertTrue(all(0.0 < p < 1.0 for p, _, _ in keys))
        self.assertEqual(sorted({k for k, *_ in rows}),
                         ['far', 'middle', 'orientation_seam', 'tail_seam'])

    def test_step(self):
        self.assertEqual(runner._step(0.5, 1), math.nextafter(0.5, 1.0))
        self.assertEqual(runner._step(0.5, -4), 0.5 - 4 * 2.0 ** -54)


if __name__ == '__main__':
    unittest.main()
