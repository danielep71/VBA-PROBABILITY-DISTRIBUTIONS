"""Negative controls for B5 provenance, B4 comparison and complete evidence."""
import copy
import unittest

import pair_candidate
from prepare_pair_host import hex_bits
from validate_scalar_host import costs, validate


class ScalarEvidence(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.design = dict(study='B5_scalar', baseline='B4', candidate_sha='1' * 40, cases=[])
        rows, timings = [], []
        points = ((.0001, .5), (.5, .25), (2.5, 1.), (2.5, 6.),
                  (100., 90.), (100., 110.), (10000., 9900.), (10000., 10100.))
        for a, x in points:
            p, q = map(hex_bits, pair_candidate.public(a, x))
            key = hex_bits(a) + ':' + hex_bits(x)
            cls.design['cases'].append(dict(id=key, expected=p + ',' + q))
            rows.append(dict(id=key, raw=p + ',' + q + ',0,' + key.replace(':', ',')))
            for n in (1, 2, 3):
                for b in (True, False):
                    # Deliberately synthetic validator controls, never performance evidence.
                    timings.append(dict(a=a, x=x, round=n, baseline=b, calls=4,
                                        p=p, q=q, elapsed_us=100. if b else 50.))
        key = hex_bits(10001.) + ':' + hex_bits(10001.)
        cls.design['cases'].append(dict(id=key, expected='DECLINED'))
        rows.append(dict(id=key, raw='DECLINED,' + key.replace(':', ',')))
        cls.report = dict(status='PASS', candidate_sha='1' * 40, error='', cleanup_error='',
                          normalization_check='PASS checks=11', arithmetic_check='PASS checks=576',
                          rows=rows, timings=timings)

    def test_complete_and_costs(self):
        points = validate(self.report, self.design)
        self.assertTrue(all(r['reduction'] == .5 for r in costs(self.report, points)))

    def test_negative_controls(self):
        variants = []
        for key, value in (('arithmetic_check', ''), ('candidate_sha', '2' * 40),
                           ('normalization_check', ''), ('cleanup_error', 'failure')):
            bad = copy.deepcopy(self.report)
            bad[key] = value
            variants.append(bad)
        for mutation in ('missing_row', 'duplicate_row', 'missing_timing', 'zero_timing', 'baseline_bits', 'echo'):
            bad = copy.deepcopy(self.report)
            if mutation == 'missing_row':
                bad['rows'].pop()
            elif mutation == 'duplicate_row':
                bad['rows'][1] = bad['rows'][0]
            elif mutation == 'missing_timing':
                bad['timings'].pop()
            elif mutation == 'zero_timing':
                bad['timings'][0]['elapsed_us'] = 0.
            elif mutation == 'baseline_bits':
                for row in bad['timings']:
                    if row['baseline']:
                        row['p'] = '0000000000000000'
            else:
                bad['rows'][0]['raw'] = bad['rows'][0]['raw'][:-1] + 'F'
            variants.append(bad)
        for bad in variants:
            with self.assertRaises(AssertionError):
                validate(bad, self.design)


if __name__ == '__main__':
    unittest.main()
