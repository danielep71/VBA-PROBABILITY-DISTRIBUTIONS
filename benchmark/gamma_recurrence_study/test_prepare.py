"""Guard candidate scope, baseline identity and evidence rejection behavior."""
import json
import tempfile
import unittest
from pathlib import Path
import prepare as P
import summarize as S


class RecurrenceStudyTests(unittest.TestCase):
    def test_reversible_pinned_candidates(self):
        source = P.SOURCE.read_text()
        self.assertEqual(P.digest(source), P.BASELINE)
        for mode in S.MODES:
            candidate = P.transform(source, mode)
            self.assertEqual(P.transform(candidate, mode, reverse=True), source)
        combined = P.transform(source, 'combined')
        self.assertEqual(combined, P.transform(P.transform(source, 'lower'), 'upper'))
        # Arithmetic implementations and the small-shape path are untouched.
        for marker in ('Private Function PROB_DD_SG_Op(', 'Private Function PROB_DD_EG_LogPair('):
            tail = source[source.index(marker):]
            block = tail[:tail.index('End Function\n') + len('End Function\n')]
            self.assertIn(block, combined)

    def test_source_drift_is_rejected(self):
        source = P.SOURCE.read_text()
        with self.assertRaises(ValueError):
            P.transform(P.transform(source, 'lower'), 'lower')
        with self.assertRaises(ValueError):
            P.transform(source.replace('    Stage = "lower geometric remainder"', '    Stage = "changed"'), 'lower')

    def test_incomplete_or_bit_drift_evidence_is_rejected(self):
        rows = [dict(a=a, x=x, route=r, mode=m, sample=n, cold=False,
                     calls=1 if r == 'inverse' else 4, elapsed_us=100., bits='3FE0000000000000')
                for a, x, r in S.CASES for m in S.MODES for n in (1, 2, 3)]
        data = dict(status='PASS', error='', candidate_sha='0' * 40, samples=rows)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'fixture.json'
            path.write_text(json.dumps(data))
            S.summarize(path, emit=False)
            rows[0]['bits'] = '3FE0000000000001'
            path.write_text(json.dumps(data))
            with self.assertRaises(AssertionError):
                S.summarize(path, emit=False)
            rows.pop(0)
            path.write_text(json.dumps(data))
            with self.assertRaises(AssertionError):
                S.summarize(path, emit=False)


if __name__ == '__main__':
    unittest.main()
