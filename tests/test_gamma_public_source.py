"""Bind the focused public port to the arithmetic and regressions tested in #242."""
import hashlib
import json
from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
PROOF = json.loads((ROOT / 'benchmark/gamma_public_integration/provenance.json').read_text())


def digest(text):
    return hashlib.sha256(text.encode()).hexdigest()


def procedure(text, name):
    pattern = r'(?ms)^(?:Private|Public) (Function|Sub) ' + re.escape(name) + r'\b.*?^End \1\n'
    matches = list(re.finditer(pattern, text))
    if len(matches) != 1:
        raise AssertionError(f'Expected one procedure: {name}')
    return matches[0].group()


def block(text, name):
    start = "' BEGIN " + name + '\n'
    end = "' END " + name + '\n\n'
    if text.count(start) != 1 or text.count(end) != 1:
        raise AssertionError(f'Expected one marked block: {name}')
    return start + text.split(start)[1].split(end)[0] + end


class PublicGammaPortTests(unittest.TestCase):
    def test_arithmetic_and_adapters_match_tested_port(self):
        source = (ROOT / 'src/M_STATS_PROBDIST_SPECIALFUNCS.bas').read_text()
        self.assertEqual(digest(block(source, 'generated public Gamma integration')),
                         PROOF['private_arithmetic_sha256'])
        self.assertEqual(digest(block(source, 'public Gamma bit types')),
                         PROOF['private_types_sha256'])
        for tail, sha in PROOF['adapter_sha256'].items():
            self.assertEqual(digest(procedure(source, 'PROB_TryGammaRegularized' + tail)), sha)

    def test_original_public_failures_and_resource_controls_remain_live(self):
        source = (ROOT / 'tests/M_STATS_PROBDIST_TEST.bas').read_text()
        self.assertEqual(digest(block(source, 'public Gamma regression coverage')),
                         PROOF['regression_block_sha256'])
        for name, sha in PROOF['regression_functions'].items():
            self.assertEqual(digest(procedure(source, name)), sha)
        for name in PROOF['entry_calls']:
            self.assertEqual(source.count('    ' + name + '\n'), 1)
        policy = json.loads((ROOT / '.github/excel-evidence-policy.json').read_text())
        self.assertEqual(PROOF['baseline_assertions'] + PROOF['added_assertions'], 1947)
        self.assertEqual(policy['expected_assertions'], PROOF['expected_assertions'])
        self.assertEqual(policy['expected_assertions'], 1947)


if __name__ == '__main__':
    unittest.main()
