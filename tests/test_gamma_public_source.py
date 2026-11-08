"""Bind the focused public port to the arithmetic and regressions tested in #242."""
import base64
import hashlib
import json
from pathlib import Path
import re
import subprocess
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

    def test_exact_source_focused_excel_evidence(self):
        artifact = json.loads((ROOT / 'benchmark/gamma_public_integration/excel_artifact.json').read_text())
        self.assertEqual(artifact['run_id'], 42524911524)
        self.assertEqual(artifact['artifact_id'], 13634354530)
        self.assertEqual(artifact['artifact_zip_sha256'],
                         '719620d6c0ab75f74a75819824bbcaf6b2f0a06f29c25027b80c75fc9a33d790')
        decoded = {}
        for name, member in artifact['members'].items():
            data = base64.b64decode(member['base64'], validate=True)
            self.assertEqual(hashlib.sha256(data).hexdigest(), member['sha256'])
            decoded[name] = data
        record = json.loads(decoded['excel-certification.json'].decode('utf-8-sig'))
        self.assertEqual(record['harness'], dict(entry_point='Test_STATS_PROBDIST_RunAll_CI',
                                               assertions=1947, passed=1947, failed=0))
        self.assertTrue(all(stage['status'] == 'PASS' for stage in record['stages'].values()))
        self.assertTrue(all(not grid['exported'] for grid in record['grids'].values()))
        self.assertEqual(record['log']['sha256'], 'sha256:' + hashlib.sha256(decoded['test-result.txt']).hexdigest())
        merge = artifact['tested_merge']
        raw = base64.b64decode(merge['commit_object_base64'], validate=True)
        self.assertEqual(hashlib.sha1(b'commit ' + str(len(raw)).encode() + b'\0' + raw).hexdigest(), merge['sha'])
        self.assertEqual(record['candidate_sha'], merge['sha'])
        self.assertEqual(merge['sha'], 'f214de606b180e3b3781c5c9806aa051c7e636f4')
        self.assertEqual(record['runner']['workflow']['sha'], merge['sha'])
        self.assertEqual(record['runner']['workflow']['run_id'], artifact['run_id'])
        head = 'fe4c90f47e7b58bd089d720046d50fdac9b4581e'
        self.assertEqual(merge['head_sha'], head)
        self.assertEqual(raw.decode().splitlines()[0], 'tree ' + merge['tree_sha'])
        self.assertEqual(subprocess.check_output(['git', 'rev-parse', head + '^{tree}'], cwd=ROOT).decode().strip(),
                         merge['tree_sha'])
        policy = json.loads(subprocess.check_output(['git', 'show', head + ':.github/excel-evidence-policy.json'], cwd=ROOT))
        self.assertEqual(policy['expected_assertions'], 1947)
        self.assertEqual({s['path'] for s in record['sources']}, set(policy['regression_sources']))
        for source in record['sources']:
            blob = subprocess.check_output(['git', 'show', head + ':' + source['path']], cwd=ROOT)
            self.assertEqual('sha256:' + hashlib.sha256(blob).hexdigest(), source['sha256'])
            self.assertEqual((ROOT / source['path']).read_text(), blob.decode().replace('\r\n', '\n'))


if __name__ == '__main__':
    unittest.main()
