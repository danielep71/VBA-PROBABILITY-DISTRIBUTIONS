"""Bind the focused public port to the arithmetic and regressions tested in #242."""
import base64
import csv
import hashlib
import io
import json
from pathlib import Path
import re
import subprocess
import unittest
import zipfile

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
    def test_fresh_export_bytes_and_declared_caller_movement(self):
        folder = ROOT / 'benchmark/gamma_public_integration'
        artifact = json.loads((folder / 'excel_export_artifact.json').read_text())
        raw = base64.b64decode(artifact['artifact_zip_base64'], validate=True)
        self.assertEqual(hashlib.sha256(raw).hexdigest(), artifact['artifact_zip_sha256'])
        self.assertEqual(artifact['artifact_zip_sha256'],
                         'fefda02d53adcd4484f14db19d20e3e5796c114eb8dc20b822690b538071043c')
        self.assertEqual(artifact['run_id'], 42527060472)
        self.assertEqual(artifact['artifact_id'], 13636002652)
        head = '35e7381967d86c2f5597de456c668e640b653b55'
        self.assertEqual(artifact['candidate_sha'], head)
        movement = json.loads((folder / 'export_movement.json').read_text())
        self.assertEqual(movement['candidate_sha'], head)
        self.assertEqual(movement['run_id'], artifact['run_id'])
        with zipfile.ZipFile(io.BytesIO(raw)) as archive:
            manifest = json.loads(archive.read('export-manifest.json').decode('utf-8-sig'))
            self.assertEqual(manifest['candidate_sha'], head)
            self.assertEqual(manifest['status'], 'PASS')
            self.assertEqual(manifest['requested'], ['main', 'holdout'])
            self.assertEqual(manifest['main_mode'], 'full')
            self.assertEqual(manifest['regression']['assertions'], 1947)
            self.assertEqual(manifest['regression']['failed'], 0)
            self.assertEqual(manifest['workflow']['run_id'], artifact['run_id'])
            for grid in manifest['grids']:
                data = archive.read(grid['artifact_path'])
                self.assertEqual('sha256:' + hashlib.sha256(data).hexdigest(), grid['raw_sha256'])
                self.assertEqual((ROOT / grid['path']).read_bytes(), data.replace(b'\r\n', b'\n'))
                baseline = subprocess.check_output(['git', 'show', head + ':' + grid['path']], cwd=ROOT)
                old_rows = list(csv.DictReader(io.StringIO(baseline.decode())))
                new_rows = list(csv.DictReader(io.StringIO(data.decode())))
                self.assertEqual(len(old_rows), len(new_rows))
                result = movement['grids'][grid['name']]
                expected = {row['line']: row for row in result['changed_observations']}
                seen = {}
                for line, (old, new) in enumerate(zip(old_rows, new_rows), start=2):
                    self.assertEqual({k: v for k, v in old.items() if k != 'observed_vba'},
                                     {k: v for k, v in new.items() if k != 'observed_vba'})
                    if old['observed_vba'] == new['observed_vba']:
                        self.assertNotIn(line, expected)
                        continue
                    row = expected[line]
                    self.assertEqual(row['before'], old['observed_vba'])
                    self.assertEqual(row['after'], new['observed_vba'])
                    self.assertEqual(row['reference'], new['reference'])
                    self.assertEqual(row['function'], new['function'])
                    self.assertEqual(row['arguments'], {k: new[k] for k in ('arg1', 'arg2', 'arg3', 'arg4')})
                    self.assertTrue(new['function'].startswith(('Gamma_', 'ChiSquare_', 'Poisson_')))
                    seen[new['function']] = seen.get(new['function'], 0) + 1
                self.assertEqual(seen, result['changed_by_function'])
                self.assertEqual(sum(seen.values()), grid['changed_observations'])
                self.assertEqual(sum(seen.values()), result['changed_rows'])
                self.assertEqual(result['unchanged_rows'] + result['changed_rows'], len(new_rows))

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
