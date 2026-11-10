"""Extended Gamma frozen evidence, rational underflow control and port isolation."""
from fractions import Fraction as F
import base64
import hashlib
import json
import math
import random
import re
import subprocess
import unittest
from public_integration import restore,EXTRA

import extended_gamma as v1
from extended_gamma_v2 import public
from install_extended_gamma_fidelity import HERE,ROOT,BEGIN,END,dataset,render,strip
from install_wide_gamma_fidelity import strip as strip_wide


class ExtendedGammaTests(unittest.TestCase):
    def test_exact_source_excel_evidence_and_unchanged_public_failures(self):
        artifact=json.loads((HERE/'excel_extended_gamma_artifact.json').read_text())
        self.assertEqual(artifact['run_id'],42515382823)
        self.assertEqual(artifact['artifact_id'],13631475342)
        self.assertEqual(artifact['artifact_zip_sha256'],'9f6facd9a70831c3b2e4230664da513fd352156e0260f853c4fef86a22b9fc2d')
        decoded={}
        for name,entry in artifact['members'].items():
            raw=base64.b64decode(entry['base64'],validate=True)
            self.assertEqual(hashlib.sha256(raw).hexdigest(),entry['sha256'])
            decoded[name]=raw
        record=json.loads(decoded['excel-certification.json'].decode('utf-8-sig'))
        sha='924ff800077a3001bd1bcc246ad3a370f1c78fba'
        self.assertEqual(record['candidate_sha'],sha)
        self.assertEqual(record['runner']['workflow']['sha'],sha)
        self.assertEqual(record['runner']['workflow']['run_id'],42515382823)
        self.assertEqual(record['harness'],dict(entry_point='Test_STATS_PROBDIST_RunAll_CI',assertions=9962,passed=9921,failed=41))
        for stage in ('import','compile','cleanup'):
            self.assertEqual(record['stages'][stage]['status'],'PASS')
        self.assertEqual(record['stages']['regression']['status'],'FAIL')
        self.assertTrue(all(not grid['exported'] for grid in record['grids'].values()))
        self.assertEqual(record['log']['sha256'],'sha256:'+hashlib.sha256(decoded['test-result.txt']).hexdigest())
        policy=json.loads(subprocess.check_output(['git','show',sha+':.github/excel-evidence-policy.json'],cwd=ROOT))
        self.assertEqual(policy['expected_assertions'],9962)
        self.assertEqual({s['path'] for s in record['sources']},set(policy['regression_sources']))
        for source in record['sources']:
            blob=subprocess.check_output(['git','show',sha+':'+source['path']],cwd=ROOT)
            self.assertEqual('sha256:'+hashlib.sha256(blob).hexdigest(),source['sha256'])
            self.assertEqual(strip_wide((ROOT/source['path']).read_text()),blob.decode().replace('\r\n','\n'))
        old=json.loads((HERE/'excel_scaled_gamma_artifact.json').read_text())
        def failures(raw):
            return re.findall(r'^\[[^\]]+\]   (.*)$',raw.decode('utf-8-sig').replace('\r\n','\n'),re.M)
        previous=failures(base64.b64decode(old['members']['test-result.txt']['base64']))
        current=failures(decoded['test-result.txt'])
        self.assertEqual(len(current),41)
        self.assertEqual(current,previous)
        self.assertFalse(any(s.startswith('EG ') for s in current))

    def test_frozen_candidate_and_complete_model_screen(self):
        protocol=json.loads((HERE/'extended_protocol.json').read_text())
        result=json.loads((HERE/'extended_screen.json').read_text())
        self.assertEqual(result['protocol_commit'],'64430e088ac7ed64ef701049b8f7f46a996ebcd8')
        self.assertEqual(result['status'],'BOUNDED MODEL PASS; NOT PUBLIC ACCEPTANCE')
        self.assertEqual(result['failures'],[])
        self.assertEqual(result['monotonicity_failures'],[])
        self.assertEqual(result['protocol_sha256'],hashlib.sha256((HERE/'extended_protocol.json').read_bytes()).hexdigest())
        for path,sha in protocol['sha256_lf'].items():
            self.assertEqual(hashlib.sha256((ROOT/path).read_bytes().replace(b'\r\n',b'\n')).hexdigest(),sha)
        self.assertEqual(len(result['rows']),220)
        self.assertEqual([{k:r[k] for k in ('a_hex','x_hex')} for r in result['rows']],protocol['points'])
        for row in result['rows']:
            a,x=[float.fromhex(row[k]) for k in ('a_hex','x_hex')]
            self.assertNotEqual(a,2.)
            self.assertEqual([v.hex() for v in public(a,x)],row['values_hex'])

    def test_new_shapes_are_disjoint_from_declared_inventory(self):
        protocol=json.loads((HERE/'extended_protocol.json').read_text())
        old=set()
        def scan(obj):
            if isinstance(obj,dict):
                for ka,kx in [('a_hex','x_hex'),('a','x')]:
                    if ka in obj and kx in obj:
                        try:
                            a=obj[ka]
                            a=float.fromhex(a) if isinstance(a,str) and ('0x' in a or a in ('inf','nan')) else float(a)
                            old.add(a.hex())
                        except (TypeError,ValueError):
                            pass
                for value in obj.values():scan(value)
            elif isinstance(obj,list):
                for value in obj:scan(value)
        for path,digest in protocol['disjointness']['audited_files'].items():
            blob=subprocess.check_output(['git','show','64430e088ac7ed64ef701049b8f7f46a996ebcd8:'+path],cwd=ROOT)
            self.assertEqual(hashlib.sha256(blob).hexdigest(),digest)
            scan(json.loads(blob))
        fresh={p['a_hex'] for p in protocol['points']}
        self.assertEqual(len(fresh),20)
        self.assertFalse(fresh&old)
        self.assertEqual(len(old),protocol['disjointness']['excluded_shape_count'])

    def test_v1_counterexample_is_preserved_and_v2_repairs_it(self):
        x=2.**-537
        self.assertEqual(v1.public(2.,x)[0],2.**-1074)
        self.assertEqual(public(2.,x),(0.,1.))
        for x in (math.nextafter(x,0.),x,math.nextafter(x,math.inf)):
            exact=F(x)**2/2
            self.assertEqual(float(exact*(1-F(x))),float(exact))
            self.assertEqual(public(2.,x),(float(exact),1.))

    def test_shape_two_against_independent_rational_enclosures(self):
        rng=random.Random(260824)
        xs=[math.ldexp(rng.uniform(.5,1.),rng.randrange(-537,-511)) for _ in range(8192)]
        xs += [2.**-1074,2.**-1022,2.**-512]
        for n in range(1,64,2):
            x=n*2.**-537
            xs += [math.nextafter(x,0.),x,math.nextafter(x,math.inf)]
        self.assertEqual(len(xs),8291)
        for x in xs:
            exact=F(x)**2/2
            lower=exact*(1-F(x))
            self.assertEqual(float(lower),float(exact),x.hex())
            self.assertEqual(public(2.,x),(float(lower),1.),x.hex())

    def test_generated_source_and_aggregate_policy(self):
        saved=json.loads((HERE/'extended_gamma_fidelity.json').read_text())
        self.assertEqual(saved,dataset())
        self.assertEqual(len(saved['rows']),587)
        self.assertEqual(saved['additional_assertions'],1772)
        text=restore((ROOT/'tests/M_STATS_PROBDIST_TEST.bas').read_text())
        block=BEGIN+text.split(BEGIN)[1].split(END)[0]+END+'\n\n'
        self.assertEqual(block,render(saved))
        self.assertEqual(text.count('    Test_ExtendedGamma_Fidelity\n'),1)
        policy=json.loads((ROOT/'.github/excel-evidence-policy.json').read_text())
        wide=json.loads((HERE/'wide_gamma_fidelity.json').read_text())
        self.assertEqual(policy['expected_assertions'],9962+wide['additional_assertions']+EXTRA)

    def test_frozen_vba_and_public_sources_are_unchanged(self):
        sha='ec8a5f98edc8d37e3e9e50c1344b2fc31c33289f'
        for path in [ROOT/'tests/M_STATS_PROBDIST_TEST.bas',*(ROOT/'src').glob('*.bas')]:
            blob=subprocess.check_output(['git','show',sha+':'+path.relative_to(ROOT).as_posix()],cwd=ROOT).decode()
            self.assertEqual(strip(path.read_text()),blob.replace('\r\n','\n'))

    def test_invalid_inputs_remain_rejected(self):
        for a,x in dataset()['invalid']:
            with self.assertRaises(ValueError):public(float.fromhex(a),float.fromhex(x))


if __name__=='__main__':
    unittest.main()
