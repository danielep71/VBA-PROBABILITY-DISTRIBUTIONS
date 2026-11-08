"""Fresh fixed-candidate references, exact conversion and retained public evidence."""
import base64
import csv
import io
import zipfile
from fractions import Fraction as F
import hashlib
import json
import math
import random
import subprocess
import sys
import unittest
import mpmath as mp
from install_scaled_fidelity import HERE,ROOT
from reference_binary64 import rounded
from wide_gamma import public


class PublicFreshTests(unittest.TestCase):
    def test_integer_rounding_against_exact_rationals(self):
        rng=random.Random(2421108)
        with mp.workdps(160):
            for _ in range(4096):
                m=rng.randrange(2**105,2**106)*rng.choice([-1,1]);e=rng.randrange(-1200,900)
                exact=F(m)*F(2)**e;value=mp.mpf(m)*mp.mpf(2)**e
                try:expected=float(exact)
                except OverflowError:expected=math.copysign(math.inf,m)
                self.assertEqual(rounded(value).hex(),expected.hex())
            for n in (0,1,2,3,2**51-1,2**51,2**52-1,2**52,2**52+1):
                for delta in (-F(1,2**100),F(0),F(1,2**100)):
                    for sign in (-1,1):
                        exact=sign*(F(n)+F(1,2)+delta)*F(2)**-1074
                        value=mp.mpf(exact.numerator)/exact.denominator
                        self.assertEqual(rounded(value).hex(),float(exact).hex())
            self.assertEqual(rounded(mp.mpf(2)**-1000000000),0.)
            self.assertEqual(rounded(-mp.mpf(2)**-1000000000).hex(),(-0.).hex())

    def test_reported_subnormal_oracle_conversion(self):
        for dps in (160,320):
            with mp.workdps(dps):
                a,x=mp.mpf('95.0625'),mp.mpf(1024)
                ref=mp.gammainc(a,x,mp.inf,regularized=True)
                sign,m,e,bits=ref._mpf_
                exact=F((-1)**sign*m)*F(2)**e
                self.assertEqual(rounded(ref).hex(),'0x0.f0c9cf0026f39p-1022')
                self.assertEqual(rounded(ref),float(exact))
                self.assertEqual(float(ref).hex(),'0x0.f0c9cf0026f3ap-1022')
                rate=1-(a-1)/x
                factor=mp.exp((a-1)*mp.log(x)-x-mp.loggamma(a))/rate
                def density(u):return mp.mpf(0) if mp.isinf(u) else mp.exp((a-1)*mp.log1p(u/(rate*x))-u/rate)
                quad=factor*mp.quad(density,[0,1,4,16,64,mp.inf])
                self.assertLess(abs(quad/ref-1),mp.mpf(10)**(-dps+15))
                self.assertEqual(rounded(quad),rounded(ref))

    def test_frozen_completion_and_retained_initial_result(self):
        p=json.loads((HERE/'public_fresh_protocol.json').read_text())
        r=json.loads((HERE/'public_fresh_complete_screen.json').read_text())
        self.assertEqual(r['status'],'FORWARD MODEL PASS; INVERSE REFERENCES READY; VBA PUBLIC ACCEPTANCE PENDING')
        self.assertEqual(r['failures'],[]);self.assertEqual(r['monotonicity_failures'],[])
        for path,sha in p['sha256_lf'].items():self.assertEqual(hashlib.sha256((ROOT/path).read_bytes().replace(b'\r\n',b'\n')).hexdigest(),sha)
        for field,name in [('protocol_sha256','public_fresh_protocol.json'),('original_screen_sha256','public_fresh_screen.json'),
                           ('conversion_sha256','reference_binary64.py'),('completion_sha256','screen_public_fresh_complete.py')]:
            self.assertEqual(r[field],hashlib.sha256((HERE/name).read_bytes()).hexdigest())
        for name in ('reference_binary64.py','screen_public_fresh_complete.py'):
            frozen=subprocess.check_output(['git','show','c8b145ebf70a94652b97e8a49a9edf68a0f50568:'+str((HERE/name).relative_to(ROOT))],cwd=ROOT)
            self.assertEqual(frozen,(HERE/name).read_bytes())
        initial=json.loads((HERE/'public_fresh_screen.json').read_text())
        self.assertEqual(initial['status'],'REJECT')
        self.assertEqual(initial['failures'],[dict(index=99,tail=1,reason='accuracy')])
        self.assertEqual(initial['monotonicity_failures'],[])
        self.assertEqual(r['rows'][99]['values_hex'],initial['rows'][99]['values_hex'])
        self.assertEqual(r['rows'][99]['reference'],initial['rows'][99]['reference'])
        self.assertNotEqual(r['rows'][99]['reference_hex'],initial['rows'][99]['reference_hex'])

    def test_fresh_cohort_and_generated_host_data(self):
        p=json.loads((HERE/'public_fresh_protocol.json').read_text());r=json.loads((HERE/'public_fresh_complete_screen.json').read_text())
        self.assertEqual(len(p['points']),206);self.assertEqual(len(p['inverses']),24)
        self.assertEqual([{k:row[k] for k in ('a_hex','x_hex')} for row in r['rows']],p['points'])
        self.assertEqual([{k:row[k] for k in ('a_hex','p_hex','q_hex')} for row in r['inverses']],p['inverses'])
        seen={}
        for row in r['rows']:
            a,x=[float.fromhex(row[k]) for k in ('a_hex','x_hex')]
            values=public(a,x);self.assertEqual([v.hex() for v in values],row['values_hex'])
            with mp.workdps(320):
                for j,v in enumerate(values):
                    ref=mp.mpf(row['reference'][j]);self.assertEqual(rounded(ref).hex(),row['reference_hex'][j])
                    if rounded(ref)<sys.float_info.min:self.assertEqual(v,rounded(ref))
                    else:self.assertLessEqual(abs(mp.mpf(v)/ref-1),mp.mpf(('3e-15','5e-15')[j]))
            if a in seen:
                self.assertLessEqual(seen[a][0],values[0]);self.assertGreaterEqual(seen[a][1],values[1])
            seen[a]=values
        self.assertEqual(len(seen),14)
        data=json.loads((HERE/'public_fresh_fidelity.json').read_text())
        self.assertEqual(data['rows'],[dict(a=x['a_hex'],x=x['x_hex'],values_hex=x['values_hex']) for x in r['rows']])
        self.assertEqual(data['inverses'],[dict(a=x['a_hex'],p=x['p_hex'],reference_hex=x['reference_hex']) for x in r['inverses']])
        self.assertEqual(data['additional_assertions'],1696)

    def test_shape_disjointness_from_declared_inventory(self):
        p=json.loads((HERE/'public_fresh_protocol.json').read_text());old=set()
        def scan(obj):
            if isinstance(obj,dict):
                for ka,kx in [('a_hex','x_hex'),('a','x')]:
                    if ka in obj and kx in obj:
                        try:
                            a=obj[ka];a=float.fromhex(a) if isinstance(a,str) and ('0x' in a or a in ('inf','nan')) else float(a);old.add(a.hex())
                        except (TypeError,ValueError):pass
                for v in obj.values():scan(v)
            elif isinstance(obj,list):
                for v in obj:scan(v)
        for path,sha in p['disjointness']['audited_files'].items():
            blob=subprocess.check_output(['git','show','5c23f636337a315a5e13170cb53aacafef0ae707:'+path],cwd=ROOT)
            self.assertEqual(hashlib.sha256(blob).hexdigest(),sha);scan(json.loads(blob))
        self.assertEqual(len(old),99)
        self.assertFalse(old&{row['a_hex'] for row in p['points']})

    def test_fresh_public_excel_run(self):
        a=json.loads((HERE/'excel_public_fresh_artifact.json').read_text());decoded={}
        self.assertEqual(a['run_id'],42524515131)
        self.assertEqual(a['artifact_id'],13635460900)
        self.assertEqual(a['artifact_zip_sha256'],'c77ad4c66cf7f03c1cd4110c359615660b807f526a10efc86fa78c62a19ebe9b')
        for n,v in a['members'].items():
            b=base64.b64decode(v['base64'],validate=True);self.assertEqual(hashlib.sha256(b).hexdigest(),v['sha256']);decoded[n]=b
        record=json.loads(decoded['excel-certification.json'].decode('utf-8-sig'))
        self.assertEqual(record['harness'],dict(entry_point='Test_STATS_PROBDIST_RunAll_CI',assertions=14878,passed=14878,failed=0))
        self.assertTrue(all(s['status']=='PASS' for s in record['stages'].values()))
        self.assertTrue(all(not g['exported'] for g in record['grids'].values()))
        self.assertEqual(record['log']['sha256'],'sha256:'+hashlib.sha256(decoded['test-result.txt']).hexdigest())
        merge=a['tested_merge'];raw=base64.b64decode(merge['commit_object_base64'],validate=True)
        self.assertEqual(hashlib.sha1(b'commit '+str(len(raw)).encode()+b'\0'+raw).hexdigest(),merge['sha'])
        self.assertEqual(record['candidate_sha'],merge['sha']);self.assertEqual(record['runner']['workflow']['sha'],merge['sha'])
        self.assertEqual(merge['sha'],'a0202380a6b6b07e102585db455ebb769b6fa0b0')
        self.assertEqual(record['runner']['workflow']['run_id'],a['run_id'])
        head='b005341702f5fb70323453bded9b2109250ee541';self.assertEqual(merge['head_sha'],head)
        self.assertEqual(raw.decode().splitlines()[0],'tree '+merge['tree_sha'])
        self.assertEqual(subprocess.check_output(['git','rev-parse',head+'^{tree}'],cwd=ROOT).decode().strip(),merge['tree_sha'])
        policy=json.loads(subprocess.check_output(['git','show',head+':.github/excel-evidence-policy.json'],cwd=ROOT))
        self.assertEqual(policy['expected_assertions'],14878)
        self.assertEqual({s['path'] for s in record['sources']},set(policy['regression_sources']))
        for s in record['sources']:
            blob=subprocess.check_output(['git','show',head+':'+s['path']],cwd=ROOT)
            self.assertEqual('sha256:'+hashlib.sha256(blob).hexdigest(),s['sha256'])
            self.assertEqual((ROOT/s['path']).read_text(),blob.decode().replace('\r\n','\n'))

    def test_first_public_excel_run_resolves_all_old_failures(self):
        a=json.loads((HERE/'excel_public_gamma_artifact.json').read_text());decoded={}
        self.assertEqual(a['run_id'],42523118255);self.assertEqual(a['artifact_id'],13634552161)
        self.assertEqual(a['artifact_zip_sha256'],'b8e8f2eb2261f209fa649ea960decca104f53deea0691dcf8b6dd755694a9487')
        for n,v in a['members'].items():
            b=base64.b64decode(v['base64'],validate=True);self.assertEqual(hashlib.sha256(b).hexdigest(),v['sha256']);decoded[n]=b
        record=json.loads(decoded['excel-certification.json'].decode('utf-8-sig'))
        self.assertEqual(record['harness'],dict(entry_point='Test_STATS_PROBDIST_RunAll_CI',assertions=13182,passed=13182,failed=0))
        self.assertTrue(all(s['status']=='PASS' for s in record['stages'].values()))
        self.assertTrue(all(not g['exported'] for g in record['grids'].values()))
        self.assertEqual(record['log']['sha256'],'sha256:'+hashlib.sha256(decoded['test-result.txt']).hexdigest())
        merge=a['tested_merge'];raw=base64.b64decode(merge['commit_object_base64'],validate=True)
        self.assertEqual(hashlib.sha1(b'commit '+str(len(raw)).encode()+b'\0'+raw).hexdigest(),merge['sha'])
        self.assertEqual(merge['sha'],'9278ce3b2ffaf6786bba42a5876066f5539d9828')
        self.assertEqual(record['candidate_sha'],merge['sha']);self.assertEqual(record['runner']['workflow']['sha'],merge['sha'])
        head='1440b180036091c7c7635eb87e155a2762bb2a30';self.assertEqual(merge['head_sha'],head)
        self.assertEqual(raw.decode().splitlines()[0],'tree '+merge['tree_sha'])
        self.assertEqual(subprocess.check_output(['git','rev-parse',head+'^{tree}'],cwd=ROOT).decode().strip(),merge['tree_sha'])
        policy=json.loads(subprocess.check_output(['git','show',head+':.github/excel-evidence-policy.json'],cwd=ROOT))
        self.assertEqual(policy['expected_assertions'],13182)
        self.assertEqual({s['path'] for s in record['sources']},set(policy['regression_sources']))
        for s in record['sources']:
            blob=subprocess.check_output(['git','show',head+':'+s['path']],cwd=ROOT)
            self.assertEqual('sha256:'+hashlib.sha256(blob).hexdigest(),s['sha256'])
            if s['path'].startswith('src/'):
                self.assertEqual((ROOT/s['path']).read_text(),blob.decode().replace('\r\n','\n'))

    def test_fresh_export_bytes_and_declared_caller_movement(self):
        folder = HERE
        artifact = json.loads((folder / 'excel_public_export_artifact.json').read_text())
        raw = base64.b64decode(artifact['artifact_zip_base64'], validate=True)
        self.assertEqual(hashlib.sha256(raw).hexdigest(), artifact['artifact_zip_sha256'])
        self.assertEqual(artifact['artifact_zip_sha256'],
                         '582cd4efbca45e6e8fff287d815a9370aca39e93b35fbb9024a5e4c71386d15b')
        self.assertEqual(artifact['run_id'], 42529642623)
        self.assertEqual(artifact['artifact_id'], 13637073116)
        head = 'e91c535a8f7ef2ba3830d9c690b39966b74cc6b3'
        self.assertEqual(artifact['candidate_sha'], head)
        movement = json.loads((folder / 'public_export_movement.json').read_text())
        self.assertEqual(movement['candidate_sha'], head)
        self.assertEqual(movement['run_id'], artifact['run_id'])
        with zipfile.ZipFile(io.BytesIO(raw)) as archive:
            manifest = json.loads(archive.read('export-manifest.json').decode('utf-8-sig'))
            self.assertEqual(manifest['candidate_sha'], head)
            self.assertEqual(manifest['status'], 'PASS')
            self.assertEqual(manifest['requested'], ['main', 'holdout'])
            self.assertEqual(manifest['main_mode'], 'full')
            self.assertEqual(manifest['regression']['assertions'], 14878)
            self.assertEqual(manifest['regression']['failed'], 0)
            self.assertEqual(manifest['workflow']['run_id'], artifact['run_id'])
            record_bytes = archive.read('excel-certification.json')
            self.assertEqual('sha256:' + hashlib.sha256(record_bytes).hexdigest(),
                             manifest['regression']['record_sha256'])
            record = json.loads(record_bytes.decode('utf-8-sig'))
            self.assertEqual(record['candidate_sha'], head)
            self.assertEqual(record['harness']['assertions'], 14878)
            self.assertEqual(record['harness']['passed'], 14878)
            self.assertEqual(record['harness']['failed'], 0)
            self.assertTrue(all(stage['status'] == 'PASS' for stage in record['stages'].values()))
            self.assertEqual(record['log']['sha256'],
                             'sha256:' + hashlib.sha256(archive.read('test-result.txt')).hexdigest())
            for source in record['sources']:
                blob = subprocess.check_output(['git', 'show', head + ':' + source['path']], cwd=ROOT)
                self.assertEqual('sha256:' + hashlib.sha256(blob).hexdigest(), source['sha256'])
                self.assertEqual((ROOT / source['path']).read_text(), blob.decode().replace('\r\n', '\n'))
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


if __name__=='__main__':unittest.main()
