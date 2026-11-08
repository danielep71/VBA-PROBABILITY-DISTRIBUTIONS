"""Wide model identity, revealed replays, independent constants and private port."""
from fractions import Fraction as F
import hashlib
import json
import subprocess
import unittest
import mpmath as mp
import sympy as sy
from install_wide_gamma_fidelity import HERE,ROOT,BEGIN,END,BASELINE,dataset,render,strip
from wide_gamma import public,STIRLING,LOG_SQRT_2PI


def digest(path):return hashlib.sha256(path.read_bytes().replace(b'\r\n',b'\n')).hexdigest()


class WideGammaTests(unittest.TestCase):
    def test_frozen_model_and_complete_screen(self):
        protocol=json.loads((HERE/'wide_protocol.json').read_text())
        result=json.loads((HERE/'wide_complete_v2_screen.json').read_text())
        self.assertEqual(result['protocol_commit'],'69b849ac7873c40db1ca8f2962a011b8996d5557')
        self.assertEqual(result['status'],'BOUNDED MODEL PASS; NOT PUBLIC ACCEPTANCE')
        self.assertEqual(result['failures'],[]);self.assertEqual(result['monotonicity_failures'],[])
        self.assertEqual(result['protocol_sha256'],digest(HERE/'wide_protocol.json'))
        self.assertEqual(result['screen_sha256'],digest(HERE/'screen_wide_complete_v2.py'))
        self.assertEqual(result['initial_attempt_sha256'],digest(HERE/'wide_initial_attempt.txt'))
        self.assertEqual(result['incomplete_completion_sha256'],digest(HERE/'wide_complete_screen.json'))
        for path,sha in protocol['sha256_lf'].items():self.assertEqual(digest(ROOT/path),sha)
        self.assertEqual([{k:r[k] for k in ('a_hex','x_hex')} for r in result['rows']],protocol['points'])
        self.assertEqual(len(result['rows']),198)
        for row in result['rows']:
            self.assertEqual([v.hex() for v in public(*[float.fromhex(row[k]) for k in ('a_hex','x_hex')])],row['values_hex'])
            self.assertEqual([m['dps'] for m in row['oracle_methods']],[160,320])
        # Retain all originally encountered oracle failures, even though completed.
        initial=json.loads((HERE/'wide_complete_screen.json').read_text())
        self.assertEqual(initial['status'],'REJECT')
        self.assertEqual([f['index'] for f in initial['failures']],[158,171,184,197])
        for filename,sha in [('screen_wide_complete.py','830eaa728c7b9195e92b263709479e166c19506d'),
                             ('screen_wide_complete_v2.py','9cbb53b02c15a0dc2ab4b59d9ab6c59acec30631')]:
            blob=subprocess.check_output(['git','show',sha+':'+(HERE/filename).relative_to(ROOT).as_posix()],cwd=ROOT)
            self.assertEqual((HERE/filename).read_bytes(),blob)

    def test_declared_shapes_are_disjoint(self):
        p=json.loads((HERE/'wide_protocol.json').read_text());old=set()
        def scan(obj):
            if isinstance(obj,dict):
                for ka,kx in [('a_hex','x_hex'),('a','x')]:
                    if ka in obj and kx in obj:
                        try:
                            a=obj[ka]
                            a=float.fromhex(a) if isinstance(a,str) and ('0x' in a or a in ('inf','nan')) else float(a)
                            old.add(a.hex())
                        except (TypeError,ValueError):pass
                for v in obj.values():scan(v)
            elif isinstance(obj,list):
                for v in obj:scan(v)
        for path,sha in p['disjointness']['audited_files'].items():
            blob=subprocess.check_output(['git','show','69b849ac7873c40db1ca8f2962a011b8996d5557:'+path],cwd=ROOT)
            self.assertEqual(hashlib.sha256(blob).hexdigest(),sha);scan(json.loads(blob))
        new={r['a_hex'] for r in p['points']}
        self.assertEqual(len(new),14);self.assertEqual(len(old),81);self.assertFalse(old&new)

    def test_revealed_outer_boundaries_repaired(self):
        old=json.loads((HERE/'routing_extended/result.json').read_text())
        self.assertEqual(old['status'],'REJECT');self.assertEqual(len(old['failures']),13)
        self.assertEqual(len(old['new_monotonicity_failures']),1)
        self.assertEqual(len(old['retained_monotonicity_failures']),2)
        result=json.loads((HERE/'wide_outer_replay.json').read_text())
        self.assertEqual(result['status'],'PASS');self.assertEqual(result['failures'],[])
        self.assertEqual(result['monotonicity_failures'],[])
        self.assertEqual(sum(len(g['rows']) for g in result['groups']),93)
        for oldgroup,newgroup in zip(old['groups'],result['groups']):
            self.assertEqual(oldgroup['name'],newgroup['name'])
            for a,b in zip(oldgroup['rows'],newgroup['rows']):
                for key in ('a_hex','x_hex','reference','reference_hex'):self.assertEqual(a[key],b[key])
                self.assertEqual([v.hex() for v in public(*[float.fromhex(b[k]) for k in ('a_hex','x_hex')])],b['values_hex'])
            for l,r in zip(newgroup['rows'],newgroup['rows'][1:]):
                for j in (0,1):
                    lv,rv=[float.fromhex(t['values_hex'][j]) for t in (l,r)]
                    if (newgroup['axis']=='x')==(j==0):self.assertLessEqual(lv,rv)
                    else:self.assertGreaterEqual(lv,rv)

    def test_exact_constants_and_analytic_truncation(self):
        for k,(num,den) in enumerate(STIRLING,1):
            self.assertEqual(sy.Rational(num,den),sy.bernoulli(2*k)/(2*k*(2*k-1)))
        omitted=F(854513,63756)/128**21
        self.assertLess(omitted,F(1,2**142))
        for dps in (110,160):
            with mp.workdps(dps):
                c=mp.log(2*mp.pi)/2;hi=float(c);lo=float(c-mp.mpf(hi))
                # Stored scaled normalization must reproduce the exact two-Double sum.
                h,l,e=LOG_SQRT_2PI.hi,LOG_SQRT_2PI.lo,LOG_SQRT_2PI.exponent
                self.assertEqual((F(h)+F(l))*F(2)**e,F(hi)+F(lo))
                self.assertLess(-1280+128+128*mp.log(10),-1075*mp.log(2))

    def test_previous_interior_bits_and_invalids(self):
        rows=json.loads((HERE/'extended_gamma_fidelity.json').read_text())['rows']
        for row in rows:
            self.assertEqual([v.hex() for v in public(float.fromhex(row['a']),float.fromhex(row['x']))],[row['p'],row['q']])
        for a,x in dataset()['invalid']:
            with self.assertRaises(ValueError):public(float.fromhex(a),float.fromhex(x))
        for a in (1.,129.,1e14):self.assertEqual(public(a,0),(0.,1.))

    def test_generated_port_and_historical_sources(self):
        saved=json.loads((HERE/'wide_gamma_fidelity.json').read_text())
        self.assertEqual(saved,dataset());self.assertEqual(len(saved['rows']),290)
        self.assertEqual(saved['additional_assertions'],886)
        text=(ROOT/'tests/M_STATS_PROBDIST_TEST.bas').read_text()
        self.assertEqual(text.count(BEGIN),1);self.assertEqual(text.count(END),1)
        self.assertEqual(text.count('    Test_WideGamma_Fidelity\n'),1)
        block=BEGIN+text.split(BEGIN)[1].split(END)[0]+END+'\n\n'
        self.assertEqual(block,render(saved))
        self.assertEqual(json.loads((ROOT/'.github/excel-evidence-policy.json').read_text())['expected_assertions'],BASELINE+886)
        sha='924ff800077a3001bd1bcc246ad3a370f1c78fba'
        for path in [ROOT/'tests/M_STATS_PROBDIST_TEST.bas',*(ROOT/'src').glob('*.bas')]:
            original=subprocess.check_output(['git','show',sha+':'+path.relative_to(ROOT).as_posix()],cwd=ROOT).decode().replace('\r\n','\n')
            self.assertEqual(strip(path.read_text()),original)


if __name__=='__main__':unittest.main()
