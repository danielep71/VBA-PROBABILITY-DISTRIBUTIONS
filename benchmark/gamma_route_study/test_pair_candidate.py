"""B4 evidence replay and fail-closed host validator negative controls."""
import copy
import json
import unittest
from pathlib import Path

import pair_candidate as candidate
import screen_pair_candidate as screen
from prepare_pair_host import hex_bits, inputs
from validate_pair_host import validate

HERE=Path(__file__).resolve().parent


class PairEvidence(unittest.TestCase):
    def test_development_replay(self):
        retained=json.loads((HERE/'pair_candidate_results.json').read_text())
        self.assertEqual(screen.evaluate(), retained)
        self.assertEqual(retained['status'],'DEVELOPMENT PASS ONLY')
        self.assertEqual(retained['retained_failures'],0)

    def test_domain_declines(self):
        for a,x in ((0.,1.),(10001.,10001.),(1.,0.),(1.,20001.),(float('nan'),1.)):
            with self.assertRaises(candidate.Declined):
                candidate.public(a,x)

    def test_incomplete_corrupt_and_duplicate_host_reports(self):
        design={'cases':inputs()}
        rows=[]
        for case in design['cases']:
            echo=case['id'].replace(':',',')
            raw=('DECLINED,'+echo if case['expected']=='DECLINED' else case['expected']+',0,'+echo)
            rows.append(dict(id=case['id'],raw=raw))
        timings=[]
        for a,x in ((.0001,.5),(.5,.25),(2.5,1.),(2.5,6.),(100.,90.),(100.,110.),(10000.,9900.),(10000.,10100.)):
            p,q=candidate.public(a,x)
            for n in (1,2,3):
                for b in (True,False):
                    timings.append(dict(a=a,x=x,round=n,baseline=b,calls=4,p=hex_bits(p),q=hex_bits(q),elapsed_us=1.))
        good=dict(normalization_check='PASS checks=11',status='PASS',candidate_sha='0'*40,error='',cleanup_error='',rows=rows,timings=timings)
        validate(good,design)  # synthetic validator control, never measurement evidence
        variants=[]
        bad=copy.deepcopy(good);bad["normalization_check"]="";variants.append(bad)
        bad=copy.deepcopy(good);bad['rows'].pop();variants.append(bad)
        bad=copy.deepcopy(good);bad['rows'][1]=bad['rows'][0];variants.append(bad)
        bad=copy.deepcopy(good);bad['rows'][0]['raw']='0000000000000000,0000000000000000,0,0000000000000000,0000000000000000';variants.append(bad)
        bad=copy.deepcopy(good);bad['cleanup_error']='COM failure';variants.append(bad)
        bad=copy.deepcopy(good);bad['timings'].pop();variants.append(bad)
        bad=copy.deepcopy(good);bad['timings'][0]['elapsed_us']=0.;variants.append(bad)
        for bad in variants:
            with self.assertRaises(AssertionError):
                validate(bad,design)


if __name__=='__main__':
    unittest.main()
