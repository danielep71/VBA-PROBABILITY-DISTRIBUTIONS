"""Fail closed on incomplete/mismatched B4 host evidence; report warm pair costs."""
import json
import math
import statistics
import sys
from pathlib import Path


def validate(report, design):
    assert report['status']=='PASS' and not report['error'] and not report['cleanup_error']
    assert len(report['candidate_sha'])==40
    int(report['candidate_sha'],16)
    assert report['normalization_check']=='PASS checks=11'
    expected = {r['id']:r['expected'] for r in design['cases']}
    assert len(expected)==len(design['cases'])==len(report['rows'])
    seen=set()
    for row in report['rows']:
        assert row['id'] not in seen and row['id'] in expected
        seen.add(row['id'])
        parts=row['raw'].split(',')
        declined=parts[0]=='DECLINED'
        assert len(parts)==(3 if declined else 5)
        assert ':'.join(parts[-2:])==row['id']
        observed='DECLINED' if declined else ','.join(parts[:2])
        assert observed==expected[row['id']]
    points=((.0001,.5),(.5,.25),(2.5,1.),(2.5,6.),(100.,90.),(100.,110.),(10000.,9900.),(10000.,10100.))
    matrix={(a,x,n,b) for a,x in points for n in (1,2,3) for b in (True,False)}
    observed={(r['a'],r['x'],r['round'],r['baseline']) for r in report['timings']}
    assert len(report['timings'])==len(matrix) and observed==matrix
    import struct
    for r in report['timings']:
        assert r['calls']==4 and math.isfinite(r['elapsed_us']) and r['elapsed_us']>0
        if not r['baseline']:
            key=struct.pack('>d',r['a']).hex().upper()+':'+struct.pack('>d',r['x']).hex().upper()
            assert r['p']+','+r['q']==expected[key]
        assert len(r['p'])==len(r['q'])==16
        int(r['p'],16);int(r['q'],16)
    for a,x in points:
        for b in (True,False):
            assert len({(r["p"],r["q"]) for r in report["timings"] if (r["a"],r["x"],r["baseline"])==(a,x,b)})==1
    return points


if __name__=='__main__':
    path=Path(sys.argv[1])
    report=json.loads((path/'pair-results.json').read_text(encoding='utf-8-sig'))
    design=json.loads((path/'inputs.json').read_text())
    points=validate(report,design)
    print('PASS: complete host fidelity and timing matrix')
    print('| A | X | Baseline pair us | B4 pair us | Speedup |')
    print('|---:|---:|---:|---:|---:|')
    for a,x in points:
        med={b:statistics.median(r['elapsed_us']/r['calls'] for r in report['timings'] if r['a']==a and r['x']==x and r['baseline']==b) for b in (True,False)}
        print(f'| {a:g} | {x:g} | {med[True]:.3f} | {med[False]:.3f} | {med[True]/med[False]:.3f}x |')
    print('Warm paired P/Q kernel cost only; not public-call, cold, inverse or release acceptance.')
