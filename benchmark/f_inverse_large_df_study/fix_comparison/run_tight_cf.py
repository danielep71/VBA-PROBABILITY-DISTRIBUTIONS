"""Frozen P2 screen; the eight new rows are evaluated only under this protocol."""
import hashlib
import itertools
import json
from pathlib import Path
import time
import mpmath as mp
import pair_cf_screen as p1
import run_baseline as baseline
import vba_order_model as model

HERE=Path(__file__).resolve().parent
NEW_ROWS=tuple(itertools.product((.2,.8),((2e4,6.),(6.,2e4),(40.,2e7),(2e7,40.))))


def evaluate(p,d1,d2):
    old=model.forward.C['PROB_NUM_EPS']
    try:
        model.forward.C['PROB_NUM_EPS']=2.0**-100
        return p1.evaluate(p,d1,d2)
    finally:
        model.forward.C['PROB_NUM_EPS']=old


def main():
    retained=json.loads((HERE/'baseline_result.json').read_text())
    rows=[]
    points=[('revealed',r['kind'],*map(float.fromhex,r['inputs']),r['error']) for r in retained['rows']]
    points += [('new','screen',p,*dfs,None) for p,dfs in NEW_ROWS]
    for population,kind,p,d1,d2,old_error in points:
        row=dict(population=population,kind=kind,inputs=[v.hex() for v in (p,d1,d2)],complement=(1-p).hex())
        started=time.monotonic()
        try:
            r50=baseline.reference(p,d1,d2,50)
            r80=baseline.reference(p,d1,d2,80)
            with mp.workdps(80):
                stable=all(abs(a-b)<=mp.mpf('1e-40')*abs(b) for a,b in zip(r50,r80))
            values={}
            for name,fn in [('F-0',model.f_inverse),('P1',p1.evaluate),('P2',evaluate)]:
                value,x,y,trace=fn(p,d1,d2)
                with mp.workdps(80):
                    error=float(abs(mp.mpf(value)-r80[0])/r80[0])
                values[name]=dict(error=error,result=value.hex(),pair=[x.hex(),y.hex()],
                                  iterations=len(trace),trace=trace)
            passed=stable and values['P2']['error']<2**-48
            if kind=='control':
                passed=passed and values['P2']['error']<=4*old_error
            row.update(passes=passed,oracle_stable=stable,variants=values)
        except (ValueError,ArithmeticError) as exc:
            row.update(passes=False,failure=str(exc))
        row['seconds']=time.monotonic()-started
        rows.append(row)
        print(population,p,d1,d2,row.get('variants',{}).get('P2',{}).get('error',row.get('failure')),row['passes'],flush=True)
    paths=[Path(__file__),Path(p1.__file__),Path(p1.DD_MODULE.__file__),Path(model.__file__),
           Path(model.forward.__file__),Path(model.normal.__file__),HERE/'TIGHT_CF_PROTOCOL.md',
           HERE/'baseline_result.json',model.ROOT/'benchmark/_ibeta.py']
    result=dict(status='SCREEN_PASS_NOT_ADOPTION' if all(r['passes'] for r in rows) else 'REJECT',
                holdout_evaluated=False,rows=rows,
                bindings={str(p.relative_to(model.ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in paths})
    (HERE/'tight_cf_result.json').write_text(json.dumps(result,indent=2)+'\n')
    print(result['status'])


if __name__=='__main__':
    main()
