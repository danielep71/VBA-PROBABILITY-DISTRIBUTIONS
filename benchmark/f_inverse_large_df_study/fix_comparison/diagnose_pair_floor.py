"""Exploratory floor decomposition of rejected P1; revealed retained rows only."""
import hashlib
import json
from pathlib import Path
import mpmath as mp
import pair_cf_screen as candidate
import run_baseline as baseline
import vba_order_model as model

HERE=Path(__file__).resolve().parent


def experiment(p,d1,d2,variant):
    old_pdf=model.forward.beta_log_pdf
    old_eps=model.forward.C['PROB_NUM_EPS']
    try:
        if variant=='exact_prefactor':
            def pdf(x,y,a,b,lx,ly):
                # The existing caller adds log(x)+log(y); retain its operation order.
                with mp.workdps(80):
                    xx,aa,bb=map(mp.mpf,(x,a,b))
                    return float((aa-1)*mp.log(xx)+(bb-1)*mp.log(1-xx)-mp.log(mp.beta(aa,bb)))
            model.forward.beta_log_pdf=pdf
        elif variant=='tight_cf':
            model.forward.C['PROB_NUM_EPS']=2.0**-100
        return candidate.evaluate(p,d1,d2)
    finally:
        model.forward.beta_log_pdf=old_pdf
        model.forward.C['PROB_NUM_EPS']=old_eps


def measures(value,x,y,ref,d1,d2):
    with mp.workdps(80):
        pair_f=mp.mpf(d2)/d1*mp.mpf(x)/mp.mpf(y)
        rounded_pair=float(pair_f)
        direct=(d2/d1)*(x/y)
        return dict(total_signed=float((mp.mpf(value)-ref)/ref),
                    pair_signed=float((pair_f-ref)/ref),
                    transform_signed=float((mp.mpf(value)-pair_f)/ref),
                    rounded_pair_error=float(abs(mp.mpf(rounded_pair)-ref)/ref),
                    direct_ratio_error=float(abs(mp.mpf(direct)-ref)/ref),
                    log_result=value.hex(),direct_result=direct.hex(),
                    rounded_pair_result=rounded_pair.hex())


def main():
    old=json.loads((HERE/'pair_cf_screen_result.json').read_text())
    rows=[]
    for r in old['rows']:
        p,d1,d2=map(float.fromhex,r['inputs'])
        ref=baseline.reference(p,d1,d2,80)[0]
        out=dict(inputs=r['inputs'],kind=r['kind'],experiments={})
        for mode in ['baseline','exact_prefactor','tight_cf']:
            value,x,y,trace=experiment(p,d1,d2,mode)
            m=measures(value,x,y,ref,d1,d2)
            m.update(pair=[x.hex(),y.hex()],iterations=len(trace),last_iteration=trace[-1])
            out['experiments'][mode]=m
        rows.append(out)
        print(p,d1,d2,{k:(v['pair_signed'],v['total_signed']) for k,v in out['experiments'].items()},flush=True)
    paths=[Path(__file__),Path(candidate.__file__),Path(model.__file__),HERE/'pair_cf_screen_result.json']
    result=dict(status='EXPLORATORY_NOT_ADOPTION',holdout_evaluated=False,rows=rows,
                bindings={str(p.relative_to(model.ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in paths})
    (HERE/'pair_floor_diagnostic.json').write_text(json.dumps(result,indent=2)+'\n')


if __name__=='__main__':
    main()
