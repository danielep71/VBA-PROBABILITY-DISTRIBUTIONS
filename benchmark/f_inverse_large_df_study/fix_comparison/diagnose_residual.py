"""Exploratory substitutions on revealed retained rows, NOT confirmatory evidence.

Never loads the prospective fitting grid, invariance set or independent holdout.
Each substitution restores the original model before the next experiment.
"""
import hashlib
import json
from pathlib import Path

import mpmath as mp
import run_baseline as baseline
import vba_order_model as model
from _ibeta import betacf, ibeta

HERE = Path(__file__).resolve().parent
ORIGINAL_FORWARD = model.forward.beta_regularized
ORIGINAL_CF = model.forward.beta_continued_fraction


def prefactor_forward(x, y, a, b, trace=None, exact_complement=False, **unused):
    if x <= 0:
        return 0.0
    if y <= 0:
        return 1.0
    xx, yy, aa, bb = map(mp.mpf, (x, y, a, b))
    if exact_complement:
        yy = 1-xx
    log_bt = aa*mp.log(xx)+bb*mp.log(yy)-mp.log(mp.beta(aa, bb))
    bt = float(mp.exp(log_bt))
    if x < (a+1.0)/(a+b+2.0):
        cf, iterations = ORIGINAL_CF(a,b,x)
        value = bt*cf/a
    else:
        cf, iterations = ORIGINAL_CF(b,a,y)
        value = 1.0-bt*cf/b
    if trace is not None:
        trace.update(substitution='prefactor',exact_complement=exact_complement,
                     iterations=iterations)
    return min(max(value,0.0),1.0)


def exact_forward(x,y,a,b,trace=None,**unused):
    if trace is not None:
        trace['substitution'] = 'entire forward with exact complement'
    return float(ibeta(mp.mpf(x),mp.mpf(a),mp.mpf(b)))


def run_one(p,d1,d2,mode,dps):
    model.forward.beta_regularized = ORIGINAL_FORWARD
    model.forward.beta_continued_fraction = ORIGINAL_CF
    try:
        with mp.workdps(dps):
            if mode == 'forward':
                model.forward.beta_regularized = exact_forward
            elif mode == 'prefactor_supplied':
                model.forward.beta_regularized = prefactor_forward
            elif mode == 'prefactor_exact_complement':
                def replacement(*args,**kwargs):
                    return prefactor_forward(*args,**kwargs,exact_complement=True)
                model.forward.beta_regularized = replacement
            elif mode in ('cf_exact_complement', 'joint_prefactor_cf'):
                def paired_forward(x,y,a,b,trace=None,**unused):
                    if x <= 0:
                        return 0.0
                    if y <= 0:
                        return 1.0
                    xx,aa,bb=map(mp.mpf,(x,a,b))
                    yy=1-xx
                    if mode == 'joint_prefactor_cf':
                        bt=float(mp.exp(aa*mp.log(xx)+bb*mp.log(yy)-mp.log(mp.beta(aa,bb))))
                    else:
                        parts={}
                        ORIGINAL_FORWARD(x,y,a,b,trace=parts)
                        bt=parts['bt']
                    if x < (a+1.0)/(a+b+2.0):
                        value=bt*float(betacf(aa,bb,xx))/a
                    else:
                        value=1.0-bt*float(betacf(bb,aa,yy))/b
                    if trace is not None:
                        trace['substitution']=mode
                    return min(max(value,0.0),1.0)
                model.forward.beta_regularized=paired_forward
            elif mode == 'cf':
                model.forward.beta_continued_fraction = lambda a,b,x: (float(betacf(mp.mpf(a),mp.mpf(b),mp.mpf(x))),-1)
            value,x,y,trace = model.f_inverse(p,d1,d2,'F-A')
            return value,x,y,trace
    finally:
        model.forward.beta_regularized = ORIGINAL_FORWARD
        model.forward.beta_continued_fraction = ORIGINAL_CF


def main():
    rows=[]
    # Already revealed retained rows only, including the clean controls.
    for kind,points in [('affected',baseline.AFFECTED),('control',baseline.CONTROLS)]:
        for p,d1,d2,_ in points:
            ref = baseline.reference(p,d1,d2,80)
            row=dict(kind=kind,inputs=[v.hex() for v in (p,d1,d2)],experiments={})
            for mode in ['baseline','forward','prefactor_supplied','prefactor_exact_complement','cf','cf_exact_complement','joint_prefactor_cf']:
                try:
                    v50=run_one(p,d1,d2,mode,50)
                    v80=run_one(p,d1,d2,mode,80)
                    with mp.workdps(80):
                        value,x,y,trace=v80
                        error=float(abs(mp.mpf(value)-ref[0])/ref[0])
                        # Exact final transform from the same rounded pair isolates C7.
                        pair_f=mp.mpf(d2)/d1*mp.mpf(x)/mp.mpf(y)
                        pair_error=float(abs(pair_f-ref[0])/ref[0])
                        transform_error=float(abs(value-pair_f)/abs(pair_f))
                    row['experiments'][mode]=dict(error=error,pair_error=pair_error,
                        transform_error=transform_error,result=value.hex(),pair=[x.hex(),y.hex()],
                        precision_stable=v50[:3]==v80[:3],iterations=len(trace),
                        last_iteration=trace[-1])
                except (ValueError,ArithmeticError) as exc:
                    row['experiments'][mode]=dict(failure=str(exc))
            rows.append(row)
            print(kind,p,d1,d2,{k:v.get('error',v.get('failure')) for k,v in row['experiments'].items()},flush=True)
    paths=[Path(__file__),Path(model.__file__),Path(model.forward.__file__),Path(model.normal.__file__),baseline.HERE.parent/'f_inverse_model.py',model.ROOT/'benchmark/_ibeta.py']
    result=dict(status='EXPLORATORY_NOT_ADOPTION',candidate='F-A',holdout_evaluated=False,
                bindings={str(p.relative_to(model.ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in paths},rows=rows)
    (HERE/'residual_diagnostic.json').write_text(json.dumps(result,indent=2)+'\n')


if __name__ == '__main__':
    main()
