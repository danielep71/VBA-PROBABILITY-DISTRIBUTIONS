"""Candidate P1: frozen retained-row screen, not a production implementation."""
import hashlib
import json
import math
from pathlib import Path
import time

import mpmath as mp
import run_baseline as baseline
import vba_order_model as model

HERE = Path(__file__).resolve().parent
DD_MODULE = model.load_model('f35_dd',model.ROOT/'benchmark/gamma_cf_study/coverage/precision_probe.py')
DD = DD_MODULE.DD
ORIGINAL = model.forward.beta_regularized


def checked(value):
    if not all(math.isfinite(v) and abs(v) <= 1e100 for v in (value.hi,value.lo)):
        raise ValueError('pair-CF range')
    return value


def cf_pair(a,b,x):
    aa_shape,bb_shape=DD(a),DD(b)
    qab=aa_shape+bb_shape
    qap=aa_shape+1.
    qam=aa_shape-1.
    tiny=model.forward.C['PROB_FPMIN']

    def clamp(v):
        checked(v)
        return DD(tiny) if abs(v.value()) < tiny else v

    c=DD(1.)
    d=DD(1.)/clamp(DD(1.)-qab*x/qap)
    h=d
    for m in range(1,model.forward.C['PROB_BETA_MAX_ITER']+1):
        m2=2*m
        aa=DD(m)*(bb_shape-m)*x/((qam+m2)*(aa_shape+m2))
        d=DD(1.)/clamp(DD(1.)+aa*d)
        c=clamp(DD(1.)+aa/c)
        h=checked(h*d*c)
        aa=-(aa_shape+m)*(qab+m)*x/((aa_shape+m2)*(qap+m2))
        d=DD(1.)/clamp(DD(1.)+aa*d)
        c=clamp(DD(1.)+aa/c)
        delta=checked(d*c)
        h=checked(h*delta)
        if abs((delta-1.).value()) <= model.forward.C['PROB_NUM_EPS']:
            return h.value(),m
    raise ValueError('pair-CF did not converge')


def forward_pair(x,y,a,b,trace=None,**unused):
    if x <= 0:
        return 0.
    if y <= 0:
        return 1.
    # Reproduce the baseline prefactor without evaluating a baseline CF.
    if a+b >= model.forward.C['PROB_IBETA_LOADER_MIN_SHAPE']:
        log_bt=model.forward.beta_log_pdf(x,y,a,b,math.log(x),math.log(y))+math.log(x)+math.log(y)
    else:
        log_bt=a*math.log(x)+b*math.log(y)-model.log_beta(a,b)
    bt=math.exp(log_bt)
    if x < (a+1.)/(a+b+2.):
        cf,it=cf_pair(a,b,DD(x))
        value=bt*cf/a
        branch='direct'
    else:
        cf,it=cf_pair(b,a,DD(1.)-DD(x))
        value=1.-bt*cf/b
        branch='reflected'
    if trace is not None:
        trace.update(cf=cf,iterations=it,bt=bt,branch=branch)
    return min(max(value,0.),1.)


def evaluate(p,d1,d2):
    saved=model.forward.beta_regularized
    try:
        model.forward.beta_regularized=forward_pair
        return model.f_inverse(p,d1,d2,'F-A')
    finally:
        model.forward.beta_regularized=saved


def main():
    old=json.loads((HERE/'baseline_result.json').read_text())
    if old['status'] != 'PASS':
        raise RuntimeError('baseline failed')
    for path,digest in old['bindings'].items():
        if hashlib.sha256((model.ROOT/path).read_bytes()).hexdigest()!=digest:
            raise RuntimeError('stale baseline '+path)
    rows=[]
    for row in old['rows']:
        p,d1,d2=map(float.fromhex,row['inputs'])
        out=dict(inputs=row['inputs'],kind=row['kind'])
        started=time.monotonic()
        try:
            value,x,y,trace=evaluate(p,d1,d2)
            r50=baseline.reference(p,d1,d2,50)
            r80=baseline.reference(p,d1,d2,80)
            with mp.workdps(80):
                stable=all(abs(a-b)<=mp.mpf('1e-40')*abs(b) for a,b in zip(r50,r80))
                error=float(abs(mp.mpf(value)-r80[0])/r80[0])
            passed=stable and error<2**-48
            if row['kind']=='control':
                passed=passed and error<=4*row['error']
            out.update(error=error,passes=passed,oracle_stable=stable,
                       result=value.hex(),pair=[x.hex(),y.hex()],trace=trace)
        except (ValueError,ArithmeticError) as exc:
            out.update(passes=False,failure=str(exc))
        out['seconds']=time.monotonic()-started
        rows.append(out)
        print(row['kind'],p,d1,d2,out.get('error',out.get('failure')),out['passes'],flush=True)
    paths=[Path(__file__),Path(DD_MODULE.__file__),HERE/'PAIR_CF_SCREEN_PROTOCOL.md',HERE/'baseline_result.json']
    result=dict(status='SCREEN_PASS_NOT_ADOPTION' if all(r['passes'] for r in rows) else 'REJECT',
                holdout_evaluated=False,rows=rows,
                bindings={str(p.relative_to(model.ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in paths})
    (HERE/'pair_cf_screen_result.json').write_text(json.dumps(result,indent=2)+'\n')
    print(result['status'])


if __name__=='__main__':
    main()
