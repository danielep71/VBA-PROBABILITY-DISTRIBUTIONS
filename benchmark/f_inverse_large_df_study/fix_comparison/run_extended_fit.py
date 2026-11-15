"""Prospective 24-row P2 gate, with explicit oracle CF exhaustion checks."""
import hashlib
import itertools
import json
import math
from pathlib import Path
import statistics
import time

import mpmath as mp
import run_baseline as baseline
import run_tight_cf as p2
import vba_order_model as model
from _ibeta import f_cdf
import _ibeta as oracle

HERE = Path(__file__).resolve().parent
POINTS = tuple(itertools.product((.05, .25, .75, .999),
    ((1e5, 2.), (1e7, 4.), (3., 1e7), (10., 1e9), (1e3, 1e3), (1e8, 1e8))))


def checked_betacf(a, b, x):
    """Same oracle operations and cap; exhaustion cannot look like success."""
    tiny = mp.mpf('1e-300')
    qab, qap, qam = a+b, a+1, a-1
    c, d = mp.mpf(1), 1-qab*x/qap
    if abs(d) < tiny:
        d = tiny
    d = 1/d
    h = d
    for m in range(1, 20000):
        m2 = 2*m
        aa = m*(b-m)*x/((qam+m2)*(a+m2))
        d = 1+aa*d
        if abs(d) < tiny:
            d = tiny
        c = 1+aa/c
        if abs(c) < tiny:
            c = tiny
        d = 1/d
        h *= d*c
        aa = -(a+m)*(qab+m)*x/((a+m2)*(qap+m2))
        d = 1+aa*d
        if abs(d) < tiny:
            d = tiny
        c = 1+aa/c
        if abs(c) < tiny:
            c = tiny
        d = 1/d
        delta = d*c
        h *= delta
        if abs(delta-1) < mp.mpf('1e-45'):
            return h
    raise ArithmeticError('oracle CF exhausted 19999 iterations')


def main():
    paths = [Path(__file__), HERE/'EXTENDED_FIT_PROTOCOL.md', Path(p2.__file__),
        Path(p2.p1.__file__), Path(p2.p1.DD_MODULE.__file__), Path(model.__file__),
        Path(model.forward.__file__), Path(model.normal.__file__),
        Path(baseline.__file__), model.ROOT/'benchmark/_ibeta.py']
    paths += list((model.ROOT/'src').glob('*.bas'))
    bindings = {str(p.relative_to(model.ROOT)): hashlib.sha256(p.read_bytes()).hexdigest() for p in paths}
    rows = []
    original_cf = oracle.betacf
    oracle.betacf = checked_betacf
    try:
        for p, (d1, d2) in POINTS:
            started = time.monotonic()
            row = dict(inputs=[v.hex() for v in (p,d1,d2)], complement=(1-p).hex(),
                       balanced_huge=d1 == d2 == 1e8, passes=False)
            try:
                r50 = baseline.reference(p,d1,d2,50)
                r80 = baseline.reference(p,d1,d2,80)
                with mp.workdps(80):
                    row['oracle_stable'] = all(abs(a-b)<=mp.mpf('1e-40')*abs(b) for a,b in zip(r50,r80))
                variants = {}
                row['variants'] = variants
                for name, fn in [('F-0', model.f_inverse), ('P2', p2.evaluate)]:
                    t0 = time.monotonic()
                    value, x, y, trace = fn(p,d1,d2)
                    if not all(math.isfinite(v) and v > 0 for v in (value,x,y)):
                        raise ArithmeticError('nonfinite/nonpositive output: '+name)
                    with mp.workdps(80):
                        errors = [float(abs(mp.mpf(v)-ref)/abs(ref)) for v,ref in zip((value,x,y),r80)]
                        residual = float(abs(f_cdf(value,d1,d2)-mp.mpf(p))/min(mp.mpf(p),1-mp.mpf(p)))
                    variants[name] = dict(result=value.hex(), pair=[x.hex(),y.hex()],
                        error=errors[0], pair_errors=errors[1:], tail_residual=residual,
                        iterations=len(trace), trace=trace, seconds=time.monotonic()-t0)
                base, new = variants['F-0'], variants['P2']
                gates = dict(oracle=row['oracle_stable'],
                    f_nonregression=new['error']<=2*base['error'],
                    clean_f_threshold=base['error']>=2**-48 or new['error']<2**-48)
                if row['balanced_huge']:
                    gates.update(pair_nonregression=all(a<=2*b for a,b in zip(new['pair_errors'],base['pair_errors'])),
                        residual_nonregression=new['tail_residual']<=2*base['tail_residual'],
                        iterations_nonregression=new['iterations']<=2*base['iterations'])
                row.update(gates=gates, passes=all(gates.values()))
            except (ValueError, ArithmeticError) as exc:
                row['evidence_blocker'] = str(exc)
            row['seconds'] = time.monotonic()-started
            rows.append(row)
            print(p,d1,d2,row.get('gates', row.get('evidence_blocker')),flush=True)
            # Checkpoint after every complete row; never discard failures.
            (HERE/'extended_fit_result.json').write_text(json.dumps(dict(
                status='INCOMPLETE', holdout_evaluated=False, bindings=bindings, rows=rows),indent=2)+'\n')
    finally:
        oracle.betacf = original_cf
    complete = all(len(r.get('variants',{})) == 2 for r in rows)
    median_gate = complete and statistics.median(r['variants']['P2']['iterations'] for r in rows) <= statistics.median(r['variants']['F-0']['iterations'] for r in rows)+2
    status = 'FIT_PASS_NOT_ADOPTION' if all(r['passes'] for r in rows) and median_gate else 'REJECT_OR_BLOCKED'
    (HERE/'extended_fit_result.json').write_text(json.dumps(dict(status=status,
        holdout_evaluated=False, median_iteration_gate=median_gate, bindings=bindings, rows=rows),indent=2)+'\n')
    print(status,flush=True)


if __name__ == '__main__':
    main()
