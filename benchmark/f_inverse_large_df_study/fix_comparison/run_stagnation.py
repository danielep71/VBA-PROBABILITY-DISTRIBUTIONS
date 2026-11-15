"""Frozen staged P3 screen: revealed fitting first, eight new rows only on pass."""
import hashlib
import itertools
import json
import math
from pathlib import Path
import statistics
import time

import mpmath as mp
import run_extended_fit as extended
import stagnation_model as p3

HERE = Path(__file__).resolve().parent
NEW_ROWS = tuple(itertools.product((.15, .85),
    ((5e4, 8.), (8., 5e4), (2000., 2000.), (2e6, 2e6))))


def median_gate(rows):
    return all(len(r.get('variants', {})) == 2 for r in rows) and (
        statistics.median(r['variants']['P3']['iterations'] for r in rows)
        <= statistics.median(r['variants']['F-0']['iterations'] for r in rows)+2)


def evaluate_row(group, kind, p, d1, d2):
    started = time.monotonic()
    row = dict(group=group, kind=kind, inputs=[v.hex() for v in (p,d1,d2)],
               complement=(1-p).hex(), passes=False)
    try:
        r50 = extended.baseline.reference(p,d1,d2,50)
        r80 = extended.baseline.reference(p,d1,d2,80)
        with mp.workdps(80):
            stable = all(abs(a-b)<=mp.mpf('1e-40')*abs(b) for a,b in zip(r50,r80))
        variants = {}
        row['variants'] = variants
        for name, fn in [('F-0', extended.model.f_inverse), ('P3', p3.evaluate)]:
            t0 = time.monotonic()
            value, x, y, trace = fn(p,d1,d2)
            if not all(math.isfinite(v) and v > 0 for v in (value,x,y)):
                raise ArithmeticError('nonfinite/nonpositive output: '+name)
            with mp.workdps(80):
                errors = [float(abs(mp.mpf(v)-ref)/abs(ref)) for v,ref in zip((value,x,y),r80)]
                residual = float(abs(extended.f_cdf(value,d1,d2)-mp.mpf(p))/min(mp.mpf(p),1-mp.mpf(p)))
            variants[name] = dict(result=value.hex(), pair=[x.hex(),y.hex()], error=errors[0],
                pair_errors=errors[1:], tail_residual=residual, iterations=len(trace),
                guard_activations=sum(t.get('guarded_stagnation', False) for t in trace),
                trace=trace, seconds=time.monotonic()-t0)
        base, new = variants['F-0'], variants['P3']
        gates = dict(oracle=stable)
        if group in ('retained', 'revealed_screen', 'new'):
            gates['f_threshold'] = new['error']<2**-48
        if kind == 'control':
            gates['control_nonregression'] = new['error']<=4*base['error']
        if group in ('fit', 'new'):
            gates.update(f_nonregression=new['error']<=2*base['error'],
                clean_f_threshold=base['error']>=2**-48 or new['error']<2**-48)
        if group == 'fit' and d1 == d2 == 1e8:
            gates.update(pair_nonregression=all(a<=2*b for a,b in zip(new['pair_errors'],base['pair_errors'])),
                residual_nonregression=new['tail_residual']<=2*base['tail_residual'],
                iterations_nonregression=new['iterations']<=2*base['iterations'])
        row.update(gates=gates, passes=all(gates.values()))
    except (ValueError, ArithmeticError) as exc:
        row['evidence_blocker'] = str(exc)
    row['seconds'] = time.monotonic()-started
    print(group,p,d1,d2,row.get('gates',row.get('evidence_blocker')),flush=True)
    return row


def main():
    paths = [Path(__file__), Path(p3.__file__), HERE/'STAGNATION_PROTOCOL.md',
        Path(extended.__file__), Path(extended.p2.__file__), Path(extended.p2.p1.__file__),
        Path(extended.p2.p1.DD_MODULE.__file__), Path(extended.model.__file__),
        Path(extended.model.forward.__file__), Path(extended.model.normal.__file__),
        Path(extended.baseline.__file__), extended.model.ROOT/'benchmark/_ibeta.py',
        HERE/'baseline_result.json'] + list((extended.model.ROOT/'src').glob('*.bas'))
    bindings = {str(p.relative_to(extended.model.ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in paths}
    retained = json.loads((HERE/'baseline_result.json').read_text())['rows']
    points = [('retained',r['kind'],*map(float.fromhex,r['inputs'])) for r in retained]
    points += [('revealed_screen','screen',p,*dfs) for p,dfs in extended.p2.NEW_ROWS]
    points += [('fit','fit',p,*dfs) for p,dfs in extended.POINTS]
    result = dict(status='INCOMPLETE',holdout_evaluated=False,new_evaluated=False,bindings=bindings,rows=[])
    def save():
        (HERE/'stagnation_result.json').write_text(json.dumps(result,indent=2)+'\n')
    original_cf = extended.oracle.betacf
    extended.oracle.betacf = extended.checked_betacf
    try:
        for point in points:
            result['rows'].append(evaluate_row(*point))
            save()
        result['fit_median_gate'] = median_gate([r for r in result['rows'] if r['group']=='fit'])
        fitting_pass = all(r['passes'] for r in result['rows']) and result['fit_median_gate']
        if fitting_pass:
            result['new_evaluated'] = True
            for p,dfs in NEW_ROWS:
                result['rows'].append(evaluate_row('new','screen',p,*dfs))
                save()
            result['new_median_gate'] = median_gate([r for r in result['rows'] if r['group']=='new'])
        result['status'] = ('SCREEN_PASS_NOT_ADOPTION' if fitting_pass
            and all(r['passes'] for r in result['rows']) and result.get('new_median_gate')
            else 'REJECT_OR_BLOCKED')
        save()
    finally:
        extended.oracle.betacf = original_cf
    print(result['status'],flush=True)


if __name__ == '__main__':
    main()
