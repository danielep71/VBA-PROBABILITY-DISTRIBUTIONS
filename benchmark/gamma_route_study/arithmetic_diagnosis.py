"""Attribute revealed B1/B3 counterexamples; oracle substitutions, not a candidate.

Reads only retained development failures. No holdout inputs are opened. Both
oracle interventions round to binary64 at the original composition boundary.
They isolate errors but are not deployable algorithms or timing evidence.
"""
import hashlib
import json
from pathlib import Path

import mpmath as mp
import preflight_phase_b as old

HERE = Path(__file__).resolve().parent
SOURCES = ('phase_b_preflight.json', 'smaller_tail_development.json')


def reference(a, x, dps):
    with mp.workdps(dps):
        aa, xx = mp.mpf(a), mp.mpf(x)
        g = mp.gamma(aa)
        p = mp.gammainc(aa, 0, xx) / g
        q = mp.gammainc(aa, xx, mp.inf) / g
        f = mp.exp(aa * mp.log(xx) - xx) / g
        return p, q, f


def direct(a, x, intervention, ref):
    rm = old.rm
    routine = 'series' if rm.below_seam(a, x) else 'cf'
    original_factor = rm.factor
    if intervention == 'reference_recurrence':
        with mp.workdps(120):
            recurrence = float(ref[0 if routine == 'series' else 1] / ref[2])
        return routine, rm.mul(original_factor(a, x), recurrence)
    if intervention == 'reference_factor':
        rm.factor = lambda _a, _x: float(ref[2])
    try:
        _, value, _ = rm.route(a, x)
        return routine, value
    finally:
        rm.factor = original_factor


def routed(a, x, rule, intervention, ref):
    if intervention == 'reference_pair':
        return float(ref[0]), float(ref[1])
    routine, v = direct(a, x, intervention, ref)
    if v is None or not 0 <= v <= 1 or (rule == 'B3' and v > .5):
        return old.public(a, x)
    if routine == 'series':
        return v, old.rm.sub(1., v) if v <= .5 else old.public(a, x)[1]
    return (old.rm.sub(1., v), v) if v <= .5 else old.public(a, x)


def evaluate():
    b1 = json.loads((HERE / SOURCES[0]).read_text())['failures']
    b3 = json.loads((HERE / SOURCES[1]).read_text())['wide']['failures']
    cases = [('B1', float.fromhex(r['a_hex']), float.fromhex(r['left_x_hex']),
              float.fromhex(r['right_x_hex'])) for r in b1]
    cases += [('B3', r['a'], float.fromhex(r['left']), float.fromhex(r['right'])) for r in b3]
    modes = ('original', 'reference_factor', 'reference_recurrence', 'reference_pair')
    rows = []
    for rule, a, left, right in cases:
        refs = [reference(a, x, 120) for x in (left, right)]
        for x, high in zip((left, right), refs):
            low = reference(a, x, 80)
            with mp.workdps(120):
                assert all(abs(l-h) <= abs(h)*mp.mpf('1e-65') for l, h in zip(low, high))
        results = {}
        for mode in modes:
            values = [routed(a, x, rule, mode, ref) for x, ref in zip((left, right), refs)]
            fail = values[1][0] < values[0][0] or values[1][1] > values[0][1]
            results[mode] = dict(monotone=not fail, values=[[v.hex() for v in pair] for pair in values])
        assert not results['original']['monotone'], 'retained failure no longer reproduces'
        assert results['reference_pair']['monotone'], 'independent reference decreases'
        rows.append(dict(rule=rule, a=a.hex(), left=left.hex(), right=right.hex(), results=results))
    counts = {rule: {mode: sum(not r['results'][mode]['monotone'] for r in rows if r['rule'] == rule)
                     for mode in modes} for rule in ('B1', 'B3')}
    return dict(kind='oracle attribution on revealed failures only; not a routing candidate',
                status='DIAGNOSIS ONLY', points=len(rows)*2, pairs=len(rows),
                remaining_failures=counts, rows=rows,
                sha256={f: hashlib.sha256((HERE/f).read_bytes()).hexdigest()
                        for f in (*SOURCES, 'route_model.py', 'preflight_phase_b.py', 'arithmetic_diagnosis.py')})


if __name__ == '__main__':
    result = evaluate()
    (HERE / 'arithmetic_diagnosis.json').write_text(json.dumps(result, indent=2)+'\n')
    print(json.dumps(result['remaining_failures'], indent=2))
