"""Frozen staged P4 screen: revealed rows first, 47 new rows only on a pass."""
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
import tail_aware_model as p4

HERE = Path(__file__).resolve().parent
BOUND, FLOOR = 2.0 ** -48, 2.0 ** -52
PROBE_ROWS = ((1 - 1e-6, 5., 2e6), (1 - 1e-10, 5., 2e6), (1 - 1e-13, 5., 2e6), (1e-10, 5., 2e6),
              (1 - 1e-10, 2e6, 5.), (1e-10, 2e6, 5.), (1 - 1e-10, 100., 1e10), (1e-12, 8., 5e4),
              (1 - 1e-12, 8., 5e4), (1 - 1e-12, 30., 40.))


def new_rows():
    rows = [('far', p, d1, d2) for p in (1e-9, 1 - 1e-9, 1e-14, 1 - 1e-14)
            for d1, d2 in ((4., 3e5), (3e5, 4.), (60., 5e8), (5e8, 60.), (300., 500.))]
    rows += [('middle', p, d1, d2) for p, (d1, d2) in
             itertools.product((.3, .7), ((9e4, 12.), (12., 9e4), (4e6, 4e6)))]
    for d1, d2 in ((6., 10.), (10., 6.), (80., 120.)):
        h = p4.tails(0.5, d1 / 2.0, d2 / 2.0)[0]
        rows += [('orientation_seam', v, d1, d2) for v in
                 (_step(h, -4), _step(h, -1), h, _step(h, 1), _step(h, 4))]
    half = (math.nextafter(.5, 0.), .5, math.nextafter(.5, 1.))
    rows += [('tail_seam', p, d1, d2) for p in half for d1, d2 in ((7., 2e4), (2e4, 7.))]
    return rows


def _step(x, n):
    for _ in range(abs(n)):
        x = math.nextafter(x, 1.0 if n > 0 else 0.0)
    return x


def evaluate_row(group, kind, p, d1, d2):
    started = time.monotonic()
    row = dict(group=group, kind=kind, inputs=[v.hex() for v in (p, d1, d2)],
               complement=(1 - p).hex(), passes=False)
    try:
        r50 = extended.baseline.reference(p, d1, d2, 50)
        r80 = extended.baseline.reference(p, d1, d2, 80)
        with mp.workdps(80):
            stable = all(abs(a - b) <= mp.mpf('1e-40') * abs(b) for a, b in zip(r50, r80))
        variants = {}
        row['variants'] = variants
        for name, fn in (('F-0', extended.model.f_inverse), ('P3', p3.evaluate), ('P4', p4.evaluate)):
            try:
                value, x, y, trace = fn(p, d1, d2)
            except (ValueError, ArithmeticError) as exc:
                if name == 'P4':
                    raise
                variants[name] = dict(failure=str(exc))
                continue
            with mp.workdps(80):
                error = float(abs(mp.mpf(value) - r80[0]) / abs(r80[0]))
            variants[name] = dict(result=value.hex(), pair=[x.hex(), y.hex()], error=error,
                                  iterations=len(trace),
                                  guard_activations=sum(t.get('guarded_stagnation', False) for t in trace))
        new, base = variants['P4'], variants['F-0']
        if 'error' not in base:
            raise ValueError('baseline failed: ' + base['failure'])
        gates = dict(oracle=stable)
        if group in ('retained', 'revealed_screen', 'new'):
            gates['f_threshold'] = new['error'] < BOUND
        if kind == 'control':
            gates['control_nonregression'] = new['error'] <= 4 * base['error']
        if group == 'fit':
            gates.update(f_nonregression=new['error'] <= 2 * base['error'],
                         clean_f_threshold=base['error'] >= BOUND or new['error'] < BOUND)
        if group == 'probe':
            gates.update(f_threshold=new['error'] < BOUND, no_worse=new['error'] <= base['error'])
        if group == 'stage2':
            gates.update(f_threshold=new['error'] < BOUND,
                         f_nonregression=new['error'] <= max(2 * base['error'], FLOOR))
        row.update(gates=gates, passes=all(gates.values()))
    except (ValueError, ArithmeticError) as exc:
        row['evidence_blocker'] = str(exc)
    row['seconds'] = time.monotonic() - started
    shown = {k: f"{v['error']:.2e}" for k, v in row.get('variants', {}).items() if 'error' in v}
    print(group, kind, p, d1, d2, shown, row.get('gates', row.get('evidence_blocker')), flush=True)
    return row


def median_gate(rows):
    return (statistics.median(r['variants']['P4']['iterations'] for r in rows)
            <= statistics.median(r['variants']['F-0']['iterations'] for r in rows) + 2)


def main():
    root = extended.model.ROOT
    paths = [Path(__file__), Path(p4.__file__), Path(p3.__file__), HERE / 'TAIL_AWARE_PROTOCOL.md',
             Path(extended.__file__), Path(p4.pair.__file__), Path(p4.pair.DD_MODULE.__file__),
             Path(extended.model.__file__), Path(extended.model.forward.__file__),
             Path(extended.model.normal.__file__), Path(extended.baseline.__file__),
             root / 'benchmark/_ibeta.py', HERE / 'stagnation_result.json']
    bindings = {str(p.relative_to(root)): hashlib.sha256(p.read_bytes()).hexdigest() for p in paths}
    revealed = json.loads((HERE / 'stagnation_result.json').read_text())['rows']
    points = [(r['group'], r['kind'], *map(float.fromhex, r['inputs'])) for r in revealed]
    points += [('probe', 'far', *row) for row in PROBE_ROWS]
    result = dict(status='INCOMPLETE', holdout_evaluated=False, stage2_evaluated=False,
                  bindings=bindings, rows=[])

    def save():
        (HERE / 'tail_aware_result.json').write_text(json.dumps(result, indent=2) + '\n')

    original_cf = extended.oracle.betacf
    extended.oracle.betacf = extended.checked_betacf
    try:
        for point in points:
            result['rows'].append(evaluate_row(*point))
            save()
        result['fit_median_gate'] = median_gate([r for r in result['rows'] if r['group'] == 'fit'])
        stage1 = all(r['passes'] for r in result['rows']) and result['fit_median_gate']
        if stage1:
            result['stage2_evaluated'] = True
            for kind, p, d1, d2 in new_rows():
                result['rows'].append(evaluate_row('stage2', kind, p, d1, d2))
                save()
            stage2 = [r for r in result['rows'] if r['group'] == 'stage2']
            result['stage2_median_gate'] = (all('variants' in r and 'error' in r['variants'].get('P4', {})
                                                for r in stage2) and median_gate(stage2))
        result['status'] = ('SCREEN_PASS_NOT_ADOPTION' if stage1 and all(r['passes'] for r in result['rows'])
                            and result.get('stage2_median_gate') else 'REJECT_OR_BLOCKED')
        save()
    finally:
        extended.oracle.betacf = original_cf
    print(result['status'], flush=True)


if __name__ == '__main__':
    main()
