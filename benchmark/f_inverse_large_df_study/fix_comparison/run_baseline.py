"""Evaluate only frozen retained-row faithfulness; never opens holdout data."""
import hashlib
import json
from pathlib import Path
import sys
import time

import mpmath as mp
import vba_order_model as model

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
from run_isolation import AFFECTED, CONTROLS
from f_inverse_model import beta_invcdf


def reference(p, d1, d2, dps):
    with mp.workdps(dps):
        x = beta_invcdf(mp.mpf(p), mp.mpf(d1) / 2, mp.mpf(d2) / 2)
        y = 1 - x
        return mp.mpf(d2) / mp.mpf(d1) * x / y, x, y


def main():
    rows = []
    started = time.monotonic()
    for kind, points in [('affected', AFFECTED), ('control', CONTROLS)]:
        for p, d1, d2, observed in points:
            row = dict(kind=kind, inputs=[v.hex() for v in (p, d1, d2)],
                       complement=(1.0-p).hex(), historical_error=observed)
            try:
                value, x, y, trace = model.f_inverse(p, d1, d2)
                r50 = reference(p, d1, d2, 50)
                r80 = reference(p, d1, d2, 80)
                with mp.workdps(80):
                    stable = all(abs(a-b) <= mp.mpf('1e-40') * abs(b) for a,b in zip(r50,r80))
                    error = float(abs(mp.mpf(value)-r80[0])/r80[0])
                    tiny = float(min(r80[1:]))
                lower, upper = 2**-52/(8*tiny), 8*2**-52/tiny
                passed = lower <= error <= upper if kind == 'affected' else error < 2**-48
                row.update(result=value.hex(), pair=[x.hex(),y.hex()], error=error,
                           bound=[lower,upper] if kind=='affected' else [0,2**-48],
                           oracle_stable=stable, passes=passed and stable,
                           iterations=len(trace), trace=trace)
            except (ValueError, ArithmeticError) as exc:
                row.update(passes=False, failure=str(exc))
            rows.append(row)
            print(kind, p, d1, d2, row.get('error',row.get('failure')), row['passes'], flush=True)
    paths = [Path(model.__file__), Path(model.forward.__file__), Path(model.normal.__file__), Path(__file__),
             HERE.parent/'FIX_COMPARISON_PREREGISTRATION.md'] + list((model.ROOT/'src').glob('*.bas'))
    result = dict(status='PASS' if all(r['passes'] for r in rows) else 'FAIL_STOP',
                  candidates_evaluated=False, holdout_evaluated=False,
                  iteration_comparison='No Excel iteration counts supplied in the retained-row tables',
                  bindings={str(p.relative_to(model.ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in paths},
                  seconds=time.monotonic()-started, rows=rows)
    (HERE/'baseline_result.json').write_text(json.dumps(result,indent=2)+'\n')
    print(result['status'])


if __name__ == '__main__':
    main()
