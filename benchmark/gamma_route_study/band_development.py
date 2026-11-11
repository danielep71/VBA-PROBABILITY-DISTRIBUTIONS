"""Development only: diagnose the rejected rule and test a one-sigma DD band.

Uses only revealed fitting shapes. Never reads or evaluates holdout inputs.
The band is abs(X-(A+1)) <= sqrt(A); its new boundaries are tested explicitly.
"""
import csv
import hashlib
import json
import math
from pathlib import Path
import mpmath as mp
import preflight_phase_b as old

HERE = Path(__file__).resolve().parent


def routed(a, x):
    if abs(x - (a + 1.)) <= math.sqrt(a):
        return old.public(a, x)
    return old.routed(a, x)


def evaluate():
    diagnosis = []
    for f in old.evaluate()['failures']:
        a = float.fromhex(f['a_hex'])
        row = dict(shape=a, points=[])
        for key in ('left_x_hex', 'right_x_hex'):
            x = float.fromhex(f[key])
            routine, value, _ = old.rm.route(a, x)
            p, q = old.routed(a, x)
            refs = []
            for prec in (80, 120):
                with mp.workdps(prec):
                    aa, xx = mp.mpf(a), mp.mpf(x)
                    refs.append((mp.gammainc(aa, 0, xx, regularized=True), mp.gammainc(aa, xx, mp.inf, regularized=True)))
            with mp.workdps(120):
                assert all(abs(u/v-1) < mp.mpf('1e-65') for u,v in zip(refs[0], refs[1]))
                ulp = mp.mpf(math.ulp(p))
                row['points'].append(dict(x_hex=x.hex(), routine=routine, direct_hex=value.hex(), p_hex=p.hex(), q_hex=q.hex(), p_error_ulps=mp.nstr((mp.mpf(p)-refs[1][0])/ulp,30), reference_p=mp.nstr(refs[1][0],80), reference_q=mp.nstr(refs[1][1],80)))
        diagnosis.append(row)
    shapes = sorted({float(r['shape']) for r in csv.DictReader((HERE/'b64_fit_grid.csv').open()) if 0 < float(r['shape']) <= 10000})
    rows, failures = [], []
    for a in shapes:
        for label, boundary in [('lower-band', a+1.-math.sqrt(a)), ('old-seam', a+1.), ('upper-band', a+1.+math.sqrt(a))]:
            xs = [math.nextafter(boundary,0.),boundary,math.nextafter(boundary,math.inf)]
            vals = [routed(a,x) for x in xs]
            baseline = [old.public(a,x) for x in xs]
            rows.append(dict(shape=a,boundary=label,x_hex=[x.hex() for x in xs],values_hex=[[v.hex() for v in pair] for pair in vals],baseline_hex=[[v.hex() for v in pair] for pair in baseline]))
            for i in range(2):
                for tail in (0,1):
                    if (vals[i+1][tail] < vals[i][tail]) if tail==0 else (vals[i+1][tail] > vals[i][tail]):
                        failures.append(dict(shape=a,boundary=label,tail='P' if tail==0 else 'Q',left_x_hex=xs[i].hex(),right_x_hex=xs[i+1].hex(),left_hex=vals[i][tail].hex(),right_hex=vals[i+1][tail].hex(),baseline_monotone=baseline[i+1][tail]>=baseline[i][tail] if tail==0 else baseline[i+1][tail]<=baseline[i][tail]))
    return dict(kind='development on revealed fit shapes; no holdout', candidate='B2: double-double where abs(X-(A+1)) <= sqrt(A); otherwise frozen B1; A2=10000', status='REJECT' if failures else 'DEVELOPMENT PASS ONLY', diagnosis=diagnosis,shapes=len(shapes),triples=len(rows),failures=failures,rows=rows,sha256={p:hashlib.sha256((HERE/p).read_bytes()).hexdigest() for p in ('band_development.py','preflight_phase_b.py','route_model.py','b64_fit_grid.csv')})


if __name__=='__main__':
    result=evaluate()
    (HERE/'band_development.json').write_text(json.dumps(result,indent=2)+'\n')
    print(json.dumps({k:v for k,v in result.items() if k not in ('rows','diagnosis','sha256')},indent=2))
    raise SystemExit(result['status']=='REJECT')
