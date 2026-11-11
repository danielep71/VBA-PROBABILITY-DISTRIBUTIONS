"""Replay the frozen rule on already revealed fit coordinates; never read holdout inputs."""
import csv
import hashlib
import json
import math
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import route_model as rm
sys.path.insert(0, str(HERE.parent / 'gamma_cf_study' / 'coverage'))
from wide_gamma import public


def routed(a, x):
    if a > 10000:
        return public(a, x)
    routine, v, _ = rm.route(a, x)
    if v is None or not math.isfinite(v) or not 0 <= v <= 1:
        return public(a, x)
    if routine == 'series':
        return v, rm.sub(1., v) if v <= .5 else public(a, x)[1]
    if v <= .5:
        return rm.sub(1., v), v
    return public(a, x)


def evaluate():
    rows = list(csv.DictReader((HERE / 'b64_fit_grid.csv').open()))
    shapes = sorted({float(r['shape']) for r in rows if 0 < float(r['shape']) <= 10000})
    failures = []
    results = []
    for a in shapes:
        seam = a + 1.
        xs = [math.nextafter(seam, 0.), seam, math.nextafter(seam, math.inf)]
        values = [routed(a, x) for x in xs]
        baseline = [public(a, x) for x in xs]
        result = dict(a_hex=a.hex(), x_hex=[x.hex() for x in xs], values_hex=[[v.hex() for v in pair] for pair in values], baseline_hex=[[v.hex() for v in pair] for pair in baseline])
        results.append(result)
        for i in range(2):
            for tail in (0, 1):
                if (values[i+1][tail] < values[i][tail]) if tail == 0 else (values[i+1][tail] > values[i][tail]):
                    failures.append(dict(a_hex=a.hex(), left_x_hex=xs[i].hex(), right_x_hex=xs[i+1].hex(), tail='P' if tail == 0 else 'Q', left=values[i][tail].hex(), right=values[i+1][tail].hex()))
    out = dict(rule_commit='63f88c895d4040cd4a7c25c2cf19b0daf34c90c6', kind='development seam probes at revealed fit shapes; not holdout', status='REJECT' if failures else 'PASS', shapes=len(shapes), failures=failures, rows=results)
    out['sha256'] = {name: hashlib.sha256((HERE / name).read_bytes()).hexdigest() for name in ('route_model.py', 'b64_fit_grid.csv', 'ROUTING_RULE.md', 'preflight_phase_b.py')}
    return out


def main():
    out = evaluate()
    (HERE / 'phase_b_preflight.json').write_text(json.dumps(out, indent=2)+'\n')
    print(json.dumps({k:v for k,v in out.items() if k != 'rows'}, indent=2))
    return out["status"] != "PASS"


if __name__ == '__main__':
    raise SystemExit(main())
