"""Revealed-fit screening of B4. Never reads the reserved holdout design/grid."""
import csv
import hashlib
import json
import struct
from fractions import Fraction
from pathlib import Path

import pair_candidate as candidate
from analyze_binary64_map import BOUNDS, UNDERFLOW, regime

HERE = Path(__file__).resolve().parent


def bits(x):
    return struct.unpack('>Q', struct.pack('>d', x))[0]


def value(n):
    return struct.unpack('>d', struct.pack('>Q', n))[0]


def evaluate():
    with (HERE/'b64_fit_grid.csv').open() as stream:
        fit = list(csv.DictReader(stream))
    retained = json.loads((HERE/'arithmetic_diagnosis.json').read_text())['rows']
    failures, declined, fit_rows, windows = [], [], [], []
    worst = {'P': 0., 'Q': 0.}
    for row in fit:
        a, x = float(row['shape']), float(row['x'])
        try:
            trace = {}
            pair = candidate.public(a, x, trace)
        except candidate.Declined as exc:
            declined.append(dict(point_id=row['point_id'], reason=str(exc)))
            continue
        limits = BOUNDS[regime(a, x)]
        for tail, v, ref, limit in zip(('P', 'Q'), pair, (row['ref_p'], row['ref_q']), limits):
            reference = Fraction(ref)
            if reference >= UNDERFLOW:
                error = float(abs(Fraction(v)-reference)/reference)
                worst[tail] = max(worst[tail], error)
                if error > limit:
                    failures.append(dict(kind='fit accuracy', point=row['point_id'], tail=tail, error=error, limit=limit))
        fit_rows.append(dict(point_id=row['point_id'], values=[v.hex() for v in pair], trace=trace))
    centers = {(float(row['shape']), float(row['shape'])+1.) for row in fit
               if 1e-8 <= float(row['shape']) <= 10000}
    centers |= {(float.fromhex(r['a']), float.fromhex(r['left'])) for r in retained}
    retained_failures = 0
    for r in retained:
        a = float.fromhex(r['a'])
        left = candidate.public(a, float.fromhex(r['left']))
        right = candidate.public(a, float.fromhex(r['right']))
        if right[0] < left[0] or right[1] > left[1]:
            retained_failures += 1
            failures.append(dict(kind='retained monotonicity', row=r))
    for a, center in sorted(centers):
        xs = [value(n) for n in range(bits(center)-16, bits(center)+17)]
        pairs = [candidate.public(a, x) for x in xs]
        for i in range(32):
            if pairs[i+1][0] < pairs[i][0] or pairs[i+1][1] > pairs[i][1]:
                failures.append(dict(kind='wide monotonicity', a=a.hex(), left=xs[i].hex(), right=xs[i+1].hex()))
        windows.append(dict(a=a.hex(), center=center.hex(), x=[x.hex() for x in xs],
                            values=[[v.hex() for v in pair] for pair in pairs]))
    return dict(kind='B4 development only; no host fidelity, timing, or holdout acceptance',
                status='REJECT' if failures else 'DEVELOPMENT PASS ONLY',
                fit_answered=len(fit_rows), fit_declined=len(declined), worst_relative=worst,
                retained_pairs=len(retained), retained_failures=retained_failures,
                windows=len(windows), window_evaluations=len(windows)*33,
                failures=failures, declined=declined, fit_rows=fit_rows,
                windows_sha256=hashlib.sha256(json.dumps(windows, sort_keys=True).encode()).hexdigest(),
                sha256={f: hashlib.sha256((HERE/f).read_bytes()).hexdigest() for f in
                        ('pair_candidate.py','screen_pair_candidate.py','b64_fit_grid.csv','arithmetic_diagnosis.json')})


if __name__ == '__main__':
    out = evaluate()
    (HERE/'pair_candidate_results.json').write_text(json.dumps(out, indent=2)+'\n')
    print(json.dumps({k:v for k,v in out.items() if k not in ('fit_rows','declined','sha256')}, indent=2))
    raise SystemExit(bool(out['failures']))
