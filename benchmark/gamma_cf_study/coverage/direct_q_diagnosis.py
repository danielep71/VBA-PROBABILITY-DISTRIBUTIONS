"""Retrospective #24 diagnostics on revealed B7 points, not an acceptance screen.

The math.fsum/log1p/expm1 arm is deliberately optimistic: these Python
operations are not claimed to reproduce VBA. Exact-arithmetic arms isolate
error sources; they are not implementable production candidates.
"""
import hashlib
import json
import math
import platform
import re
from pathlib import Path

import mpmath as mp
import seam_diagnosis as model

HERE = Path(__file__).resolve().parent
X, D = model.X, model.D


def direct(a, x, mode):
    if a > 1 or x >= a + 1:
        return model.public(a, x)
    if mode == 'binary':
        term = -x
        terms = []
        for n in range(1, 100001):
            value = D(X(term) / (X(a) + n))
            terms.append(value)
            if n > 4 and abs(value) < 2**-54 * abs(math.fsum(terms)):
                break
            term = D(-X(term) * X(x) / (n + 1))
        else:
            raise RuntimeError('series cap')
        z = D(X(a) * X(math.fsum(terms)))
        lp = D(X(a) * X(math.log(x)) - X(model.lg(a)) + X(math.log1p(z)))
        if lp > 0 or not math.isfinite(lp):
            raise RuntimeError('invalid log probability')
        return math.exp(lp), -math.expm1(lp)
    with mp.workdps(100):
        aa, xx = mp.mpf(a), mp.mpf(x)
        if mode == 'ideal_lower':
            return (float(mp.gammainc(aa, 0, xx, regularized=True)),
                    float(mp.gammainc(aa, xx, mp.inf, regularized=True)))
        term = -xx
        total = term / (aa + 1)
        for n in range(2, 1000):
            term *= -xx / n
            value = term / (aa + n)
            total += value
            if abs(value) < abs(total) * mp.mpf('1e-90'):
                break
        else:
            raise RuntimeError('diagnostic precision series cap')
        if mode == 'exact_series':
            z = float(aa * total)
            lp = float(aa * mp.mpf(math.log(x)) - mp.mpf(model.lg(a))
                       + mp.mpf(math.log1p(z)))
            return math.exp(lp), -math.expm1(lp)
        if mode != 'exact_except_lgamma':
            raise ValueError('unknown diagnostic arm')
        lp = aa * mp.log(xx) - mp.mpf(model.lg(a)) + mp.log1p(aa * total)
        return float(mp.exp(lp)), float(-mp.expm1(lp))


def main():
    inputs = [HERE / 'b7_screen.json', HERE / 'exact_seam_run_42470083523.txt',
              HERE / 'direct_q_diagnosis.py', HERE / 'seam_diagnosis.py',
              HERE.parent / 'model.py',
              model.ROOT / 'benchmark/beta_balanced_survival/beta_route_model.py']
    inputs += sorted((model.ROOT / 'src').glob('M_STATS_PROBDIST_*.bas'))
    report = {'status': 'DIAGNOSTIC ONLY; binary direct-Q arm REJECTED',
              'python': platform.python_version(), 'mpmath': mp.__version__,
              'sha256_lf': {str(p.relative_to(model.ROOT)): hashlib.sha256(
                  p.read_bytes().replace(b'\r\n', b'\n')).hexdigest() for p in inputs},
              'exact_excel': [], 'arms': {}}
    log = inputs[1].read_text()
    if 'TOTAL=1945;PASS=1905;FAIL=40' not in log:
        raise ValueError('unexpected Excel result')
    for line in log.splitlines():
        if 'B6 X seam 05 ' not in line:
            continue
        surface = 0 if 'Cumulative' in line else 1
        for side, raw in re.findall(r'(left|right)=([^ ]+)', line):
            # Encoding subtracts the *binary64* parse of hi in VBA.
            # Reconstruct with two binary64 parses and a binary64 addition.
            hi, lo = map(float, raw.split(';'))
            observed = hi + lo
            x = math.nextafter(1.5, 0) if side == 'left' else 1.5
            predicted = model.public(.5, x)[surface]
            report['exact_excel'].append({'surface': surface, 'side': side,
                'raw': raw, 'observed_hex': observed.hex(),
                'model_hex': predicted.hex(), 'equal': observed == predicted})
    if len(report['exact_excel']) != 4 or not all(r['equal'] for r in report['exact_excel']):
        raise ValueError('exact original-seam model comparison failed')
    revealed = json.loads(inputs[0].read_text())['rows']
    for mode in ['binary', 'exact_series', 'exact_except_lgamma', 'ideal_lower']:
        arm = {'points': 0, 'accuracy_failures': [], 'new_monotonicity_failures': [],
               'all_monotonicity_failures': []}
        previous = {}
        for row in revealed:
            a, x = float.fromhex(row['a_hex']), float.fromhex(row['x_hex'])
            if a > 1:
                continue
            if 'reference' not in row:
                raise ValueError('missing revealed reference')
            values, baseline = direct(a, x, mode), model.public(a, x)
            arm['points'] += 1
            with mp.workdps(120):
                errors = [float(abs(mp.mpf(v) / mp.mpf(ref) - 1))
                          for v, ref in zip(values, row['reference'])]
            key = (row['role'], a)
            identity = {'role': row['role'], 'index': row['index'],
                        'a_hex': a.hex(), 'x_hex': x.hex()}
            if errors[0] > 3e-15 or errors[1] > 5e-15:
                arm['accuracy_failures'].append(dict(identity, errors=errors))
            if key in previous:
                px, pv, pb = previous[key]
                if math.nextafter(px, math.inf) == x:
                    for j in range(2):
                        fails = pv[j] > values[j] if j == 0 else pv[j] < values[j]
                        base_ok = pb[j] <= baseline[j] if j == 0 else pb[j] >= baseline[j]
                        if fails:
                            item = dict(identity, surface=j, previous_x_hex=px.hex(),
                                        left_hex=pv[j].hex(), right_hex=values[j].hex())
                            arm['all_monotonicity_failures'].append(item)
                            if base_ok:
                                arm['new_monotonicity_failures'].append(item)
            previous[key] = (x, values, baseline)
        report['arms'][mode] = arm
    (HERE / 'direct_q_diagnosis.json').write_text(json.dumps(report, indent=2) + '\n')
    print(report['status'])
    for name, arm in report['arms'].items():
        print(name, arm['points'], 'accuracy failures', len(arm['accuracy_failures']),
              'new monotonicity failures', len(arm['new_monotonicity_failures']))
    # A diagnostic generated successfully is not candidate acceptance.
    rejected = any(report['arms']['binary'][key] for key in
                   ['accuracy_failures', 'new_monotonicity_failures'])
    return 1 if rejected else 0


if __name__ == '__main__':
    raise SystemExit(main())
