"""Evaluate frozen two-Double feasibility design; never tune its candidate."""
import hashlib
import json
import math
import platform
import time
from pathlib import Path

import mpmath as mp
import precision_v2 as dd
import seam_diagnosis as model
from screen_b7 import reference

HERE = Path(__file__).resolve().parent
ROOT = model.ROOT


def digest(path):
    return hashlib.sha256(path.read_bytes().replace(b'\r\n', b'\n')).hexdigest()


def main():
    start = time.monotonic()
    protocol = json.loads((HERE / 'precision_v2_protocol.json').read_text())
    for name, expected in protocol['sha256_lf'].items():
        if digest(ROOT / name) != expected:
            raise ValueError('frozen input mismatch: ' + name)
    report = {'status': 'INCOMPLETE', 'protocol_commit':
              '7519fd2111e8c511c1cdf4fe5aeacfbec748d056',
              'python': platform.python_version(), 'mpmath': mp.__version__,
              'rows': [], 'failures': [], 'monotonicity_failures': [],
              'input_hashes': protocol['sha256_lf'],
              'screen_sha256': digest(Path(__file__)),
              'protocol_sha256': digest(HERE / 'precision_v2_protocol.json'),
              'reference_code_sha256': digest(HERE / 'screen_b7.py')}
    previous = {}
    for i, point in enumerate(protocol['points']):
        a, x = float.fromhex(point['a_hex']), float.fromhex(point['x_hex'])
        row = dict(point, index=i)
        try:
            values = dd.public(a, x)
            if not all(math.isfinite(v) and 0 <= v <= 1 for v in values):
                raise ValueError('invalid probability')
            row['values_hex'] = [v.hex() for v in values]
            refs, agreement = reference(a, x)
            with mp.workdps(120):
                errors = [float(abs(mp.mpf(v) / r - 1)) for v, r in zip(values, refs)]
                row['reference'] = [mp.nstr(r, 100) for r in refs]
            row['relative_errors'] = errors
            row['quadrature_disagreement'] = agreement
            if errors[0] > protocol['thresholds']['P_relative'] or errors[1] > protocol['thresholds']['Q_relative']:
                report['failures'].append({'index': i, 'reason': 'accuracy'})
            if a in previous:
                px, pv = previous[a]
                if math.nextafter(px, math.inf) == x:
                    for j in range(2):
                        if pv[j] > values[j] if j == 0 else pv[j] < values[j]:
                            report['monotonicity_failures'].append(
                                {'index': i, 'surface': j, 'previous_x_hex': px.hex()})
            previous[a] = (x, values)
        except (ValueError, RuntimeError, OverflowError, ZeroDivisionError) as exc:
            row['error'] = str(exc)
            report['failures'].append({'index': i, 'reason': str(exc)})
            previous.pop(a, None)
        report['rows'].append(row)
        if (i + 1) % 22 == 0:
            print('evaluated', i + 1, flush=True)
    for name, expected in protocol['sha256_lf'].items():
        if digest(ROOT / name) != expected:
            raise ValueError('input changed during evaluation: ' + name)
    left = dd.public(.5, math.nextafter(1.5, 0))
    right = dd.public(.5, 1.5)
    report['original_seam_fixed'] = left[0] <= right[0] and left[1] >= right[1]
    reject = report['failures'] or report['monotonicity_failures'] or not report['original_seam_fixed']
    report['status'] = 'REJECT' if reject else 'FEASIBILITY PASS; NOT VBA OR FULL-DOMAIN ACCEPTANCE'
    report['runtime_seconds'] = time.monotonic() - start
    (HERE / 'precision_v2_screen.json').write_text(json.dumps(report, indent=2) + '\n')
    print(report['status'], 'failures', len(report['failures']),
          'monotonicity', len(report['monotonicity_failures']))
    return 1 if reject else 0


if __name__ == '__main__':
    raise SystemExit(main())
