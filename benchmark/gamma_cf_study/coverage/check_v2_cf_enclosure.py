"""Post-hoc fraction-only audit on revealed v2 points; no acceptance claim."""
import hashlib
import json
import sys
import time
from fractions import Fraction as F
from pathlib import Path
import precision_v2 as v2
from cf_enclosure import enclose

HERE = Path(__file__).resolve().parent


def main():
    start = time.monotonic()
    report = {'status': 'INCOMPLETE', 'scope': 'fraction H only; revealed v2 set',
              'relative_error_guard': '2^-94 (diagnostic; not probability contract)',
              'enclosure_relative_width': '2^-105', 'rows': [], 'failures': []}
    inputs = ['precision_v2_protocol.json', 'precision_v2.py', 'precision_probe.py',
              'precision_coefficients.json', 'cf_enclosure.py', Path(__file__).name]
    report['sha256_lf'] = {name: hashlib.sha256((HERE/name).read_bytes().replace(b'\r\n', b'\n')).hexdigest()
                           for name in inputs}
    for point in json.loads((HERE/'precision_v2_protocol.json').read_text())['points']:
        a, x = float.fromhex(point['a_hex']), float.fromhex(point['x_hex'])
        if x < a + 1:
            continue
        observed = {}
        def capture(frame, event, arg):
            if event == 'return' and frame.f_code is v2.upper.__code__:
                observed.update(h=frame.f_locals['h'], depth=frame.f_locals['n'])
        old = sys.getprofile()
        try:
            sys.setprofile(capture)
            v2.public(a, x)
        finally:
            sys.setprofile(old)
        lo, hi, depth = enclose(a, x)
        value = F(observed['h'].hi) + F(observed['h'].lo)
        error_bound = max(abs(value-F(lo)), abs(value-F(hi))) / F(lo)
        row = dict(point, enclosure=[str(lo), str(hi)], enclosure_depth=depth,
                   lentz_depth=observed['depth'], h_hex=[observed['h'].hi.hex(), observed['h'].lo.hex()],
                   relative_error_upper_bound=float(error_bound))
        report['rows'].append(row)
        if error_bound > F(1, 2**94):
            report['failures'].append(len(report['rows'])-1)
    report['status'] = 'FAIL' if report['failures'] else 'FRACTION DIAGNOSTIC PASS; NOT P/Q ACCEPTANCE'
    report['runtime_seconds'] = time.monotonic()-start
    (HERE/'precision_v2_cf_enclosure.json').write_text(json.dumps(report, indent=2)+'\n')
    print(report['status'], len(report['rows']), 'rows', len(report['failures']), 'failures')
    print('max error bound', max(r['relative_error_upper_bound'] for r in report['rows']))
    return bool(report['failures'])


if __name__ == '__main__':
    raise SystemExit(main())
