"""Supplementary research diagnostics; not outward-rounded certificates."""
import hashlib
import json
import math
from fractions import Fraction
from pathlib import Path
import mpmath as mp

HERE = Path(__file__).resolve().parent


def truncation_bound():
    # Cauchy radius 7/5, |g| <= 55; degree 72 and |a-.5| <= .5.
    return 55 * Fraction(5, 14)**73 / (1 - Fraction(5, 14))


def loggamma1p(a):
    if a < mp.mpf('.01'):
        term = a * a
        total = -mp.euler * a
        for n in range(2, 10001):
            delta = (-1)**n * mp.zeta(n) * term / n
            total += delta
            if abs(delta) < abs(total) * mp.eps:
                return total
            term *= a
        raise RuntimeError('reference loggamma1p cap')
    return mp.loggamma(1 + a)


def pinelis(a, x):
    a, x = mp.mpf(a), mp.mpf(x)
    if not mp.isfinite(a) or not mp.isfinite(x) or not 0 < a <= 1 or x <= 0:
        raise ValueError('bounds require 0<a<=1 and finite x>0')
    if a == 1:
        q = mp.exp(-x)
        return q, q
    lg = loggamma1p(a)
    b = mp.exp(lg / (a - 1))
    common = mp.exp(-x + a * mp.log(x) - lg)
    lower = common * a * (x + 3 - a) / (x*x + (4-2*a)*x + (a-1)*(a-2))
    upper = common * mp.expm1(a * mp.log1p(b/x)) / b
    if not 0 < lower < upper:
        raise RuntimeError('invalid bounds')
    return lower, upper


def diagnostic_contains(value, bounds):
    # Conversion margin only. This is NOT a certified interval implementation.
    lo, hi = bounds
    return math.nextafter(float(lo), -math.inf) <= value <= math.nextafter(float(hi), math.inf)


def main():
    points = {}
    files = ['b7_screen.json', 'precision_screen.json', 'precision_v2_screen.json']
    for name in files:
        for row in json.loads((HERE / name).read_text())['rows']:
            a = float.fromhex(row['a_hex'])
            if a <= 1:
                points[(row['a_hex'], row['x_hex'])] = row['reference'][1]
    report = {'status': 'SUPPLEMENTARY DIAGNOSTICS ONLY', 'rows': [],
              'failures': [], 'truncation_exact_rational': str(truncation_bound()),
              'truncation_upper_decimal': '1.949e-31',
              'rounding_certified': False}
    for (ah, xh), rawref in points.items():
        a, x = float.fromhex(ah), float.fromhex(xh)
        with mp.workdps(90):
            lo = pinelis(a, x)
        with mp.workdps(120):
            hi = pinelis(a, x)
            stability = max(abs(u/v-1) for u, v in zip(lo, hi))
            ref = mp.mpf(rawref)
            # References printed to 100 digits: equality at a=1 is numerical.
            inside = hi[0] <= ref <= hi[1] if a < 1 else abs(ref/hi[0]-1) < mp.mpf('1e-95')
            width = (hi[1]-hi[0])/ref
            row = {'a_hex': ah, 'x_hex': xh, 'relative_width': mp.nstr(width, 30),
                   'stability_relative': mp.nstr(stability, 20), 'reference_inside': bool(inside)}
            if stability > mp.mpf('1e-80') or not inside:
                report['failures'].append(row)
            report['rows'].append(row)
    report['sha256_lf'] = {name: hashlib.sha256((HERE/name).read_bytes().replace(b'\r\n', b'\n')).hexdigest()
                           for name in files + ['research_checks.py']}
    (HERE / 'research_checks.json').write_text(json.dumps(report, indent=2) + '\n')
    print(len(points), 'revealed points; failures', len(report['failures']))
    return 1 if report['failures'] else 0


if __name__ == '__main__':
    raise SystemExit(main())
