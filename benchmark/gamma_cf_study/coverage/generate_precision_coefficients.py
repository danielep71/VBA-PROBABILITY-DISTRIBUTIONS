"""Generate two-part degree-72 Taylor coefficients of logGamma(1+a)/(a(a-1))."""
import json
from pathlib import Path
import mpmath as mp


def generate(dps):
    with mp.workdps(dps):
        coefficients = []
        for n in range(73):
            if n == 0:
                ell = mp.loggamma(mp.mpf('1.5'))
            elif n == 1:
                ell = mp.digamma(mp.mpf('1.5'))
            else:
                ell = (-1)**n * mp.zeta(n, mp.mpf('1.5')) / n
            c = -4 * ell + (4 * coefficients[n-2] if n >= 2 else 0)
            coefficients.append(c)
        return [[float(c).hex(), float(c - mp.mpf(float(c))).hex()] for c in coefficients]


if __name__ == '__main__':
    lo, hi = generate(100), generate(160)
    if lo != hi:
        raise RuntimeError('coefficient precision disagreement')
    Path(__file__).with_name('precision_coefficients.json').write_text(
        json.dumps({'degree': 72, 'precisions_dps': [100, 160],
                    'coefficients_hi_lo_hex': hi}, indent=2) + '\n')
