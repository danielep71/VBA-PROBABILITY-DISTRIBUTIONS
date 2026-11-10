"""
#307 diagnosis: what the binary64 route achieves once its deviance is exact.

The binary64 prefactor is x^A e^-x / Gamma(A+1) = exp(-stirlerr(A) - D) /
sqrt(2 pi A), with Loader's deviance D = A log(A/X) + X - A. An absolute error
in D becomes a relative error of the prefactor. This replays the map's points
from shape 0.25 in Python binary64:
  - the lower series below X = A + 1, the continued fraction from it, with
    the recurrences, stopping rules and small-shape backward re-evaluation of
    PROB_TryGammaSeriesP and PROB_TryGammaContinuedFractionQ;
  - the prefactor assembled in binary64 from D and stirlerr(A), each exact and
    rounded once ("exact D"), or D evaluated directly in binary64 ("direct D");
  - the far tail formed as 1 - other, as the pre-#242 adapter did.
It scores both tails against the governing contract, using the map's own
references. A third replay stops the continued fraction at machine epsilon
above shape 1 instead of PROB_NUM_EPS (3E-14): a second, separate change.
Exploratory: a diagnosis for the design of a routing rule, not a model of
production and not a validation.

Run: python3 analyze_exact_deviance.py
"""
import csv
import math
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import analyze_binary64_map as M  # noqa: E402

BOUNDS = {"all": (3e-15, 5e-15), "large_shape": (2e-13, 2e-13), "extreme_shape": (2e-12, 5e-13)}
MACH_EPS = 2.0 ** -52
NUM_EPS = 3e-14          # PROB_NUM_EPS: the continued fraction's stop above shape 1
FPMIN = 1e-300           # PROB_FPMIN


def series(a, x):
    """PROB_TryGammaSeriesP's recurrence: the sum starts at 1/A; stop at MACH_EPS."""
    ap, total = a, 1.0 / a
    term = total
    for _ in range(100000):
        ap += 1.0
        term = term * x / ap
        total += term
        if abs(term) <= abs(total) * MACH_EPS:
            return total
    return None


def continued_fraction(a, x, eps):
    """PROB_TryGammaContinuedFractionQ's Lentz recurrence. Above shape 1 it stops
    at eps (production: NUM_EPS); at shape 1 or below at MACH_EPS and then
    re-evaluates the same finite fraction backwards, as production does."""
    if a <= 1.0:
        eps = MACH_EPS
    b = x + 1.0 - a
    if abs(b) < FPMIN:
        b = FPMIN
    c, d = 1.0 / FPMIN, 1.0 / b
    h = d
    for i in range(1, 100001):
        an = -float(i) * (float(i) - a)
        b += 2.0
        d = an * d + b
        if abs(d) < FPMIN:
            d = FPMIN
        c = b + an / c
        if abs(c) < FPMIN:
            c = FPMIN
        d = 1.0 / d
        delta = d * c
        h = h * delta
        if abs(delta - 1.0) <= eps:
            if a <= 1.0:
                tail = 0.0
                for k in range(i, 0, -1):
                    tail = (-float(k) * (float(k) - a)) / (x + 1.0 - a + 2.0 * float(k) + tail)
                h = 1.0 / (x + 1.0 - a + tail)
            return h
    return None


def replay(exact_d, cf_eps=NUM_EPS):
    with mp.workdps(60):
        refs = {r["point_id"]: (mp.mpf(r["ref_p"]), mp.mpf(r["ref_q"]))
                for r in csv.DictReader(open(M.GRID, encoding="utf-8"))}
        out = []
        for r in csv.DictReader(open(M.GRID, encoding="utf-8")):
            a, x = float(r["shape"]), float(r["x"])
            if a < 0.25:
                continue
            A, X = mp.mpf(a), mp.mpf(x)
            d = float(A * mp.log(A / X) + X - A) if exact_d else a * math.log(a / x) + x - a
            stir = float(mp.loggamma(A + 1) - ((A + mp.mpf(0.5)) * mp.log(A) - A + mp.log(2 * mp.pi) / 2))
            # x^A e^-x / Gamma(A): the factor of both production routines
            base = a * math.exp(-stir - d) / math.sqrt(2 * math.pi * a)
            ref_p, ref_q = refs[r["point_id"]]
            if x < a + 1.0:
                s = series(a, x)
                p = None if s is None else base * s
                q = None if p is None else 1.0 - p
            else:
                s = continued_fraction(a, x, cf_eps)
                q = None if s is None else base * s
                p = None if q is None else 1.0 - q
            if p is None:
                continue
            bp, bq = BOUNDS[M.regime(a, x)]
            e_p = float(abs(mp.mpf(p) / ref_p - 1)) if ref_p > mp.mpf("1e-300") else 0.0
            e_q = float(abs(mp.mpf(q) / ref_q - 1)) if ref_q > mp.mpf("1e-300") else 0.0
            out.append((a, x, M.regime(a, x), e_p, e_q, e_p <= bp and e_q <= bq))
        return out


def main():
    for label, exact, eps in (("exact D, production stops", True, NUM_EPS),
                              ("direct D, production stops", False, NUM_EPS),
                              ("exact D, continued fraction to machine epsilon", True, MACH_EPS)):
        rows = replay(exact, eps)
        meets = sum(r[5] for r in rows)
        worst = max(max(r[3], r[4]) for r in rows)
        print(f"{label}: {meets} of {len(rows)} points from shape 0.25 meet their governing "
              f"contract on both tails; worst relative error {worst:.2e}")
        if exact:
            for a, x, reg, e_p, e_q, ok in rows:
                if not ok:
                    print(f"  A={a:<8g} X={x:<12.6g} {reg:<13} P {e_p:.2e}  Q {e_q:.2e}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
