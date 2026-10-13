"""
Fit the #72 candidate normal-tail kernel exactly as PREREGISTRATION.md §2
fixes it, and write kernel_coefficients.json.

The rational P(z)/D(z), degree (7, 7) with D(0) = 1, approximates

    f(z) = Q(z) * exp(z^2 / 2),    0 <= z <= 3,

where Q is the standard-normal upper tail. The fit is deterministic:
linearised least squares for relative error at the 500 Chebyshev nodes of
[0, 3], weights 1 / (f * D_prev), Lawson reweighting by the node errors,
120 iterations at 60 digits, keeping the iterate with the smallest node
maximum.

Every constant is also recorded the way VBA must spell it. The VBE keeps a
numeric literal to 15 significant digits, so a constant is the Double sum
hi + lo of two 15-digit literals. hi is the target Double rounded to 15
digits, and lo is the exact remainder rounded to 15 digits. The pair is
accepted only if the sum is bit-identical to the target Double.

    python derive_kernel.py     writes kernel_coefficients.json
"""
import json
import os
import struct
import time
from fractions import Fraction

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "kernel_coefficients.json")

LOWER, UPPER = 0, 3
DEGREE_P = DEGREE_D = 7
NODES = 500
ITERATIONS = 120
DPS = 60


def target(z):
    """f(z) = Q(z) * exp(z^2/2) at the working precision."""
    z = mp.mpf(z)
    return mp.erfc(z / mp.sqrt(2)) / 2 * mp.exp(z * z / 2)


def chebyshev_nodes(a, b, n):
    a, b = mp.mpf(a), mp.mpf(b)
    return [(a + b) / 2 + (b - a) / 2 * mp.cos(mp.pi * (i + mp.mpf(1) / 2) / n)
            for i in range(n)]


def horner(coefficients, x):
    """coefficients in ascending powers."""
    acc = mp.mpf(0)
    for c in reversed(coefficients):
        acc = acc * x + c
    return acc


def fit():
    """Linearised relative least squares with Lawson reweighting."""
    xs = chebyshev_nodes(LOWER, UPPER, NODES)
    fs = [target(x) for x in xs]
    lam = [mp.mpf(1)] * NODES
    d_prev = [mp.mpf(1)] * NODES
    best = None
    for _ in range(ITERATIONS):
        rows, rhs = [], []
        for x, fx, weight, dp in zip(xs, fs, lam, d_prev):
            w = mp.sqrt(weight) / (fx * dp)
            # P(x) - f(x) * (D(x) - 1) = f(x), with D(0) = 1.
            rows.append([w * x ** i for i in range(DEGREE_P + 1)]
                        + [-w * fx * x ** j for j in range(1, DEGREE_D + 1)])
            rhs.append(w * fx)
        solution = mp.qr_solve(mp.matrix(rows), mp.matrix(rhs))[0]
        p = [solution[i] for i in range(DEGREE_P + 1)]
        d = [mp.mpf(1)] + [solution[DEGREE_P + j] for j in range(1, DEGREE_D + 1)]
        errors = [abs(horner(p, x) / horner(d, x) / fx - 1) for x, fx in zip(xs, fs)]
        worst = max(errors)
        if best is None or worst < best[0]:
            best = (worst, p, d)
        d_prev = [horner(d, x) for x in xs]
        total = sum(weight * e for weight, e in zip(lam, errors))
        lam = [weight * e / total for weight, e in zip(lam, errors)]
    return best


def double_hex(x):
    return struct.pack(">d", x).hex()


def vba_literal_pair(x):
    """(hi, lo) 15-significant-digit literals whose Double sum is exactly x."""
    hi_text = format(x, ".15g")
    remainder = Fraction(x) - Fraction(float(hi_text))
    lo_text = "0" if remainder == 0 else format(float(remainder), ".15g")
    return hi_text, lo_text, float(hi_text) + float(lo_text) == x


def constant_record(value):
    x = float(value)                       # mpmath rounds to nearest
    hi, lo, exact = vba_literal_pair(x)
    return {"value": mp.nstr(value, 40), "double": repr(x), "double_hex": double_hex(x),
            "hi": hi, "lo": lo, "pair_is_exact": exact}


def main():
    mp.mp.dps = DPS
    started = time.time()
    node_max, p, d = fit()
    inv_sqrt_two_pi = 1 / mp.sqrt(2 * mp.pi)
    record = {
        "study": "normal upper-tail kernel (#72)",
        "preregistration": "PREREGISTRATION.md",
        "target": "f(z) = Q(z) * exp(z^2/2) on [0, 3]",
        "fit": {"interval": [LOWER, UPPER], "degree": [DEGREE_P, DEGREE_D], "nodes": NODES,
                "iterations": ITERATIONS, "dps": DPS, "node_max_relative_error": mp.nstr(node_max, 6)},
        "numerator": [constant_record(c) for c in p],
        "denominator": [constant_record(c) for c in d],
        "inv_sqrt_two_pi": constant_record(inv_sqrt_two_pi),
        "split": 3.0,
        "continued_fraction_levels": 60,
        "exponent_split_scale": 256,
        "runtime_s": round(time.time() - started, 1),
    }
    with open(OUT, "w", encoding="utf-8", newline="\n") as f:
        json.dump(record, f, indent=1)
        f.write("\n")
    print(f"node max relative error {mp.nstr(node_max, 4)}; wrote {os.path.basename(OUT)} "
          f"in {record['runtime_s']} s")


if __name__ == "__main__":
    main()
