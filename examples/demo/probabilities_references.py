"""Regenerate the independent reference values of the Probabilities sheet with mpmath.

Same method as explorer_references.py, whose 60-digit definitions of the 16
families it reuses, and no call to the library, the builder or Excel. Each case
is one setting of the inputs of the sheet: the family and its parameters, the
points a and b and the probability p, with F(a), S(a) = P(X > a), F(b), S(b),
P(a < X <= b) = S(a) - S(b) and the quantile Q(p), the smallest x with
F(x) >= p. The interval is the difference of two survival values taken at 60
digits, so it carries no cancellation.

The cases are, for every family: the sheet's defaults (a and b at 35 % and 65 %
of the plot window, rounded down for a discrete family) with the probabilities
0.05, 0.5, 0.95; a case in each tail (a and b at the 1E-6 and 1E-3 quantiles,
and at the 0.999 and 1 - 1E-6 quantiles); a discrete family also gets a
fractional a and b, which the library truncates; and the probabilities at the
ends of the support, 1E-12 and 1 - 1E-12.

Run with mpmath installed: python examples/demo/probabilities_references.py
"""
import json
import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import explorer_references as R

try:
    from mpmath import mp, mpf
except ImportError:  # the package tests then check the stored file only
    mp = mpf = None

WINDOW_SHARES = (0.35, 0.65)


def evaluate(family, kind, params, point):
    """F and S at a point: the discrete families read the point as its integer part."""
    if kind == "discrete":
        point = math.floor(point)
        _, cdf, survival = R.discrete(family, params, point)
    else:
        _, cdf, survival = R.continuous(family, params, point)
    return cdf, survival


def default_points(family, kind, params):
    low, high = R.default_window(family, params)
    a, b = (low + share * (high - low) for share in WINDOW_SHARES)
    if kind == "discrete":
        a, b = math.floor(a), math.floor(b)
    return float(a), float(b)


def case(family, kind, params, a, b, p, label):
    f_a, s_a = evaluate(family, kind, params, a)
    f_b, s_b = evaluate(family, kind, params, b)
    q = R.quantile(family, kind, params, p) if 0 < p < 1 else None
    return {"family": family, "parameters": params, "label": label, "a": a, "b": b, "p": p,
            "cdf_a": float(f_a), "survival_a": float(s_a), "cdf_b": float(f_b), "survival_b": float(s_b),
            "interval": float(s_a - s_b), "quantile": None if q is None else float(q)}


def references():
    mp.dps = 60
    cases = []
    for family, (kind, sets) in R.FAMILIES.items():
        params = sets[0][0]
        a, b = default_points(family, kind, params)
        for p in (0.05, 0.5, 0.95):
            cases.append(case(family, kind, params, a, b, p, "default"))
        for lower_p, upper_p, label in ((1e-6, 1e-3, "lower tail"), (0.999, 1 - 1e-6, "upper tail")):
            qa = float(R.quantile(family, kind, params, lower_p))
            qb = float(R.quantile(family, kind, params, upper_p))
            if kind == "discrete":
                qa, qb = math.floor(qa), math.floor(qb)
            if qb > qa or kind == "continuous":
                cases.append(case(family, kind, params, qa, qb, lower_p, label))
        for p in (1e-12, 1 - 1e-12):
            cases.append(case(family, kind, params, a, b, p, "probability near an end"))
        if kind == "discrete":
            cases.append(case(family, kind, params, a + 0.7, b + 0.2, 0.5, "fractional points"))
    return {"method": __doc__, "absolute_tolerance": 1e-12, "relative_tolerance": 1e-9,
            "roundtrip_tolerance": 1e-8, "cases": cases}


if __name__ == "__main__":
    if mp is None:
        sys.exit("mpmath is required to regenerate the references")
    Path(__file__).with_name("probabilities-references.json").write_text(
        json.dumps(references(), indent=1) + "\n", encoding="utf-8")
