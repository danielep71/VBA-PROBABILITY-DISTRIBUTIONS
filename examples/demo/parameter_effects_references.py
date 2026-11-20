"""Regenerate the independent reference values of the Parameter Effects sheet with mpmath.

Same method as explorer_references.py, whose 60-digit definitions of the 16 families it reuses,
and no call to the library, the builder or Excel. Each case is one setting of the sheet: a
family at its default base parameters, the parameter to vary and its four values. For each of the
four curves it holds the mean, standard deviation and skewness, and P(X > mean + 3 sd).

The moments are not the closed forms of the sheet. A continuous family's raw moments are
integrated numerically from its density (mpmath quad over the support, split at several points),
a discrete family's are summed from its mass function over the support (or, for the unbounded
supports, until the tail is below 1E-40); the mean, variance and skewness follow from the raw
moments. Whether a moment exists is a rule of the distribution, stated here: the Student t has a
mean above 1 degree of freedom, a finite variance above 2 and a skewness above 3; the F has a mean
above 2 denominator degrees of freedom, a finite variance above 4 and a skewness above 6; a
variance between the two thresholds is infinite, below the mean's threshold it is undefined.
A moment that does not exist is "undefined" or "infinite", as the sheet shows it.

The shared x axis is the union of the four plot windows (explorer_references.default_window at
each curve's parameters), as in the sheet; five points of the grid carry the density and the
cumulative function of each curve, so a test can check the spilled columns. A discrete family's
tail is taken at the integer part of mean + 3 sd, as the library reads a point.

The cases are, for every family and every parameter: the default values of the four curves; and,
for the parameter that is varied first, four values between them (halfway from each default to
the next, the first kept), as a user would type.

Run with mpmath installed: python examples/demo/parameter_effects_references.py
"""
import json
import math
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import explorer_references as R

try:
    from mpmath import mp, mpf
except ImportError:  # the package tests then check the stored file only
    mp = mpf = None

SAMPLE_INDEXES = (0, 50, 100, 150, 200)
RAW_TAIL = mpf(10) ** -40 if mpf else None


def exists(family, params):
    """Which moments exist, as a rule of the distribution: (mean, variance, skewness), each one of
    True, 'infinite' (the variance only) or 'undefined'."""
    if family == "Student t":
        v = params[0]
        return ("undefined" if v <= 1 else True,
                True if v > 2 else ("infinite" if v > 1 else "undefined"),
                True if v > 3 else "undefined")
    if family == "F":
        d2 = params[1]
        return ("undefined" if d2 <= 2 else True,
                True if d2 > 4 else ("infinite" if d2 > 2 else "undefined"),
                True if d2 > 6 else "undefined")
    return (True, True, True)


def support(family, params):
    if family in ("Normal", "Student t"):
        return (-mp.inf, mp.inf)
    if family == "Beta":
        return (mpf(0), mpf(1))
    if family == "Uniform":
        return (mpf(params[0]), mpf(params[1]))
    return (mpf(0), mp.inf)


def pieces(lo, hi, center):
    """Split points for the integration: the finite ends, and decades around a centre."""
    if lo == -mp.inf:
        cuts = [-mp.inf, center - 1000, center - 100, center - 10, center - 1, center, center + 1, center + 10,
                center + 100, center + 1000, mp.inf]
    elif hi == mp.inf:
        cuts = [lo, lo + 1, lo + 10, lo + 100, lo + 1000, lo + 10000, mp.inf]
    else:
        cuts = [lo, (lo + hi) / 2, hi]
    return cuts


def raw_moments(family, kind, params):
    p = [mpf(value) for value in params]
    if kind == "continuous":
        lo, hi = support(family, params)
        cuts = pieces(lo, hi, p[0] if family == "Normal" else mpf(0))

        def density(x):
            return R.continuous(family, params, x)[0]
        return [mpf(1)] + [mp.quad(lambda x, n=n: x ** n * density(x), cuts) for n in (1, 2, 3)]
    total = [mpf(0)] * 4
    k = 0
    while True:
        mass = R.discrete(family, params, k)[0]
        for n in range(4):
            total[n] += mpf(k) ** n * mass
        k += 1
        if family == "Binomial" and k > params[0]:
            break
        if family == "Hypergeometric" and k > min(params[0], params[1]):
            break
        if family == "Discrete uniform" and k > params[1]:
            break
        if family not in ("Binomial", "Hypergeometric", "Discrete uniform"):
            if k > 20 and R.discrete(family, params, k)[2] < RAW_TAIL:
                break
    return total


def moments(family, kind, params):
    """Mean, standard deviation, skewness and the tail P(X > mean + 3 sd), as numbers or text."""
    rules = exists(family, params)
    if family == "Discrete uniform":                       # the support starts at the lower bound
        low, high = int(params[0]), int(params[1])
        masses = [(k, R.discrete(family, params, k)[0]) for k in range(low, high + 1)]
        m = [sum(mpf(k) ** n * mass for k, mass in masses) for n in range(4)]
    else:
        m = raw_moments(family, kind, params)
    mean = m[1] / m[0]
    variance = m[2] / m[0] - mean ** 2
    sd = mp.sqrt(variance)
    skew = (m[3] / m[0] - 3 * mean * m[2] / m[0] + 2 * mean ** 3) / sd ** 3
    out = {"mean": float(mean) if rules[0] is True else rules[0]}
    out["sd"] = float(sd) if rules[1] is True else rules[1]
    out["skewness"] = float(skew) if rules[2] is True else rules[2]
    if rules[0] is True and rules[1] is True:
        point = mean + 3 * sd
        if kind == "discrete":
            # A point that is an integer up to the rounding of the sums (Poisson(1): 1 + 3 = 4) is that
            # integer, as it is in the sheet's double arithmetic.
            point = int(mp.floor(mpf(mp.nstr(point, 12))))
        evaluate = R.continuous if kind == "continuous" else R.discrete
        top = support(family, params)[1] if kind == "continuous" else None
        if top is not None and point >= top:
            out["tail"] = 0.0                              # beyond the end of a bounded support
        else:
            out["tail"] = float(evaluate(family, params, point)[2])
    else:
        out["tail"] = "n/a"
    return out


def curve_parameters(base, slot, value):
    values = list(base)
    values[slot] = value
    return values


def case(family, kind, base, slot, values, label):
    curves = []
    windows = []
    for value in values:
        params = curve_parameters(base, slot, value)
        curves.append({"value": value, **moments(family, kind, params)})
        windows.append(R.default_window(family, params))
    low, high = min(w[0] for w in windows), max(w[1] for w in windows)
    evaluate = R.continuous if kind == "continuous" else R.discrete
    if kind == "continuous":
        count = 201
        indexes = list(SAMPLE_INDEXES)
        xs = [low + (high - low) * index / 200 for index in indexes]
    else:
        first = math.ceil(low)
        count = int(high) - first + 1
        indexes = [round(position * (count - 1) / 4) for position in range(5)]
        xs = [first + index for index in indexes]
    sample = []
    for index, x in zip(indexes, xs):
        row = {"index": index, "x": float(x), "density": [], "cdf": []}
        for value in values:
            params = curve_parameters(base, slot, value)
            try:
                density, cdf, _ = evaluate(family, params, x)
                row["density"].append(float(density))
                row["cdf"].append(float(cdf))
            except (ValueError, ZeroDivisionError):
                row["density"].append(None)
                row["cdf"].append(None)
        sample.append(row)
    return {"family": family, "label": label, "base": base, "vary": slot + 1, "values": values,
            "lower": float(low), "upper": float(high), "rows": count, "curves": curves, "sample": sample}


def between(values, integer):
    """Four values a user might type instead: the first kept, then halfway to each next default."""
    out = [values[0]]
    for earlier, later in zip(values, values[1:]):
        middle = (earlier + later) / 2
        out.append(math.floor(middle) if integer else middle)
    return out


def references():
    mp.dps = 30
    effects = json.loads((HERE / "parameter-effects.json").read_text(encoding="utf-8"))["families"]
    families = {f["name"]: f for f in json.loads((HERE / "explorer-families.json").read_text(encoding="utf-8"))["families"]}
    cases = []
    for name, (kind, _) in R.FAMILIES.items():
        family = families[name]
        base = [p["default"] for p in family["parameters"]]
        entry = effects[name]
        for slot, parameter in enumerate(entry["parameters"]):
            cases.append(case(name, kind, base, slot, parameter["curves"], "default values"))
        slot = entry["vary"] - 1
        values = between(entry["parameters"][slot]["curves"], family["parameters"][slot]["integer"])
        if values != entry["parameters"][slot]["curves"] and len(set(values)) == 4:
            cases.append(case(name, kind, base, slot, values, "typed values"))
    return {"method": __doc__, "absolute_tolerance": 1e-12, "relative_tolerance": 1e-9, "cases": cases}


if __name__ == "__main__":
    if mpmath_missing := (mp is None):
        sys.exit("mpmath is required: pip install mpmath")
    data = references()
    (HERE / "parameter-effects-references.json").write_text(json.dumps(data, indent=1) + "\n", encoding="utf-8")
    print(len(data["cases"]), "cases")
