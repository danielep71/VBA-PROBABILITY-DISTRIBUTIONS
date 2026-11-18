"""Regenerate the independent Explorer reference values with mpmath.

Every value is computed from the definition of the distribution at 60 digits,
by formulas written here, with no call to the library, the builder or Excel:
densities from the closed forms, continuous cumulative functions from the
regularized incomplete gamma and beta functions, discrete ones from the finite
sums or the closed forms. Survival is P(X > x), taken from the complement
function when there is one, so a tail is not computed as 1 - CDF.

Each case carries the parameters in the order of the Explorer's four input
cells and the probe point. Each family also carries its default plot window,
written out here, and five points of its default grid, so a test can check the
spilled x column and the density and CDF columns. This fixture checks the demo's
wiring and arithmetic; it does not replace the library's accuracy contracts.

Run with mpmath installed: python examples/demo/explorer_references.py
"""
import json
import sys
from pathlib import Path

try:
    import mpmath
    from mpmath import mp, mpf
except ImportError:  # the package tests then check the stored file only
    mpmath = None

DEFAULT_POINTS = 201
SAMPLE_INDEXES = (0, 50, 100, 150, 200)


def continuous(family, params, x):
    """Density, cumulative and survival of a continuous family at x."""
    x = mpf(x)
    p = [mpf(value) for value in params]
    if family == "Normal":
        z = (x - p[0]) / p[1]
        return (mp.npdf(x, p[0], p[1]), mp.erfc(-z / mp.sqrt(2)) / 2, mp.erfc(z / mp.sqrt(2)) / 2)
    if family == "Lognormal":
        if x <= 0:
            return (mpf(0), mpf(0), mpf(1))
        z = (mp.log(x) - p[0]) / p[1]
        return (mp.npdf(mp.log(x), p[0], p[1]) / x, mp.erfc(-z / mp.sqrt(2)) / 2, mp.erfc(z / mp.sqrt(2)) / 2)
    if family == "Student t":
        v = p[0]
        density = mp.gamma((v + 1) / 2) / (mp.sqrt(v * mp.pi) * mp.gamma(v / 2)) * (1 + x * x / v) ** (-(v + 1) / 2)

        def lower_tail(t):  # P(T <= t) for t <= 0
            return mp.betainc(v / 2, mpf(1) / 2, 0, v / (v + t * t), regularized=True) / 2
        return (density, lower_tail(x) if x <= 0 else 1 - lower_tail(-x),
                lower_tail(-x) if x >= 0 else 1 - lower_tail(x))
    if family == "Chi-square":
        k = p[0]
        density = x ** (k / 2 - 1) * mp.exp(-x / 2) / (2 ** (k / 2) * mp.gamma(k / 2))
        return (density, mp.gammainc(k / 2, 0, x / 2, regularized=True),
                mp.gammainc(k / 2, x / 2, mp.inf, regularized=True))
    if family == "F":
        d1, d2 = p[0], p[1]
        density = (d1 ** (d1 / 2) * d2 ** (d2 / 2) * x ** (d1 / 2 - 1)
                   / ((d2 + d1 * x) ** ((d1 + d2) / 2) * mp.beta(d1 / 2, d2 / 2)))
        return (density, mp.betainc(d1 / 2, d2 / 2, 0, d1 * x / (d1 * x + d2), regularized=True),
                mp.betainc(d2 / 2, d1 / 2, 0, d2 / (d1 * x + d2), regularized=True))
    if family == "Gamma":
        k, scale = p[0], p[1]
        density = x ** (k - 1) * mp.exp(-x / scale) / (mp.gamma(k) * scale ** k)
        return (density, mp.gammainc(k, 0, x / scale, regularized=True),
                mp.gammainc(k, x / scale, mp.inf, regularized=True))
    if family == "Beta":
        a, b = p[0], p[1]
        density = x ** (a - 1) * (1 - x) ** (b - 1) / mp.beta(a, b)
        return (density, mp.betainc(a, b, 0, x, regularized=True), mp.betainc(b, a, 0, 1 - x, regularized=True))
    if family == "Exponential":
        return (p[0] * mp.exp(-p[0] * x), -mp.expm1(-p[0] * x), mp.exp(-p[0] * x))
    if family == "Weibull":
        k, scale = p[0], p[1]
        u = (x / scale) ** k
        return (k / scale * (x / scale) ** (k - 1) * mp.exp(-u), -mp.expm1(-u), mp.exp(-u))
    if family == "Uniform":
        low, high = p[0], p[1]
        inside = low <= x <= high
        return (1 / (high - low) if inside else mpf(0), min(max((x - low) / (high - low), 0), 1),
                min(max((high - x) / (high - low), 0), 1))
    raise KeyError(family)


def discrete(family, params, k):
    """Mass, cumulative and survival of a discrete family at the integer k."""
    p = [mpf(value) for value in params]
    if family == "Binomial":
        n, q = int(p[0]), p[1]

        def mass(j):
            return mp.binomial(n, j) * q ** j * (1 - q) ** (n - j)
        cdf = mp.fsum(mass(j) for j in range(0, k + 1))
        return (mass(k), cdf, mp.fsum(mass(j) for j in range(k + 1, n + 1)))
    if family == "Poisson":
        mean = p[0]
        cdf = mp.gammainc(k + 1, mean, mp.inf, regularized=True)
        return (mp.exp(-mean) * mean ** k / mp.factorial(k), cdf, mp.gammainc(k + 1, 0, mean, regularized=True))
    if family == "Geometric":
        q = p[0]
        return (q * (1 - q) ** k, 1 - (1 - q) ** (k + 1), (1 - q) ** (k + 1))
    if family == "Negative binomial":
        r, q = int(p[0]), p[1]

        def mass(j):
            return mp.binomial(j + r - 1, j) * q ** r * (1 - q) ** j
        cdf = mp.betainc(r, k + 1, 0, q, regularized=True)
        return (mass(k), cdf, mp.betainc(k + 1, r, 0, 1 - q, regularized=True))
    if family == "Hypergeometric":
        n, successes, size = int(p[0]), int(p[1]), int(p[2])

        def mass(j):
            return mp.binomial(successes, j) * mp.binomial(size - successes, n - j) / mp.binomial(size, n)
        low, high = max(0, n - (size - successes)), min(n, successes)
        return (mass(k) if low <= k <= high else mpf(0),
                mp.fsum(mass(j) for j in range(low, min(k, high) + 1)),
                mp.fsum(mass(j) for j in range(max(k + 1, low), high + 1)))
    if family == "Discrete uniform":
        low, high = int(p[0]), int(p[1])
        count = high - low + 1
        return (mpf(1) / count if low <= k <= high else mpf(0), mpf(min(max(k - low + 1, 0), count)) / count,
                mpf(min(max(high - k, 0), count)) / count)
    raise KeyError(family)


# family: (kind, [(parameters, [probe points]), ...], default window as a function of the defaults)
FAMILIES = {
    "Normal": ("continuous", [([0, 1], [-3, -1, 0, 0.5, 2]), ([2, 3], [-4, 2, 5, 11]), ([-5, 0.25], [-5.5, -5, -4.75])]),
    "Lognormal": ("continuous", [([0, 0.5], [0.2, 1, 2, 5]), ([1, 1], [0.5, 2.718281828459045, 20])]),
    "Student t": ("continuous", [([5, ], [-4, -1, 0, 1, 3]), ([1, ], [-10, 0.5, 6]), ([30, ], [-2, 2.5]), ([2.5, ], [-0.3, 7])]),
    "Chi-square": ("continuous", [([4, ], [0.5, 3, 8, 15]), ([1, ], [0.1, 1, 4]), ([30, ], [20, 30, 50])]),
    "F": ("continuous", [([5, 10], [0.3, 1, 3, 8]), ([1, 1], [0.5, 2]), ([20, 4], [0.7, 2.5, 6])]),
    "Gamma": ("continuous", [([2, 1], [0.5, 2, 6]), ([0.5, 2], [0.2, 1, 5]), ([20, 0.5], [6, 10, 15])]),
    "Beta": ("continuous", [([2, 5], [0.05, 0.3, 0.8]), ([0.5, 0.5], [0.1, 0.5, 0.9]), ([30, 40], [0.3, 0.43, 0.6])]),
    "Exponential": ("continuous", [([1], [0, 0.5, 3]), ([0.25], [1, 8, 30])]),
    "Weibull": ("continuous", [([1.5, 1], [0.2, 1, 3]), ([0.7, 2], [0.5, 4, 12]), ([5, 3], [2, 3, 4])]),
    "Uniform": ("continuous", [([0, 1], [-0.5, 0, 0.5, 1, 2]), ([-2, 6], [-3, 0, 5.9])]),
    "Binomial": ("discrete", [([20, 0.4], [0, 3, 8, 12, 20]), ([1, 0.5], [0, 1]), ([100, 0.02], [0, 2, 6])]),
    "Poisson": ("discrete", [([4], [0, 3, 9]), ([0.5], [0, 1, 4]), ([60], [40, 60, 80])]),
    "Geometric": ("discrete", [([0.3], [0, 2, 10]), ([0.05], [0, 20, 60]), ([1], [0, 1])]),
    "Negative binomial": ("discrete", [([3, 0.4], [0, 4, 12]), ([1, 0.5], [0, 3]), ([10, 0.7], [2, 4, 9])]),
    "Hypergeometric": ("discrete", [([10, 30, 100], [0, 3, 7, 10]), ([5, 5, 10], [5]), ([20, 50, 200], [2, 5, 9])]),
    "Discrete uniform": ("discrete", [([1, 6], [0, 1, 3, 6, 7]), ([-3, 4], [-4, -3, 0, 4])]),
}


def default_window(family, params):
    """Plot window the Explorer proposes for the default parameters."""
    a = params
    if family == "Normal":
        return (a[0] - 4 * a[1], a[0] + 4 * a[1])
    if family == "Lognormal":
        return (0, float(mp.exp(a[0] + 3 * a[1])))
    if family == "Student t":
        return (-6, 6)
    if family == "Chi-square":
        return (0, a[0] + 5 * float(mp.sqrt(2 * a[0])) + 2)
    if family == "F":
        return (0, 5)
    if family == "Gamma":
        return (0, (a[0] + 5 * float(mp.sqrt(a[0]))) * a[1])
    if family == "Beta":
        return (0, 1)
    if family == "Exponential":
        return (0, 8 / a[0])
    if family == "Weibull":
        return (0, a[1] * float(mp.log(1000)) ** (1 / a[0]))
    if family == "Uniform":
        return (a[0] - 0.1 * (a[1] - a[0]), a[1] + 0.1 * (a[1] - a[0]))
    if family == "Binomial":
        return (0, a[0])
    if family == "Poisson":
        return (0, int(-(-(a[0] + 8 * float(mp.sqrt(a[0])) + 10) // 1)))
    if family == "Geometric":
        return (0, 1 if a[0] >= 1 else min(500, int(-(-float(mp.log(0.001) / mp.log(1 - a[0])) // 1))))
    if family == "Negative binomial":
        failures = a[0] * (1 - a[1]) / a[1] + 6 * float(mp.sqrt(a[0] * (1 - a[1]))) / a[1] + 10
        return (0, min(1000, int(-(-failures // 1))))
    if family == "Hypergeometric":
        return (max(0, a[0] - (a[2] - a[1])), min(a[0], a[1]))
    if family == "Discrete uniform":
        return (a[0], a[1])
    raise KeyError(family)


def number(value):
    return float(value)


def references():
    mp.dps = 60
    cases, grids = [], []
    for family, (kind, sets) in FAMILIES.items():
        evaluate = continuous if kind == "continuous" else discrete
        for params, points in sets:
            for point in points:
                density, cdf, survival = evaluate(family, params, point)
                cases.append({"family": family, "parameters": params, "x": point, "density": number(density),
                              "cdf": number(cdf), "survival": number(survival)})
        defaults = sets[0][0]
        low, high = default_window(family, defaults)
        if kind == "continuous":
            indexes = list(SAMPLE_INDEXES)
            xs = [low + (high - low) * index / (DEFAULT_POINTS - 1) for index in indexes]
            count = DEFAULT_POINTS
        else:
            count = int(high) - int(low) + 1
            indexes = [round(slot * (count - 1) / 4) for slot in range(5)]
            xs = [int(low) + index for index in indexes]
        sample = []
        for index, x in zip(indexes, xs):
            try:
                density, cdf, _ = evaluate(family, defaults, x)
                sample.append({"index": index, "x": number(x), "density": number(density), "cdf": number(cdf)})
            except (ValueError, ZeroDivisionError):
                sample.append({"index": index, "x": number(x), "density": None, "cdf": None})
        grids.append({"family": family, "parameters": defaults, "lower": number(low), "upper": number(high),
                      "rows": count, "sample": sample})
    return {"method": __doc__, "absolute_tolerance": 1e-12, "relative_tolerance": 1e-9,
            "cases": cases, "grids": grids}


if __name__ == "__main__":
    if mpmath is None:
        sys.exit("mpmath is required to regenerate the references")
    Path(__file__).with_name("explorer-references.json").write_text(
        json.dumps(references(), indent=1) + "\n", encoding="utf-8")
