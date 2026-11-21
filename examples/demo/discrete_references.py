"""Write the cases of the Discrete sheet (Binomial and Poisson counts) with their mpmath references.

The sheet computes, for one model at a time, the mass at a threshold k, the four ways of reading a
threshold (at most k, fewer than k, at least k, more than k), the quantile of a probability q and the
two cumulative values that frame it, F(Q - 1) < q <= F(Q). Below it, a table gives the same five
quantities for every count of a window around the mean.

A case is one set of inputs. For each, this script stores the expected value of every result cell and,
for every count of the window and a few counts beyond it, the mass, the cumulative value and the
survival value. The references are computed with mpmath at 50 digits from the Double inputs the sheet
holds (a probability such as 0.02 is the Double nearest to it; a Poisson mean is the Double product of
the rate and the exposure, as Excel forms it):

  Binomial  P(X = j) = C(n, j) p^j (1 - p)^(n - j); the cumulative and survival values are sums of
            those masses, each over its own side (a survival value is never 1 - CDF);
  Poisson   P(X = j) = exp(-m) m^j / j!; the cumulative value is the regularized upper incomplete
            gamma Q(j + 1, m) and the survival value the lower one, P(j + 1, m).

The quantile is the smallest k with F(k) >= q, found on the exact cumulative values. Cross-checks by
another method (exact rationals for the Binomial, a Decimal series for the Poisson, scipy where it is
installed) are in tests/test_demo_discrete.py.

Run with mpmath installed: python examples/demo/discrete_references.py
"""
import json
import math
import platform
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUTPUT = HERE / "discrete-cases.json"
DIGITS = 50
SHOWN = 17          # enough to give back the Double nearest to the reference
# The window of the table: the counts from mean - SPREAD * sd to mean + SPREAD * sd, widened to hold
# the threshold and the quantile, cut to the support. MARGIN more counts on each side are stored so the
# check does not depend on a rounding at the ends of the window.
SPREAD = 5
MARGIN = 3
MAX_ROWS = 1000

try:
    from mpmath import mp, mpf
except ImportError:  # the tests then check the stored file only
    mp = mpf = None

# name, model, n, p, rate, exposure, k, q, purpose
CASES = [
    {"name": "Loan defaults, Binomial", "model": "Binomial", "trials": 100, "probability": 0.02,
     "rate": 0.02, "exposure": 100, "threshold": 5, "level": 0.99,
     "purpose": "The default example: 100 independent loans, each defaulting with probability 2%."},
    {"name": "Loan defaults, Poisson", "model": "Poisson", "trials": 100, "probability": 0.02,
     "rate": 0.02, "exposure": 100, "threshold": 5, "level": 0.99,
     "purpose": "The same portfolio as a Poisson count: 0.02 defaults per loan-year over 100 loan-years."},
    {"name": "Fair coin", "model": "Binomial", "trials": 20, "probability": 0.5,
     "rate": 0.02, "exposure": 100, "threshold": 10, "level": 0.5,
     "purpose": "A symmetric Binomial: the median is the quantile of 0.5."},
    {"name": "Step boundary", "model": "Binomial", "trials": 2, "probability": 0.5,
     "rate": 0.02, "exposure": 100, "threshold": 1, "level": 0.75,
     "purpose": "q equals F(1) = 0.75 exactly: the smallest k with F(k) >= q is 1, not 2."},
    {"name": "Threshold zero", "model": "Poisson", "trials": 100, "probability": 0.02,
     "rate": 1.5, "exposure": 2, "threshold": 0, "level": 0.05,
     "purpose": "k = 0: fewer than 0 is impossible and at least 0 is certain; F(0) = exp(-3) is just below q."},
    {"name": "Threshold at n", "model": "Binomial", "trials": 10, "probability": 0.3,
     "rate": 0.02, "exposure": 100, "threshold": 10, "level": 0.999,
     "purpose": "k = n: at most n is certain and more than n impossible."},
    {"name": "Large mean", "model": "Poisson", "trials": 100, "probability": 0.02,
     "rate": 5, "exposure": 100, "threshold": 550, "level": 0.999,
     "purpose": "A Poisson mean of 500: the window is about 220 counts wide."},
    {"name": "Rare events", "model": "Binomial", "trials": 5000, "probability": 0.001,
     "rate": 0.02, "exposure": 100, "threshold": 12, "level": 0.9999,
     "purpose": "5000 trials with p = 0.001: the Binomial is close to a Poisson of mean 5."},
    {"name": "Far quantile", "model": "Poisson", "trials": 100, "probability": 0.02,
     "rate": 0.1, "exposure": 5, "threshold": 3, "level": 0.999999,
     "purpose": "A mean of 0.5 and q = 0.999999: the quantile lies beyond mean + 5 standard deviations."},
]
DEFAULT_CASE = "Loan defaults, Binomial"


def model_mean_sd(case):
    """The mean and the standard deviation as the sheet forms them for the window, in Double arithmetic."""
    if case["model"] == "Binomial":
        n, p = float(case["trials"]), float(case["probability"])
        return n * p, math.sqrt(n * p * (1 - p))
    mean = float(case["rate"]) * float(case["exposure"])
    return mean, math.sqrt(mean)


MASSES = {}


def masses(case):
    """Every Binomial mass 0..n, computed once per case: C(n, j) p^j (1 - p)^(n - j)."""
    if case["name"] not in MASSES:
        n, p = case["trials"], mpf(float(case["probability"]))
        MASSES[case["name"]] = [mp.binomial(n, j) * p ** j * (1 - p) ** (n - j) for j in range(n + 1)]
    return MASSES[case["name"]]


def mass(case, j):
    if case["model"] == "Binomial":
        if j < 0 or j > case["trials"]:
            return mpf(0)
        return masses(case)[j]
    m = mpf(float(case["rate"]) * float(case["exposure"]))
    if j < 0:
        return mpf(0)
    return mp.exp(-m) * m ** j / mp.factorial(j)


def cumulative(case, j):
    """P(X <= j), from the lower side."""
    if j < 0:
        return mpf(0)
    if case["model"] == "Binomial":
        if j >= case["trials"]:
            return mpf(1)
        return mp.fsum(masses(case)[: j + 1])
    m = mpf(float(case["rate"]) * float(case["exposure"]))
    return mp.gammainc(j + 1, m, mp.inf, regularized=True)


def survival(case, j):
    """P(X > j), from the upper side: never 1 - CDF."""
    if j < 0:
        return mpf(1)
    if case["model"] == "Binomial":
        if j >= case["trials"]:
            return mpf(0)
        return mp.fsum(masses(case)[j + 1:])
    m = mpf(float(case["rate"]) * float(case["exposure"]))
    return mp.gammainc(j + 1, 0, m, regularized=True)


def quantile(case):
    """The smallest k with F(k) >= q, q being the Double the sheet holds."""
    q = mpf(float(case["level"]))
    mean, sd = model_mean_sd(case)
    j = max(0, int(mean) - 1)
    while j > 0 and cumulative(case, j - 1) >= q:
        j -= 1
    while cumulative(case, j) < q:
        j += 1
    return j


def window(case, quantile_value):
    """The counts of the table, as the sheet's names DiscreteLower and DiscreteUpper define them."""
    mean, sd = model_mean_sd(case)
    k = case["threshold"]
    lower = max(0, min(k, quantile_value, math.floor(mean - SPREAD * sd)))
    upper = max(k, quantile_value, math.ceil(mean + SPREAD * sd))
    if case["model"] == "Binomial":
        upper = min(case["trials"], upper)
    return lower, upper


def text(value):
    if value == 0 or value == int(value):
        return str(int(value))
    return mp.nstr(value, SHOWN, strip_zeros=False)


def number(value):
    return float(value)


def build_case(case):
    k = case["threshold"]
    q_index = quantile(case)
    lower, upper = window(case, q_index)
    if upper - lower + 1 > MAX_ROWS:
        raise ValueError(f'{case["name"]}: the window has {upper - lower + 1} rows')
    if case["model"] == "Binomial":
        n, p = case["trials"], mpf(float(case["probability"]))
        mean, sd = n * p, mp.sqrt(n * p * (1 - p))
    else:
        mean = mpf(float(case["rate"]) * float(case["exposure"]))
        sd = mp.sqrt(mean)
    results = {
        "mean": mean, "sd": sd, "mass": mass(case, k), "at_most": cumulative(case, k),
        "fewer_than": cumulative(case, k - 1), "at_least": survival(case, k - 1), "more_than": survival(case, k),
        "quantile": mpf(q_index), "before_quantile": cumulative(case, q_index - 1),
        "at_quantile": cumulative(case, q_index),
    }
    last = upper + MARGIN
    if case["model"] == "Binomial":
        last = min(last, case["trials"])
    table = [{"k": j, "mass": number(mass(case, j)), "at_most": number(cumulative(case, j)),
              "more_than": number(survival(case, j))}
             for j in range(max(0, lower - MARGIN), last + 1)]
    return {
        "name": case["name"], "model": case["model"], "purpose": case["purpose"],
        "inputs": {"trials": case["trials"], "probability": case["probability"], "rate": case["rate"],
                   "exposure": case["exposure"], "threshold": k, "level": case["level"]},
        "results": {key: number(value) for key, value in results.items()},
        "results_text": {key: text(value) for key, value in results.items()},
        "window": {"lower": lower, "upper": upper},
        "table": table,
    }


def build():
    mp.dps = DIGITS
    MASSES.clear()
    return {
        "schema": "discrete-cases-1",
        "default_case": DEFAULT_CASE,
        "reference": {"method": f"mpmath at {DIGITS} significant digits from the Double inputs; "
                                "survival values summed or integrated on their own side",
                      "digits": SHOWN},
        "environment": {"python": platform.python_version(), "mpmath": __import__("mpmath").__version__},
        "window": {"spread": SPREAD, "margin": MARGIN, "max_rows": MAX_ROWS},
        "cases": [build_case(case) for case in CASES],
    }


def main():
    if mp is None:
        print("mpmath is not installed", file=sys.stderr)
        return 1
    data = build()
    OUTPUT.write_text(json.dumps(data, indent=1) + "\n", encoding="utf-8")
    print(OUTPUT, len(data["cases"]), "cases", sum(len(case["table"]) for case in data["cases"]), "table rows")
    return 0


if __name__ == "__main__":
    sys.exit(main())
