"""Write the scenarios of the Tails sheet (Extreme Tails and Numerical Limits) with their mpmath references.

Each scenario is a family at fixed parameters and a ladder of points that runs from the body of the
distribution to far beyond the smallest positive number an Excel cell holds (about 2.2E-308), so the
sheet can show, side by side, what two different things do to a small probability:

  * cancellation: 1 - CDF loses its digits and becomes exactly 0 when the true tail is about 1E-16 or
    smaller, although the number is representable; the direct survival function keeps its digits;
  * underflow: when the true tail is below the range of a Double, every method returns 0 (or an
    error), because the number cannot be stored. The log-mass functions of the discrete families
    (LogPMF) are the library's way of going past that wall; there is no log-survival function.

Two kinds of scenario:

  tail   the upper tail P(X > x) of a continuous or discrete family: the library's survival function
         against 1 - the library's cumulative function;
  log    the natural logarithm of the mass of a discrete family: the library's LogPMF against
         LN(PMF), which is an error where the mass underflows to 0.

The reference of every point is computed with mpmath at 60 digits and stored with 34: the value
itself as text (a tail below 1E-308 cannot be a Double), and its depth, -log10 of the probability
(of the mass, for a log scenario), which is a Double. The sheet grades the values of Excel against it.
A tail is never referenced by 1 - CDF.

Cross-checks that use another method are in tests/test_demo_tails.py: exact integer sums for the
Binomial, the Geometric, the Negative binomial and the Hypergeometric, a Decimal sum for the Poisson
tail, closed forms for the Exponential and Weibull, scipy's logsf where scipy is installed.

Run with mpmath installed: python examples/demo/tails_references.py
"""
import json
import platform
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUTPUT = HERE / "tails-cases.json"
DIGITS = 60
SHOWN = 34
# The smallest positive normal Double: below it an Excel cell holds 0.
SMALLEST_NORMAL = 2.2250738585072014e-308

try:
    from mpmath import mp, mpf
except ImportError:  # the tests then check the stored file only
    mp = mpf = None


def normal_tail(x):
    return mp.erfc(mpf(x) / mp.sqrt(2)) / 2


def exponential_tail(x):
    return mp.exp(-mpf(x))


def weibull_tail(x):
    return mp.exp(-mpf(x) ** 2)


def gamma_tail(x):
    return mp.gammainc(5, mpf(x), mp.inf, regularized=True)


def chisquare_tail(x):
    return mp.gammainc(mpf(3) / 2, mpf(x) / 2, mp.inf, regularized=True)


def student_tail(x):
    df = mpf(5)
    return mp.betainc(df / 2, mpf(1) / 2, 0, df / (df + mpf(x) ** 2), regularized=True) / 2


def poisson_tail(k):
    return mp.gammainc(mpf(k) + 1, 0, 10, regularized=True)


def binomial_tail(k):
    return mp.betainc(mpf(k) + 1, 2000 - mpf(k), 0, mpf(1) / 2, regularized=True)


def log_binomial(k):
    return mp.log(mp.binomial(2000, k)) + 2000 * mp.log(mpf(1) / 2)


def log_poisson(k):
    return k * mp.log(10) - 10 - mp.loggamma(mpf(k) + 1)


def log_geometric(k):
    return mp.log(mpf(1) / 2) + k * mp.log(mpf(1) / 2)


def log_negbinomial(f):
    return mp.log(mp.binomial(f + 4, f)) + 5 * mp.log(mpf(1) / 2) + f * mp.log(mpf(1) / 2)


def log_hypergeometric(k):
    return mp.log(mp.binomial(5000, k) * mp.binomial(5000, 2000 - k) / mp.binomial(10000, 2000))


# name, kind, family title, parameters text, x label, direct call, alternative path, points, reference
SCENARIOS = [
    {"name": "Standard Normal", "kind": "tail", "title": "Standard Normal upper tail, z = 1 to 40",
     "parameters": "Mean 0, standard deviation 1; z in standard deviations.", "x_label": "z",
     "direct": "K_STATS_NormalStandard_Survival({x})", "alternative": "1-K_STATS_NormalStandard_Cumulative({x})",
     "points": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 37, 38, 39, 40],
     "reference": normal_tail,
     "method": "erfc(z / sqrt(2)) / 2",
     "contract": "1E-13 in the tested region (standardized values up to 6)."},
    {"name": "Exponential", "kind": "tail", "title": "Exponential upper tail, x = 1 to 800",
     "parameters": "Rate 1 (mean 1); x in units of the mean.", "x_label": "x",
     "direct": "K_STATS_Exponential_Survival({x},1)", "alternative": "1-K_STATS_Exponential_Cumulative({x},1)",
     "points": [1, 5, 10, 20, 30, 36, 40, 50, 100, 200, 300, 400, 500, 600, 700, 710, 720, 730, 740, 745, 750, 800],
     "reference": exponential_tail,
     "method": "exp(-x)",
     "contract": "1E-13 in the tested region."},
    {"name": "Weibull", "kind": "tail", "title": "Weibull upper tail, x = 1 to 28",
     "parameters": "Shape 2, scale 1; x in units of the scale.", "x_label": "x",
     "direct": "K_STATS_Weibull_Survival({x},2,1)", "alternative": "1-K_STATS_Weibull_Cumulative({x},2,1)",
     "points": [1, 2, 3, 4, 5, 6, 7, 8, 10, 12, 15, 18, 20, 22, 24, 25, 26, 27, 28],
     "reference": weibull_tail,
     "method": "exp(-x^2)",
     "contract": "1E-13 in the tested region."},
    {"name": "Gamma", "kind": "tail", "title": "Gamma upper tail, x = 5 to 800",
     "parameters": "Shape 5, scale 1; x in units of the scale.", "x_label": "x",
     "direct": "K_STATS_Gamma_Survival({x},5,1)", "alternative": "1-K_STATS_Gamma_Cumulative({x},5,1)",
     "points": [5, 10, 20, 30, 40, 50, 60, 80, 100, 150, 200, 300, 400, 500, 600, 700, 740, 760, 780, 800],
     "reference": gamma_tail,
     "method": "regularized upper incomplete gamma Q(5, x)",
     "contract": "1E-13 in the tested region."},
    {"name": "Chi-square", "kind": "tail", "title": "Chi-square upper tail, 3 degrees of freedom, x = 1 to 1700",
     "parameters": "3 degrees of freedom.", "x_label": "x",
     "direct": "K_STATS_ChiSquare_Survival({x},3)", "alternative": "1-K_STATS_ChiSquare_Cumulative({x},3)",
     "points": [1, 5, 10, 20, 40, 60, 80, 120, 200, 300, 400, 600, 800, 1000, 1200, 1400, 1500, 1550, 1600, 1700],
     "reference": chisquare_tail,
     "method": "regularized upper incomplete gamma Q(3/2, x/2)",
     "contract": "2.6E-10 in the tested region (cumulative and survival)."},
    {"name": "Student t", "kind": "tail", "title": "Student t upper tail, 5 degrees of freedom, t = 1 to 1E70",
     "parameters": "5 degrees of freedom; t in standard-error units.", "x_label": "t",
     "direct": "K_STATS_StudentT_Survival({x},5)", "alternative": "1-K_STATS_StudentT_Cumulative({x},5)",
     "points": [1, 3, 10, 30, 100, 1e3, 1e4, 1e6, 1e8, 1e10, 1e15, 1e20, 1e30, 1e40, 1e50, 1e60, 1e62, 1e64, 1e66, 1e70],
     "reference": student_tail,
     "method": "regularized incomplete beta I(5/(5+t^2); 5/2, 1/2) / 2",
     "contract": "1.3E-12 in the tested region."},
    {"name": "Poisson", "kind": "tail", "title": "Poisson upper tail P(X > k), mean 10, k = 10 to 310",
     "parameters": "Mean 10; k a count.", "x_label": "k",
     "direct": "K_STATS_Poisson_Survival({x},10)", "alternative": "1-K_STATS_Poisson_Cumulative({x},10)",
     "points": [10, 15, 20, 30, 40, 50, 60, 80, 100, 150, 200, 250, 280, 290, 300, 310],
     "reference": poisson_tail,
     "method": "regularized lower incomplete gamma P(k + 1, 10)",
     "contract": "2E-13 in the tested region."},
    {"name": "Binomial", "kind": "tail", "title": "Binomial upper tail P(X > k), 2000 trials, p = 0.5, k = 1000 to 1999",
     "parameters": "2000 trials, success probability 0.5; k a count.", "x_label": "k",
     "direct": "K_STATS_Binomial_Survival({x},2000,0.5)", "alternative": "1-K_STATS_Binomial_Cumulative({x},2000,0.5)",
     "points": [1000, 1050, 1100, 1150, 1175, 1200, 1300, 1400, 1500, 1600, 1700, 1800, 1850, 1900, 1950, 1990, 1999],
     "reference": binomial_tail,
     "method": "regularized incomplete beta I(0.5; k + 1, 2000 - k)",
     "contract": "5E-8 in the tested region (the loosest survival contract)."},
    {"name": "Binomial LogPMF", "kind": "log", "title": "Binomial ln P(X = k), 2000 trials, p = 0.5, k = 1000 to 2000",
     "parameters": "2000 trials, success probability 0.5; k a count.", "x_label": "k",
     "direct": "K_STATS_Binomial_LogPMF({x},2000,0.5)", "alternative": "LN(K_STATS_Binomial_PMF({x},2000,0.5))",
     "points": [1000, 1100, 1200, 1400, 1600, 1700, 1800, 1850, 1900, 1950, 1980, 2000],
     "reference": log_binomial,
     "method": "ln C(2000, k) + 2000 ln(0.5)",
     "contract": "1E-13 (LogPMF) in the tested region."},
    {"name": "Poisson LogPMF", "kind": "log", "title": "Poisson ln P(X = k), mean 10, k = 10 to 800",
     "parameters": "Mean 10; k a count.", "x_label": "k",
     "direct": "K_STATS_Poisson_LogPMF({x},10)", "alternative": "LN(K_STATS_Poisson_PMF({x},10))",
     "points": [10, 50, 100, 150, 200, 250, 300, 350, 400, 500, 600, 800],
     "reference": log_poisson,
     "method": "k ln 10 - 10 - ln Gamma(k + 1)",
     "contract": "1E-13 (LogPMF) in the tested region."},
    {"name": "Geometric LogPMF", "kind": "log", "title": "Geometric ln P(X = k), p = 0.5, k = 1 to 1500",
     "parameters": "Success probability 0.5; k the number of failures before the first success.", "x_label": "k",
     "direct": "K_STATS_Geometric_LogPMF({x},0.5)", "alternative": "LN(K_STATS_Geometric_PMF({x},0.5))",
     "points": [1, 10, 100, 500, 900, 1000, 1050, 1070, 1100, 1200, 1300, 1500],
     "reference": log_geometric,
     "method": "ln p + k ln(1 - p)",
     "contract": "1E-13 (LogPMF) in the tested region."},
    {"name": "Negative Binomial LogPMF", "kind": "log", "title": "Negative binomial ln P(X = f), 5 successes, p = 0.5, f = 10 to 1500",
     "parameters": "5 successes, success probability 0.5; f the number of failures.", "x_label": "f",
     "direct": "K_STATS_NegativeBinomial_LogPMF({x},5,0.5)", "alternative": "LN(K_STATS_NegativeBinomial_PMF({x},5,0.5))",
     "points": [10, 100, 500, 800, 1000, 1050, 1100, 1200, 1300, 1500],
     "reference": log_negbinomial,
     "method": "ln C(f + 4, f) + 5 ln p + f ln(1 - p)",
     "contract": "2E-13 (LogPMF) in the tested region."},
    {"name": "Hypergeometric LogPMF", "kind": "log",
     "title": "Hypergeometric ln P(X = k), 2000 drawn, 5000 successes of 10000, k = 1000 to 2000",
     "parameters": "Sample of 2000 from a population of 10000 holding 5000 successes; k the successes in the sample.",
     "x_label": "k",
     "direct": "K_STATS_Hypergeometric_LogPMF({x},2000,5000,10000)",
     "alternative": "LN(K_STATS_Hypergeometric_PMF({x},2000,5000,10000))",
     "points": [1000, 1100, 1200, 1300, 1400, 1500, 1600, 1700, 1750, 1800, 1900, 2000],
     "reference": log_hypergeometric,
     "method": "ln[C(5000, k) C(5000, 2000 - k) / C(10000, 2000)]",
     "contract": "2E-13 (LogPMF) in the tested region."},
]


def point(scenario, x):
    value = scenario["reference"](x)
    depth = float(-mp.log10(value)) if scenario["kind"] == "tail" else float(-value / mp.log(10))
    stored = float(value)
    representable = scenario["kind"] == "log" or stored >= SMALLEST_NORMAL
    return {"x": x, "reference": mp.nstr(value, SHOWN),
            "reference_value": stored if representable else 0.0,
            "depth": depth}


def build():
    mp.dps = DIGITS
    scenarios = []
    for scenario in SCENARIOS:
        scenarios.append({
            "name": scenario["name"], "kind": scenario["kind"], "title": scenario["title"],
            "parameters": scenario["parameters"], "x_label": scenario["x_label"],
            "direct": scenario["direct"], "alternative": scenario["alternative"],
            "method": scenario["method"], "contract": scenario["contract"],
            "points": [point(scenario, x) for x in scenario["points"]],
        })
    return {
        "schema": "tails-cases-1",
        "default_scenario": "Standard Normal",
        "smallest_normal": SMALLEST_NORMAL,
        "reference": {"method": f"mpmath at {DIGITS} significant digits, stored with {SHOWN}", "digits": SHOWN},
        "environment": {"python": platform.python_version(), "mpmath": __import__("mpmath").__version__},
        "grades": {"reference": 12, "report": 6},
        "scenarios": scenarios,
    }


def main():
    if mp is None:
        print("mpmath is not installed", file=sys.stderr)
        return 1
    data = build()
    OUTPUT.write_text(json.dumps(data, indent=1) + "\n", encoding="utf-8")
    print(OUTPUT, len(data["scenarios"]), "scenarios", sum(len(s["points"]) for s in data["scenarios"]), "points")
    return 0


if __name__ == "__main__":
    sys.exit(main())
