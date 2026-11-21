"""Write the cases of the Diagnostics sheet (errors and diagnostics) with their expected outcomes.

The sheet shows what the library returns on valid input, at the edges of a support and on input it refuses,
and how to read the reason. Two parts:

  the try-it area   one function of a catalogue, its arguments typed by the user (any value: a probability of
                    1, a negative shape, a text, an empty cell), the value it returns, the kind of that value,
                    the Status text the library writes, and what both mean;
  the case table    a fixed list of calls, each with the outcome the library's documented contract gives it
                    (docs/user-manual/v1.0.0, chapters 6, 7 and 9; docs/PUBLIC_API.md, "What a function
                    returns on bad input"), the outcome the sheet observes and a check that tells an error
                    shown on purpose from an unexpected one.

Every expected outcome comes from a rule of the documentation, quoted in the case (`rule`). A number is
computed here with mpmath at 50 digits, a refused call is `#NUM!`, a text argument `#VALUE!`. Where a case
expects a Status text, `status_contains` is a fragment of the library's own message, and
tests/test_demo_diagnostics.py checks that the fragment is in the source of src/modules/: nothing on the sheet
is a diagnostic the library does not write.

Arguments are numbers, or {"text": "..."} for a text and {"empty": true} for an empty cell.

Run with mpmath installed: python examples/demo/diagnostics_references.py
"""
import json
import platform
import sys
from fractions import Fraction
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUTPUT = HERE / "diagnostics-cases.json"
DIGITS = 50
SHOWN = 17          # enough to give back the Double nearest to the reference
SWEEP_POINTS = 101

try:
    from mpmath import mp, mpf
except ImportError:  # the tests then check the stored file only
    mp = mpf = None

EMPTY = {"empty": True}

# The functions of the try-it area: the parameter names of docs/PUBLIC_API.txt, an example call, the range
# over which the chart sweeps the first argument, and the domain as chapter 6 of the manual states it.
FUNCTIONS = [
    ("K_STATS_NormalStandard_Cumulative", ["Z"], [1.5], [-10, 10],
     "Any finite Z. Far in the upper tail the result is 1: the tail is below Double resolution, use Survival."),
    ("K_STATS_NormalStandard_Survival", ["Z"], [3], [-10, 40],
     "Any finite Z. The upper tail is computed directly, never as 1 - Cumulative; beyond about Z = 38 it is below the Double range and returns 0."),
    ("K_STATS_NormalStandard_InverseCumulative", ["Probability"], [0.975], [0, 1],
     "Probability strictly between 0 and 1: 0 and 1 return #NUM!."),
    ("K_STATS_NormalStandard_IntervalProbability", ["LowerZ", "UpperZ"], [-1, 1], [-3, 3],
     "UpperZ not below LowerZ: bounds in the wrong order return #NUM!."),
    ("K_STATS_Normal_Density", ["X", "Mean", "StdDev"], [1, 0, 1], [-40, 40],
     "StdDev > 0; |X|, |Mean| and StdDev below 1E100. A density too small for a Double is 0."),
    ("K_STATS_Normal_Cumulative", ["X", "Mean", "StdDev"], [1.5, 0, 1], [-5, 5],
     "StdDev > 0; |X|, |Mean| and StdDev below 1E100. Outside these: #NUM!."),
    ("K_STATS_Normal_InverseCumulative", ["Probability", "Mean", "StdDev"], [0.975, 0, 1], [0, 1],
     "Probability strictly between 0 and 1 (0.95, not 95); StdDev > 0; |Mean| and StdDev below 1E100."),
    ("K_STATS_Lognormal_Mean", ["MeanLog", "StdDevLog"], [0.5, 0.8], [0, 800],
     "StdDevLog > 0; |MeanLog| and StdDevLog below 1E100. A mean beyond the largest Double returns #NUM!."),
    ("K_STATS_StudentT_Cumulative", ["X", "DegreesFreedom"], [1.5, 10], [-6, 6],
     "DegreesFreedom > 0 and at most 1E8; it need not be a whole number."),
    ("K_STATS_StudentT_InverseCumulative", ["Probability", "DegreesFreedom"], [0.975, 5], [0, 1],
     "Probability strictly between 0 and 1; DegreesFreedom > 0 and at most 1E8."),
    ("K_STATS_ChiSquare_Density", ["X", "DegreesFreedom"], [3, 4], [-1, 10],
     "DegreesFreedom > 0. Below 0 the density is 0; at X = 0 it is unbounded, #NUM!, when DegreesFreedom < 2."),
    ("K_STATS_Gamma_Density", ["X", "Shape", "ScaleParam"], [1, 0.5, 1], [-1, 5],
     "Shape > 0 and at most 1E20, ScaleParam > 0 (a scale, not a rate). At X = 0 unbounded, #NUM!, when Shape < 1."),
    ("K_STATS_Gamma_Cumulative", ["X", "Shape", "ScaleParam"], [3, 2, 1.5], [-2, 15],
     "Shape > 0 and ScaleParam > 0 (a scale, not a rate). Below 0 the result is 0."),
    ("K_STATS_Gamma_InverseCumulative", ["Probability", "Shape", "ScaleParam"], [0.95, 2, 1.5], [0, 1],
     "Probability strictly between 0 and 1; Shape > 0; ScaleParam > 0. A quantile below the Double range is 0."),
    ("K_STATS_Beta_Density", ["X", "Alpha", "Beta"], [0.3, 0.5, 2], [-0.5, 1.5],
     "Alpha > 0 and Beta > 0. Unbounded, #NUM!, at X = 0 when Alpha < 1 and at X = 1 when Beta < 1."),
    ("K_STATS_Beta_Cumulative", ["X", "Alpha", "Beta"], [0.3, 2, 3], [-0.5, 1.5],
     "Alpha > 0 and Beta > 0. Below 0 the result is 0 and above 1 it is 1: the natural limits, not errors."),
    ("K_STATS_Exponential_Density", ["X", "Lambda"], [2, 0.7], [-2, 10],
     "Lambda > 0 (a rate, not a scale). Below 0 the density is 0."),
    ("K_STATS_Exponential_Survival", ["X", "Lambda"], [2, 0.7], [-2, 10],
     "Lambda > 0 (a rate, not a scale). Below 0 the result is 1."),
    ("K_STATS_Uniform_Cumulative", ["X", "LowerBound", "UpperBound"], [3, 0, 10], [-5, 15],
     "LowerBound < UpperBound. Outside the bounds the result is 0 or 1."),
    ("K_STATS_Binomial_PMF", ["NumberSuccesses", "Trials", "ProbSuccess"], [3, 10, 0.3], [-2, 12],
     "Counts are truncated toward zero. NumberSuccesses from 0 to Trials; ProbSuccess from 0 to 1."),
    ("K_STATS_Binomial_Cumulative", ["NumberSuccesses", "Trials", "ProbSuccess"], [3, 10, 0.3], [-2, 12],
     "Counts are truncated toward zero. NumberSuccesses from 0 to Trials; ProbSuccess from 0 to 1; Trials up to 10,000,000."),
    ("K_STATS_Binomial_InverseCumulative", ["Probability", "Trials", "ProbSuccess"], [0.8, 10, 0.3], [0, 1],
     "Probability strictly between 0 and 1; ProbSuccess strictly between 0 and 1; Trials up to 10,000,000."),
    ("K_STATS_Poisson_PMF", ["NumberEvents", "Mean"], [3, 2.5], [-2, 12],
     "NumberEvents from 0, truncated toward zero; Mean not negative."),
    ("K_STATS_Poisson_Cumulative", ["NumberEvents", "Mean"], [3, 2.5], [-2, 12],
     "NumberEvents from 0, truncated toward zero; Mean not negative and at most 10,000,000."),
    ("K_STATS_Hypergeometric_PMF", ["SampleSuccesses", "SampleSize", "PopulationSuccesses", "PopulationSize"],
     [2, 10, 8, 40], [-2, 12],
     "SampleSize and PopulationSuccesses at most PopulationSize. A count the population makes impossible returns #NUM!."),
    ("K_STATS_DiscreteUniform_PMF", ["X", "LowerBound", "UpperBound"], [3, 0, 10], [-5, 15],
     "LowerBound not above UpperBound. Outside the support, or at a value that is not whole, the mass is 0."),
    ("K_STATS_DiscreteUniform_LogPMF", ["X", "LowerBound", "UpperBound"], [3, 0, 10], [-5, 15],
     "LowerBound not above UpperBound. Where the mass is 0 its logarithm is undefined: #NUM!."),
    ("K_STATS_DiscreteUniform_Cumulative", ["X", "LowerBound", "UpperBound"], [2.8, -3, 5], [-6, 8],
     "LowerBound not above UpperBound. Any real X: a step function, 0 below the support and 1 above it."),
]
DEFAULT_FUNCTION = "K_STATS_Gamma_InverseCumulative"

# The case table. group: valid (a number), limit (the natural limit outside or at the edge of a support),
# underflow (a true value outside the Double range, returned as its limit), invalid (refused on purpose),
# excel (what Excel does to an argument before the call). rule: where the documentation states the outcome.
NUM = "#NUM!"
VALUE = "#VALUE!"
CASES = [
    ("valid", "A valid call", "K_STATS_Normal_Cumulative", [1.5, 0, 1], None, None,
     "9.1: a call that succeeds returns a number.",
     "P(X <= 1.5) for a standard Normal.", "Nothing to correct."),
    ("limit", "Below a continuous support", "K_STATS_Beta_Cumulative", [-0.5, 2, 3], None, None,
     "7.8: continuous distributions return the natural limit and no error.",
     "The Beta lives on 0 to 1: no mass lies below -0.5.", "Nothing to correct: 0 is the answer, not an error."),
    ("limit", "Above a continuous support", "K_STATS_Uniform_Cumulative", [12, 0, 10], None, None,
     "7.8: above the support the cumulative probability is 1.",
     "All of the Uniform(0, 10) lies below 12.", "Nothing to correct."),
    ("limit", "A density outside the support", "K_STATS_Exponential_Density", [-1, 0.7], None, None,
     "7.8: below the support the density is 0.",
     "The Exponential has no mass below 0.", "Nothing to correct."),
    ("limit", "The top of a discrete support", "K_STATS_Binomial_Cumulative", [10, 10, 0.3], None, None,
     "7.8: k = n is in the Binomial support; P(X <= n) = 1.",
     "Ten successes in ten trials is the largest count.", "Nothing to correct."),
    ("limit", "A count truncated toward zero", "K_STATS_Poisson_PMF", [2.9, 3], None, None,
     "7.6: count arguments are truncated toward zero, 2.9 is read as 2.",
     "The mass of 2 events, not of 2.9.", "Pass whole counts to avoid the surprise."),
    ("limit", "A Discrete Uniform at a real X", "K_STATS_DiscreteUniform_Cumulative", [2.8, -3, 5], None, None,
     "7.6: the Discrete Uniform cumulative accepts any real X, a step function.",
     "Six of the nine integers from -3 to 5 are at most 2.8.", "Nothing to correct."),
    ("limit", "A Discrete Uniform outside its support", "K_STATS_DiscreteUniform_PMF", [20, 0, 10], None, None,
     "7.6: the Discrete Uniform PMF returns 0 outside the support.",
     "20 is not one of the integers from 0 to 10.", "Nothing to correct; its LogPMF is #NUM! (below)."),
    ("underflow", "A density too small for a Double", "K_STATS_Normal_Density", [40, 0, 1], None, None,
     "9.4: 0 for a density far in a tail, the true value is too small for a Double.",
     "The true value, about 1.5E-348, is below the smallest Double.", "Work with logarithms if the size matters."),
    ("underflow", "A cumulative probability rounded to 1", "K_STATS_NormalStandard_Cumulative", [10], None, None,
     "9.4: 1 for a cumulative probability whose upper tail is below Double resolution.",
     "The true value is 1 - 7.6E-24, stored as 1.", "For the upper tail call Survival, as in the next case."),
    ("underflow", "The same tail, kept by Survival", "K_STATS_NormalStandard_Survival", [10], None, None,
     "7.3: every distribution has a Survival function that computes the upper tail directly.",
     "P(Z > 10) about 7.6E-24, which 1 - Cumulative loses.", "Use Survival for a small upper tail."),
    ("underflow", "A Gamma quantile below the Double range", "K_STATS_Gamma_InverseCumulative", [1e-300, 0.01, 1],
     None, None,
     "9.4: 0 for a Gamma quantile whose true value is below the Double range.",
     "With Shape 0.01 the 1E-300 quantile is about 1E-30000.", "Nothing to correct: 0 is the nearest Double."),
    ("invalid", "Probability 0 in an inverse", "K_STATS_NormalStandard_InverseCumulative", [0], NUM, "Probability",
     "7.5: every inverse requires a probability strictly between 0 and 1.",
     "The 0-quantile of the Normal is minus infinity.", "Pass a probability strictly between 0 and 1."),
    ("invalid", "Probability 1 in an inverse", "K_STATS_StudentT_InverseCumulative", [1, 5], NUM, "Probability",
     "7.5: the values 0 and 1 return #NUM!.",
     "The 1-quantile of a Student t is plus infinity.", "Pass a probability strictly between 0 and 1."),
    ("invalid", "A probability typed as a percentage", "K_STATS_Normal_InverseCumulative", [95, 0, 1], NUM,
     "Probability", "7.1: enter 5 percent as 0.05, not as 5.",
     "95 is not a probability: 95% is 0.95.", "Divide by 100: pass 0.95."),
    ("invalid", "A zero standard deviation", "K_STATS_Normal_Cumulative", [1, 0, 0], NUM, "StdDev",
     "6.1: StdDev must be strictly positive.",
     "A Normal with no spread has no density.", "Pass a positive StdDev."),
    ("invalid", "A negative shape", "K_STATS_Gamma_Cumulative", [1, -2, 1.5], NUM, "Shape",
     "6.3: shape, scale and rate parameters must be strictly positive.",
     "No Gamma has a negative shape.", "Pass a positive Shape."),
    ("invalid", "Uniform bounds in the wrong order", "K_STATS_Uniform_Cumulative", [5, 10, 0], NUM, None,
     "6.3: LowerBound must be less than UpperBound.",
     "The bounds are swapped: 10 is not below 0.", "Swap the bounds: Uniform_Cumulative(5,0,10)."),
    ("invalid", "An interval with its bounds swapped", "K_STATS_NormalStandard_IntervalProbability", [1, -1], NUM,
     "UpperZ", "9.2: an interval with the upper bound below the lower bound.",
     "The upper bound -1 is below the lower bound 1.", "Pass the bounds in order: (-1, 1)."),
    ("invalid", "A negative count", "K_STATS_Poisson_Cumulative", [-1, 3], NUM, None,
     "7.8: discrete distributions are stricter, a negative count returns #NUM!.",
     "A count of events cannot be negative.", "For P(X < 0) = 0 give the answer directly, as in 7.2."),
    ("invalid", "A count above the number of trials", "K_STATS_Binomial_PMF", [12, 10, 0.3], NUM, "Trials",
     "7.8: a Binomial count above Trials returns #NUM!.",
     "Twelve successes cannot occur in ten trials.", "Keep NumberSuccesses from 0 to Trials."),
    ("invalid", "A success probability above 1", "K_STATS_Binomial_Cumulative", [12, 20, 1.35], NUM, "ProbSuccess",
     "9.5: ProbSuccess must be between 0 and 1.",
     "1.35 is not a probability.", "Pass ProbSuccess from 0 to 1."),
    ("invalid", "An unbounded density at an endpoint", "K_STATS_Beta_Density", [0, 0.5, 2], NUM, "unbounded",
     "6.3: the Beta density at X = 0 returns #NUM! when Alpha < 1.",
     "The density grows without bound as X falls to 0.", "Ask for a probability (Cumulative) instead."),
    ("invalid", "The logarithm of a zero mass", "K_STATS_DiscreteUniform_LogPMF", [20, 0, 10], NUM, "Log-mass",
     "7.6: when the outcome has probability exactly zero, LogPMF returns #NUM!.",
     "The mass of 20 is 0, and ln(0) has no finite value.", "Check the support first; the PMF gives 0."),
    ("invalid", "Degrees of freedom above the limit", "K_STATS_StudentT_Cumulative", [1, 1e9], NUM, None,
     "6.2: Student t cumulative degrees of freedom above 1E8 return #NUM!.",
     "1E9 is above the tested limit of 1E8.", "Use the Normal, which the t approaches."),
    ("invalid", "A location of 1E100 or more", "K_STATS_Normal_Cumulative", [0, 1e100, 1], NUM, "Mean",
     "9.2: location, scale or shape of size 1E100 or more, where checked.",
     "Mean = 1E100 is outside the supported magnitude.", "Rescale the problem."),
    ("invalid", "A moment beyond the largest Double", "K_STATS_Lognormal_Mean", [800, 1], NUM, "overflow",
     "9.2: a quantile or moment beyond the largest Double.",
     "exp(800.5) is above 1.8E308.", "Work with MeanLog: ln of the mean is MeanLog + StdDevLog^2 / 2."),
    ("excel", "A text argument", "K_STATS_Normal_Cumulative", [{"text": "abc"}, 0, 1], VALUE, None,
     "PUBLIC_API.md: a text returns #VALUE!; Excel refuses to convert it before the function is called.",
     "The function never runs; Excel gives #VALUE!.", "Enter a number, not text."),
    ("excel", "An empty cell read as 0", "K_STATS_Normal_Cumulative", [EMPTY, 0, 1], None, None,
     "PUBLIC_API.md: an empty cell returns what 0 returns.",
     "The empty X is read as 0, so the answer is P(X <= 0) = 0.5.", "Fill the cell: an empty input gives no error."),
    ("excel", "An empty cell as a probability", "K_STATS_NormalStandard_InverseCumulative", [EMPTY], NUM,
     "Probability", "PUBLIC_API.md: an empty cell returns what 0 returns; 7.5: a probability of 0 returns #NUM!.",
     "The empty cell is read as probability 0.", "Fill the cell with a probability."),
]

# The try-it area, as ci/Test-Diagnostics.ps1 sets it: the function, its arguments, the outcome and the
# fragment of the Status the library writes ("" for none, the call succeeded).
TRIALS = [
    ("the default example", "K_STATS_Gamma_InverseCumulative", [0.95, 2, 1.5], None, ""),
    ("probability 1", "K_STATS_Gamma_InverseCumulative", [1, 2, 1.5], NUM, "Probability"),
    ("a negative shape", "K_STATS_Gamma_InverseCumulative", [0.95, -2, 1.5], NUM, "Shape"),
    ("a text probability", "K_STATS_Gamma_InverseCumulative", [{"text": "high"}, 2, 1.5], VALUE, "Not called"),
    ("an empty X", "K_STATS_Normal_Cumulative", [EMPTY, 0, 1], None, ""),
    ("an unbounded density", "K_STATS_Beta_Density", [0, 0.5, 2], NUM, "unbounded"),
    ("an impossible count, four arguments", "K_STATS_Hypergeometric_PMF", [9, 10, 8, 40], NUM, None),
    ("a discrete quantile", "K_STATS_Binomial_InverseCumulative", [0.8, 10, 0.3], None, ""),
]


def number(value):
    return mpf(0) if value == EMPTY else mpf(value)


def binomial_pmf(k, n, p):
    p = Fraction(p)
    return Fraction(1) * __import__("math").comb(n, k) * p ** k * (1 - p) ** (n - k)


def gamma_quantile(p, shape, scale):
    """The p-quantile of a Gamma(shape, scale): bisection in log x, then the value at 50 digits."""
    p = mpf(p)
    target = lambda x: mp.gammainc(shape, 0, x, regularized=True) - p  # noqa: E731
    lo, hi = mpf(10) ** -400000, mpf(10) ** 6
    if target(lo) > 0:
        return mpf(0)
    for _ in range(4000):
        mid = mp.sqrt(lo * hi)
        if target(mid) > 0:
            hi = mid
        else:
            lo = mid
        if hi / lo - 1 < mpf(10) ** -45:
            break
    return lo * mpf(scale)


def reference(function, args):
    """The value the library should return, at 50 digits; None for a call the contract refuses."""
    a = [number(value) for value in args]
    name = function.replace("K_STATS_", "")
    if name == "Normal_Cumulative":
        return mp.ncdf((a[0] - a[1]) / a[2])
    if name == "NormalStandard_Cumulative":
        return mp.ncdf(a[0])
    if name == "NormalStandard_Survival":
        return mp.ncdf(-a[0])
    if name == "Normal_Density":
        return mp.npdf(a[0], a[1], a[2])
    if name == "Beta_Cumulative":
        x = min(max(a[0], mpf(0)), mpf(1))
        return mp.betainc(a[1], a[2], 0, x, regularized=True)
    if name == "Uniform_Cumulative":
        return min(max((a[0] - a[1]) / (a[2] - a[1]), mpf(0)), mpf(1))
    if name == "Exponential_Density":
        return mpf(0) if a[0] < 0 else a[1] * mp.exp(-a[1] * a[0])
    if name == "Binomial_Cumulative":
        k, n = int(args[0]), int(args[1])
        return mpf(sum(binomial_pmf(j, n, args[2]) for j in range(k + 1)).numerator) / \
            sum(binomial_pmf(j, n, args[2]) for j in range(k + 1)).denominator
    if name == "Binomial_InverseCumulative":
        n, total = int(args[1]), Fraction(0)
        for k in range(n + 1):
            total += binomial_pmf(k, n, args[2])
            if total >= Fraction(args[0]):
                return mpf(k)
    if name == "Poisson_PMF":
        k = int(args[0])
        return a[1] ** k * mp.exp(-a[1]) / mp.factorial(k)
    if name == "DiscreteUniform_Cumulative":
        lo, hi = int(args[1]), int(args[2])
        return mpf(sum(1 for j in range(lo, hi + 1) if j <= args[0])) / (hi - lo + 1)
    if name == "DiscreteUniform_PMF":
        lo, hi = int(args[1]), int(args[2])
        inside = args[0] == int(args[0]) and lo <= args[0] <= hi
        return mpf(1) / (hi - lo + 1) if inside else mpf(0)
    if name == "Gamma_InverseCumulative":
        return gamma_quantile(args[0], a[1], a[2])
    raise ValueError(f"no reference for {function}")


def expected(function, args, refused):
    if refused is not None:
        return {"kind": refused}
    value = reference(function, args)
    # Below the smallest positive Double (4.9E-324) the nearest Double is 0.
    stored = 0.0 if abs(value) < mpf("2.4703282292062328e-324") else float(value)
    return {"kind": "number", "value": stored, "text": mp.nstr(value, SHOWN)}


def build():
    mp.dps = DIGITS
    functions = [{"name": name, "params": params, "example": example, "sweep": sweep, "domain": domain}
                 for name, params, example, sweep, domain in FUNCTIONS]
    known = {item["name"] for item in functions}
    cases = []
    for index, (group, title, function, args, refused, fragment, rule, cause, action) in enumerate(CASES, 1):
        if function not in known:
            raise ValueError(f"{function} is not in the catalogue")
        cases.append({"id": f"c{index:02d}", "group": group, "title": title, "function": function, "args": args,
                      "expected": expected(function, args, refused), "status_contains": fragment, "rule": rule,
                      "cause": cause, "action": action})
    trials = [{"name": name, "function": function, "args": args, "expected": expected(function, args, refused),
               "status_contains": fragment} for name, function, args, refused, fragment in TRIALS]
    return {
        "schema": "diagnostics-cases-1",
        "default_function": DEFAULT_FUNCTION,
        "sweep_points": SWEEP_POINTS,
        "reference": {"method": f"the outcome each rule of the documentation gives; a number computed with mpmath "
                                f"at {DIGITS} significant digits (the Binomial by exact fractions)",
                      "sources": ["docs/user-manual/v1.0.0/06-distribution-families.md",
                                  "docs/user-manual/v1.0.0/07-conventions.md",
                                  "docs/user-manual/v1.0.0/09-errors-and-diagnostics.md",
                                  "docs/PUBLIC_API.md"],
                      "digits": SHOWN},
        "environment": {"python": platform.python_version(), "mpmath": __import__("mpmath").__version__},
        "functions": functions,
        "cases": cases,
        "trials": trials,
    }


def main():
    if mp is None:
        print("mpmath is not installed", file=sys.stderr)
        return 1
    data = build()
    OUTPUT.write_text(json.dumps(data, indent=1) + "\n", encoding="utf-8")
    print(OUTPUT, len(data["functions"]), "functions", len(data["cases"]), "cases", len(data["trials"]), "trials")
    return 0


if __name__ == "__main__":
    sys.exit(main())
