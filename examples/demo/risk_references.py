"""Write the cases of the Risk sheet (VaR, expected shortfall, tails and default counts) with their mpmath references.

The sheet works in losses: a position of value V earns a return R over a horizon of h days, and the loss is
L = -V R, positive when money is lost. The daily return has mean mu and standard deviation sigma, and the
days are independent, so over the horizon the loss has

  mean  m = -V mu h            standard deviation  sd = V sigma sqrt(h)

as the sheet forms them in Double arithmetic. Two models share that mean and that variance:

  Normal      L = m + sd Z, Z standard normal;
  Student-t   L = m + s T, T a Student t with nu > 2 degrees of freedom and s = sd sqrt((nu - 2) / nu), so
              that Var(L) = s^2 nu / (nu - 2) = sd^2. With nu <= 2 the variance is not finite and there is no
              such scale.

For a confidence level a the value at risk is the a-quantile of the loss and the expected shortfall the mean
loss beyond it:

  Normal      VaR = m + sd z,  ES = m + sd phi(z) / (1 - a),  z the a-quantile of Z;
  Student-t   VaR = m + s t,   ES = m + s (nu + t^2) / (nu - 1) f(t) / (1 - a),  t the a-quantile of T.

The exceedance probability of a loss threshold l is P(L > l) under each model. The default count is a
Binomial: n loans that default independently with probability p each; its VaR is the a-quantile Q of the
count (the smallest k with F(k) >= a) and the loss VaR is Q times the loss per default.

A case is one set of inputs. For each, this script stores the expected value of every result cell and, for
the 101 rows of the table, the loss x = m + sd z, the two densities and the two exceedance probabilities.
The table runs from z = -4 to z = top standard deviations from the mean, top = MAX(6, INT((ES - m) / sd) + 1)
with ES the larger of the two expected shortfalls, so the far tail of a heavy Student t stays in view; the
row j has z = j (top + 4) / 100 - 4, as the sheet forms it. The references are computed with mpmath
at 50 digits from the Double values the sheet holds (m, sd, s and x formed in Double arithmetic as Excel forms
them). Cross-checks by another method (numerical integration of the expected shortfall, exact fractions for
the Binomial, SciPy where it is installed) are in tests/test_demo_risk.py.

Run with mpmath installed: python examples/demo/risk_references.py
"""
import json
import math
import platform
import sys
from fractions import Fraction
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUTPUT = HERE / "risk-cases.json"
DIGITS = 50
SHOWN = 17          # enough to give back the Double nearest to the reference
# The table: ROWS losses from FIRST_Z to top standard deviations from the mean loss, top at least MIN_TOP.
ROWS = 101
FIRST_Z = -4
MIN_TOP = 6

try:
    from mpmath import mp, mpf
except ImportError:  # the tests then check the stored file only
    mp = mpf = None

INPUT_KEYS = ("value", "mean_return", "volatility", "horizon", "confidence", "df", "threshold", "loans",
              "default_probability", "loss_per_default")
CASES = [
    {"name": "Default portfolio", "value": 1000000, "mean_return": 0.0003, "volatility": 0.01, "horizon": 10,
     "confidence": 0.99, "df": 4, "threshold": 50000, "loans": 100, "default_probability": 0.02,
     "loss_per_default": 10000,
     "purpose": "The default example: a position of 1,000,000 with a daily volatility of 1% over 10 days, "
                "99% VaR and ES, a Student t with 4 degrees of freedom; 100 loans with a default probability "
                "of 2% and a loss of 10,000 each."},
    {"name": "Expected shortfall at 97.5%", "value": 1000000, "mean_return": 0.0003, "volatility": 0.01,
     "horizon": 10, "confidence": 0.975, "df": 5, "threshold": 60000, "loans": 250,
     "default_probability": 0.01, "loss_per_default": 4000,
     "purpose": "ES at 97.5% next to the VaR at the same level: the ES is the larger of the two."},
    {"name": "Near the variance limit", "value": 1000000, "mean_return": 0.0003, "volatility": 0.01,
     "horizon": 10, "confidence": 0.999, "df": 2.5, "threshold": 100000, "loans": 1000,
     "default_probability": 0.005, "loss_per_default": 1000,
     "purpose": "nu = 2.5, just above the limit of a finite variance: the matched scale s is small and the "
                "far tail of the Student t is much heavier than the Normal's."},
    {"name": "Thin tails", "value": 1000000, "mean_return": 0.0003, "volatility": 0.01, "horizon": 10,
     "confidence": 0.99, "df": 200, "threshold": 50000, "loans": 100, "default_probability": 0.02,
     "loss_per_default": 10000,
     "purpose": "nu = 200: the Student t at matched variance is close to the Normal."},
    {"name": "One day, zero mean", "value": 2500000, "mean_return": 0, "volatility": 0.012, "horizon": 1,
     "confidence": 0.95, "df": 6, "threshold": 20000, "loans": 50, "default_probability": 0.1,
     "loss_per_default": 20000,
     "purpose": "A one-day horizon with a zero mean return: the VaR is the volatility times the quantile."},
    {"name": "Expected loss", "value": 1000000, "mean_return": -0.002, "volatility": 0.008, "horizon": 5,
     "confidence": 0.9, "df": 3, "threshold": 0, "loans": 500, "default_probability": 0.001,
     "loss_per_default": 50000,
     "purpose": "A negative mean return, so the mean loss is positive; the threshold 0 asks for the "
                "probability of any loss; the 90% default count is 1."},
    {"name": "Rare tail", "value": 5000000, "mean_return": 0.0002, "volatility": 0.015, "horizon": 20,
     "confidence": 0.9999, "df": 3, "threshold": 1000000, "loans": 10000, "default_probability": 0.0005,
     "loss_per_default": 2500,
     "purpose": "99.99% over 20 days with nu = 3: the Student-t VaR is far beyond the Normal's, and a loss "
                "of 1,000,000 is many times likelier."},
]
DEFAULT_CASE = "Default portfolio"


def doubles(case):
    """The Double values the sheet forms from the inputs: the mean and the standard deviation of the loss, the
    Student-t scale. Excel evaluates -D6*D7*D9 as ((-V) mu) h, and D22*SQRT((D11-2)/D11) as written."""
    value, mu, sigma, horizon = (float(case[key]) for key in ("value", "mean_return", "volatility", "horizon"))
    nu = float(case["df"])
    mean = (-value) * mu * horizon
    sd = value * sigma * math.sqrt(horizon)
    scale = sd * math.sqrt((nu - 2) / nu)
    return mean, sd, scale


def table_top(case, results):
    """The upper end of the table in standard deviations: MAX(6, INT((ES - m) / sd) + 1), ES the larger of the
    two expected shortfalls, in Double arithmetic from the Double results."""
    mean, sd, _ = doubles(case)
    shortfall = max(float(results["normal_es"]), float(results["t_es"]))
    return max(MIN_TOP, math.floor((shortfall - mean) / sd) + 1)


def table_x(case, top):
    """The losses of the table: z = j (top - FIRST_Z) / (ROWS - 1) + FIRST_Z and x = m + sd z, j = 0..ROWS - 1,
    in Double arithmetic as Excel evaluates SEQUENCE(101,1,0,1)*(top+4)/100-4."""
    mean, sd, _ = doubles(case)
    rows = []
    for j in range(ROWS):
        z = j * (top - FIRST_Z) / (ROWS - 1) + FIRST_Z
        rows.append((z, mean + sd * z))
    return rows


def t_survival(u, nu):
    """P(T > u) for a Student t with nu degrees of freedom, from the regularized incomplete beta."""
    half = mp.betainc(nu / 2, mpf(1) / 2, 0, nu / (nu + u * u), regularized=True) / 2
    return half if u >= 0 else 1 - half


def t_density(u, nu):
    return (mp.gamma((nu + 1) / 2) / (mp.sqrt(nu * mp.pi) * mp.gamma(nu / 2))
            * (1 + u * u / nu) ** (-(nu + 1) / 2))


def t_quantile(level, nu):
    """The level-quantile of T, for level > 1/2: the root of P(T > t) = 1 - level."""
    target = 1 - level
    low, high = mpf(0), mpf(1)
    while t_survival(high, nu) > target:
        high *= 2
    for _ in range(400):
        middle = (low + high) / 2
        if t_survival(middle, nu) > target:
            low = middle
        else:
            high = middle
    return mp.findroot(lambda t: t_survival(t, nu) - target, (low + high) / 2)


def normal_quantile(level):
    return mp.sqrt(2) * mp.erfinv(2 * level - 1)


def binomial_quantile(n, p, level):
    """The smallest k with F(k) >= level, by exact fractions of the Double p and the Double level."""
    p, level = Fraction(float(p)), Fraction(float(level))
    a, b = p.numerator, p.denominator - p.numerator
    whole = p.denominator ** n
    numerator, below = b ** n, 0
    for k in range(n + 1):
        below += numerator
        if below * level.denominator >= level.numerator * whole:
            return k, Fraction(whole - below, whole)
        numerator = numerator * (n - k) * a // ((k + 1) * b)
    raise ValueError("no quantile")


def references(case):
    mean_d, sd_d, scale_d = doubles(case)
    mean, sd, scale = mpf(mean_d), mpf(sd_d), mpf(scale_d)
    level = mpf(float(case["confidence"]))
    nu = mpf(float(case["df"]))
    threshold = mpf(float(case["threshold"]))
    z = normal_quantile(level)
    t = t_quantile(level, nu)
    tail = 1 - level
    quantile, beyond = binomial_quantile(case["loans"], case["default_probability"], case["confidence"])
    return {
        "mean_loss": mean, "sd_loss": sd,
        "normal_var": mean + sd * z,
        "normal_es": mean + sd * mp.npdf(z) / tail,
        "t_scale": scale,
        "t_var": mean + scale * t,
        "t_es": mean + scale * (nu + t * t) / (nu - 1) * t_density(t, nu) / tail,
        "normal_exceed": 1 - mp.ncdf((threshold - mean) / sd) if threshold < mean else mp.ncdf(-(threshold - mean) / sd),
        "t_exceed": t_survival((threshold - mean) / scale, nu),
        "expected_defaults": mpf(case["loans"]) * mpf(float(case["default_probability"])),
        "default_quantile": mpf(quantile),
        "default_var": mpf(quantile) * mpf(float(case["loss_per_default"])),
        "default_exceed": mpf(beyond.numerator) / beyond.denominator,
    }


def text(value):
    if value == 0 or value == int(value):
        return str(int(value))
    return mp.nstr(value, SHOWN, strip_zeros=False)


def build_case(case):
    results = references(case)
    _, _, scale_d = doubles(case)
    mean, sd, scale, nu = mpf(doubles(case)[0]), mpf(doubles(case)[1]), mpf(scale_d), mpf(float(case["df"]))
    table = []
    top = table_top(case, results)
    for z, x in table_x(case, top):
        u = (mpf(x) - mean) / scale
        normal_u = (mpf(x) - mean) / sd
        table.append({
            "z": z, "x": x,
            "normal_density": float(mp.npdf(normal_u) / sd),
            "t_density": float(t_density(u, nu) / scale),
            "normal_exceed": float(mp.ncdf(-normal_u)),
            "t_exceed": float(t_survival(u, nu)),
        })
    return {
        "name": case["name"], "purpose": case["purpose"],
        "inputs": {key: case[key] for key in INPUT_KEYS},
        "results": {key: float(value) for key, value in results.items()},
        "results_text": {key: text(value) for key, value in results.items()},
        "top": top,
        "table": table,
    }


def build():
    mp.dps = DIGITS
    return {
        "schema": "risk-cases-1",
        "default_case": DEFAULT_CASE,
        "reference": {"method": f"mpmath at {DIGITS} significant digits from the Double values the sheet holds; "
                                "the Student t from the regularized incomplete beta, its quantile by bisection "
                                "and Newton; the Binomial quantile by exact fractions",
                      "digits": SHOWN},
        "environment": {"python": platform.python_version(), "mpmath": __import__("mpmath").__version__},
        "table": {"rows": ROWS, "first_z": FIRST_Z, "min_top": MIN_TOP},
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
