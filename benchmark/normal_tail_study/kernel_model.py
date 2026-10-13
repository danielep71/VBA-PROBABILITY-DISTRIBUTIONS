"""
Binary64 model of the #72 candidate normal-tail kernel, of the current VBA
kernel, and of the public Normal, Lognormal and Standard Normal paths that use
them (PREREGISTRATION.md §2 and §7).

Every function here follows the VBA statement by statement. Python's float is
IEEE binary64 with round-to-nearest, as VBA's Double is. math.exp and math.log
may differ from VBA Exp and Log in the last bit; the Excel-wave tolerances in
§7 allow for that.

The current kernel is the port in
../student_t_large_df_study/binary64_route_model.py, which reproduces the 12
committed Excel normal-tail observations within 2 ulps.
"""
import json
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "student_t_large_df_study"))
import binary64_route_model as current  # noqa: E402

COEFFICIENTS_PATH = os.path.join(HERE, "kernel_coefficients.json")

TAIL_UNDERFLOW_Z = 38.5                 # PROB_TAIL_UNDERFLOW_Z
PDF_UNDERFLOW_Z = 38.6                  # PROB_PDF_UNDERFLOW_Z
Z_SATURATION = 40.0                     # PROB_NF_Z_SATURATION
BOUND = 2.0 ** -50                      # preregistered B


def _literal(record):
    """The Double VBA builds from a hi + lo literal pair."""
    return float(record["hi"]) + float(record["lo"])


def load_coefficients(path=COEFFICIENTS_PATH):
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    return {
        "numerator": [_literal(r) for r in data["numerator"]],
        "denominator": [_literal(r) for r in data["denominator"]],
        "inv_sqrt_two_pi": _literal(data["inv_sqrt_two_pi"]),
        "split": float(data["split"]),
        "levels": int(data["continued_fraction_levels"]),
        "scale": float(data["exponent_split_scale"]),
        "raw": data,
    }


COEFFICIENTS = load_coefficients() if os.path.exists(COEFFICIENTS_PATH) else None


# ---------------------------------------------------------------------------
# Candidate kernel (§2)
# ---------------------------------------------------------------------------

def exp_neg_half_square(z, k=None):
    """E(z) = exp(-z^2/2) for z >= 0, with an exact main argument."""
    k = k or COEFFICIENTS
    zh = math.floor(z * k["scale"]) / k["scale"]          # Int(Z * 256) / 256
    return math.exp(-0.5 * (zh * zh)) * math.exp(-0.5 * ((z - zh) * (z + zh)))


def candidate_upper_tail(z, k=None):
    """Q(z) for z >= 0, as the candidate PROB_NormalUpperTailPositive."""
    k = k or COEFFICIENTS
    if z >= TAIL_UNDERFLOW_Z:
        return 0.0
    e = exp_neg_half_square(z, k)
    if z < k["split"]:
        p = k["numerator"]
        d = k["denominator"]
        a = p[-1]
        for c in p[-2::-1]:
            a = a * z + c
        b = d[-1]
        for c in d[-2::-1]:
            b = b * z + c
        return e * (a / b)
    denominator = z
    for level in range(k["levels"], 0, -1):
        denominator = z + float(level) / denominator
    return e * (k["inv_sqrt_two_pi"] / denominator)


def candidate_pdf(z, k=None):
    """PROB_NormalPDF with the split exponential."""
    k = k or COEFFICIENTS
    if abs(z) >= PDF_UNDERFLOW_Z:
        return 0.0
    return k["inv_sqrt_two_pi"] * exp_neg_half_square(abs(z), k)


# ---------------------------------------------------------------------------
# Current kernel, unchanged code
# ---------------------------------------------------------------------------

current_upper_tail = current.vba_upper_tail


def current_pdf(z):
    if abs(z) >= PDF_UNDERFLOW_Z:
        return 0.0
    return current.INV_SQRT_TWO_PI * math.exp(-0.5 * z * z)


KERNELS = {
    "candidate": (candidate_upper_tail, candidate_pdf),
    "current": (current_upper_tail, current_pdf),
}


# ---------------------------------------------------------------------------
# Shared normal-family helpers, parameterised by kernel
# ---------------------------------------------------------------------------

def cdf(z, kernel="candidate"):
    q = KERNELS[kernel][0]
    return q(-z) if z < 0.0 else 1.0 - q(z)


def survival(z, kernel="candidate"):
    q = KERNELS[kernel][0]
    return q(z) if z >= 0.0 else 1.0 - q(-z)


def interval(lower_z, upper_z, kernel="candidate"):
    """PROB_NormalIntervalProbability, returning (value, branch, tail terms)."""
    q = KERNELS[kernel][0]
    if lower_z >= 0.0:
        t1, t2 = survival(lower_z, kernel), survival(upper_z, kernel)
        value, branch = t1 - t2, "upper"
    elif upper_z <= 0.0:
        t1, t2 = q(-upper_z), q(-lower_z)
        value, branch = t1 - t2, "lower"
    else:
        t1, t2 = q(-lower_z), q(upper_z)
        value, branch = 1.0 - t1 - t2, "crossing"
    value = min(max(value, 0.0), 1.0)
    return value, branch, (t1, t2)


_ACKLAM = {
    "A": (-39.6968302866538, 220.946098424521, -275.928510446969,
          138.357751867269, -30.6647980661472, 2.50662827745924),
    "B": (-54.4760987982241, 161.585836858041, -155.698979859887,
          66.8013118877197, -13.2806815528857),
    "C": (-7.78489400243029E-03, -0.322396458041136, -2.40075827716184,
          -2.54973253934373, 4.37466414146497, 2.93816398269878),
    "D": (7.78469570904146E-03, 0.32246712907004, 2.445134137143, 3.75440866190742),
}


def inv_cdf_raw(p):
    """PROB_NormalInvCDFRaw (Acklam), M_STATS_PROBDIST_CORE."""
    a1, a2, a3, a4, a5, a6 = _ACKLAM["A"]
    b1, b2, b3, b4, b5 = _ACKLAM["B"]
    c1, c2, c3, c4, c5, c6 = _ACKLAM["C"]
    d1, d2, d3, d4 = _ACKLAM["D"]
    if p < 0.02425:
        q = math.sqrt(-2.0 * math.log(p))
        return ((((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6)
                / ((((d1 * q + d2) * q + d3) * q + d4) * q + 1.0))
    if p > 0.97575:
        q = math.sqrt(-2.0 * math.log(1.0 - p))
        return -((((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6)
                 / ((((d1 * q + d2) * q + d3) * q + d4) * q + 1.0))
    q = p - 0.5
    r = q * q
    return ((((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q
            / (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1.0))


def inv_cdf(p, kernel="candidate"):
    """PROB_NormalInvCDF: the Acklam seed and one Halley step."""
    pdf = KERNELS[kernel][1]
    x = inv_cdf_raw(p)
    pdf_x = pdf(x)
    cdf_x = cdf(x, kernel)
    if pdf_x > 0.0 and cdf_x > 0.0 and cdf_x < 1.0:
        e = cdf_x - p
        u = e / pdf_x
        x = x - u / (1.0 + x * u / 2.0)
    return x


def inv_survival(p, kernel="candidate"):
    return -inv_cdf(p, kernel)


def standardize(value, location, scale):
    """PROB_TryStandardize on finite, in-range arguments."""
    return (value + (-location)) / scale


def affine(offset, scale, value):
    """PROB_TryAffineTransform on finite, in-range arguments."""
    return offset + scale * value


# ---------------------------------------------------------------------------
# Public functions affected by the repair (§7)
# ---------------------------------------------------------------------------

def _z_or_saturated(x, mean, sd):
    z = standardize(x, mean, sd)
    if math.isinf(z):
        z = math.copysign(Z_SATURATION, z)
    return z


def public(function, args, kernel="candidate"):
    """Model a K_STATS_* function on its finite binary64 arguments.

    Returns (value, z) where z is the standardised argument the kernel saw
    (None for inverse functions).
    """
    family, _, surface = function.partition("_")
    if family == "NormalStandard":
        z = args[0]
        if surface == "Survival":
            return survival(z, kernel), z
        if surface == "Cumulative":
            return cdf(z, kernel), z
        if surface == "Density":
            return KERNELS[kernel][1](z), z
        if surface == "IntervalProbability":
            return interval(args[0], args[1], kernel)[0], None
        if surface == "InverseCumulative":
            return inv_cdf(args[0], kernel), None
        if surface == "InverseSurvival":
            return inv_survival(args[0], kernel), None
    if family == "Normal":
        if surface in ("Survival", "Cumulative"):
            z = _z_or_saturated(args[0], args[1], args[2])
            return (survival(z, kernel) if surface == "Survival" else cdf(z, kernel)), z
        if surface == "IntervalProbability":
            lo = _z_or_saturated(args[0], args[2], args[3])
            hi = _z_or_saturated(args[1], args[2], args[3])
            return interval(lo, hi, kernel)[0], None
        if surface == "InverseCumulative":
            return affine(args[1], args[2], inv_cdf(args[0], kernel)), None
        if surface == "InverseSurvival":
            return affine(args[1], args[2], inv_survival(args[0], kernel)), None
    if family == "Lognormal":
        if surface in ("Survival", "Cumulative"):
            x = args[0]
            if x <= 0.0:
                return (1.0 if surface == "Survival" else 0.0), None
            z = _z_or_saturated(math.log(x), args[1], args[2])
            return (survival(z, kernel) if surface == "Survival" else cdf(z, kernel)), z
        if surface == "InverseCumulative":
            return math.exp(affine(args[1], args[2], inv_cdf(args[0], kernel))), None
        if surface == "InverseSurvival":
            return math.exp(affine(args[1], args[2], inv_survival(args[0], kernel))), None
    raise ValueError(f"not modelled: {function}")
