"""
Binary64 model of the balanced Beta_Survival and Beta_Cumulative route (#163).

The model ports, statement by statement, the VBA procedures that the Beta CDF
and survival reach for shapes of at least 0.5:

  K_STATS_Beta_Survival(X, a, b)    PROB_TryBetaRegularized(1 - X, X, b, a)
  K_STATS_Beta_Cumulative(X, a, b)  PROB_TryBetaRegularized(X, 1 - X, a, b)
  PROB_TryBetaRegularized           prefactor X^A Y^B / B(A, B): the literal
                                    A Log X + B Log Y - LogBeta below
                                    A + B = 1000, the Loader decomposition at
                                    and above; then the Lentz continued
                                    fraction, direct or reflected
  PROB_TryBetaContinuedFraction     modified Lentz, stop |Del - 1| <= 3E-14
  PROB_LogBeta                      half-integer shortcut, or the general
                                    LogGamma(A) + LogGamma(B) - LogGamma(A + B)
  PROB_LogGamma                     Lanczos g = 7, n = 9 (Z >= 0.5)
  PROB_LogGammaHalfDiff             direct difference or asymptotic series
  PROB_TryBetaLogPdf, PROB_TryDeviancePart, PROB_StirlingError, PROB_Log1p
                                    the Loader branch

Arithmetic semantics. VBA evaluates each expression with 64-bit-significand
(x87 extended) intermediates and rounds to a Double only when the value is
stored in a Double variable, passed ByVal As Double, or returned. Log and Exp
return Doubles. The model reproduces that with an mpmath context at 64 bits
(X) and float() at every store (D). The semantics were identified, not
assumed: on the committed Excel observations of PROB_LogGamma with Z >= 0.5,
plain binary64 evaluation reproduces 11 of 57 exactly and 64-bit
intermediates reproduce all 57 (results.json, "semantics"). Python's math.log
and math.exp stand in for VBA Log and Exp.

Constants are read from the VBA source at import, so the model cannot drift
from a transcription. The procedures are bound by SHA-256 of their text in
results.json; test_beta_route_model.py fails when one changes.

Scope. The model refuses (NotModelled) whatever it does not port: shapes
below 0.5 (the LogGamma series and reflection branches), ratios below 0.1
with the larger shape at least one (PROB_LogGammaDelta), the exact 0.1 ratio
seam (see log_beta) and the tiny-product guard of PROB_TryBetaLogPdf. That
leaves the whole balanced domain of the shapes the balanced contracts cover.

The substitutions of beta_regularized (exact_lb, lanczos, prefactor) replace
one component at a time, to attribute the error and to evaluate repair
candidates; measure_balanced_survival.py uses them. The defaults are the
source.
"""
import ast
import hashlib
import math
import operator
import os
import re

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
SOURCES = {
    "core": os.path.join(ROOT, "src", "M_STATS_PROBDIST_CORE.bas"),
    "special": os.path.join(ROOT, "src", "M_STATS_PROBDIST_SPECIALFUNCS.bas"),
    "continuous": os.path.join(ROOT, "src", "M_STATS_PROBDIST_CONTINUOUS.bas"),
}

# Procedures the model ports, by source file. Their text is hashed into
# results.json.
BOUND_PROCEDURES = (
    ("continuous", "K_STATS_Beta_Survival"),
    ("continuous", "K_STATS_Beta_Cumulative"),
    ("continuous", "K_STATS_Beta_Density"),
    ("special", "PROB_TryBetaRegularized"),
    ("special", "PROB_TryBetaContinuedFraction"),
    ("special", "PROB_LogBeta"),
    ("special", "PROB_LogGamma"),
    ("special", "PROB_LogGammaHalfDiff"),
    ("special", "PROB_TryBetaLogPdf"),
    ("special", "PROB_TryDeviancePart"),
    ("special", "PROB_StirlingError"),
    ("core", "PROB_Log1p"),
    ("core", "PROB_TryExp"),
)
CONSTANT_NAMES = (
    "PROB_HALF_LOG_TWO_PI", "PROB_HALF_LOG_PI", "PROB_NUM_EPS", "PROB_FPMIN",
    "PROB_BETA_MAX_ITER", "PROB_IBETA_LOADER_MIN_SHAPE", "PROB_HALF_DIFF_CUTOFF",
    "PROB_LOGBETA_STABLE_RATIO", "PROB_LANCZOS_G", "PROB_BD0_MAX_ITER",
    "PROB_MACH_EPS", "PROB_BLP_TINY_PRODUCT", "PROB_DENSITY_SHAPE_MAX",
) + tuple(f"PROB_LANCZOS_P{k}" for k in range(9))


class NotModelled(ValueError):
    """The arguments reach a VBA branch this model does not port."""


# --------------------------------------------------------------------------
# source reading


def _read(path):
    with open(path, encoding="utf-8") as f:
        return f.read().replace("\r\n", "\n")


def _logical_lines(text):
    """Source lines with ' _' continuations joined."""
    out, buf = [], ""
    for line in text.split("\n"):
        if line.rstrip().endswith(" _"):
            buf += line.rstrip()[:-2] + " "
            continue
        out.append(buf + line)
        buf = ""
    return out


_CONST = re.compile(r"^\s*(?:Public|Private)\s+Const\s+(\w+)\s+As\s+(Double|Long)\s*=\s*([^']+)")
_BINOPS = {ast.Add: operator.add, ast.Sub: operator.sub,
           ast.Mult: operator.mul, ast.Div: operator.truediv}


def _eval_const(expr, known, kind):
    """Evaluate a VBA constant expression in Double, as the VBE does."""
    text = re.sub(r"(\d)#", r"\1", expr.strip())

    def ev(node):
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
            return float(node.value)
        if isinstance(node, ast.Name) and node.id in known:
            return known[node.id]
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.USub):
            return -ev(node.operand)
        if isinstance(node, ast.BinOp) and type(node.op) in _BINOPS:
            return _BINOPS[type(node.op)](ev(node.left), ev(node.right))
        raise ValueError(f"unsupported constant expression {expr!r}")

    value = ev(ast.parse(text, mode="eval").body)
    return int(value) if kind == "Long" else value


def read_constants():
    """Every Double or Long constant of the three source files, by name."""
    known = {}
    for key in ("core", "special", "continuous"):
        for line in _logical_lines(_read(SOURCES[key])):
            m = _CONST.match(line)
            if m:
                try:
                    known[m.group(1)] = _eval_const(m.group(3), known, m.group(2))
                except ValueError:
                    continue
    missing = [n for n in CONSTANT_NAMES if n not in known]
    if missing:
        raise RuntimeError(f"constants not found in the VBA source: {missing}")
    return {n: known[n] for n in CONSTANT_NAMES}


def procedure_text(key, name):
    """The text of one VBA procedure, header to End Function, LF line ends."""
    text = _read(SOURCES[key])
    m = re.search(rf"^(?:Public|Private) Function {name}\(.*?^End Function$",
                  text, re.M | re.S)
    if not m:
        raise RuntimeError(f"{name} not found in {SOURCES[key]}")
    return m.group(0)


def source_binding():
    """SHA-256 of every bound procedure's text, keyed 'file:procedure'."""
    return {f"{key}:{name}": hashlib.sha256(procedure_text(key, name).encode()).hexdigest()
            for key, name in BOUND_PROCEDURES}


C = read_constants()
SOURCE_LANCZOS = tuple(C[f"PROB_LANCZOS_P{k}"] for k in range(9))

# --------------------------------------------------------------------------
# VBA arithmetic

INTERMEDIATE_BITS = 64          # x87 extended significand
_X87 = mp.MPContext()
_X87.prec = INTERMEDIATE_BITS


class intermediate_precision:
    """Context manager: evaluate intermediates at `bits` (53 = plain binary64).

    Used only to show which semantics the Excel observations follow."""

    def __init__(self, bits):
        self.bits = bits

    def __enter__(self):
        self.saved = _X87.prec
        _X87.prec = self.bits

    def __exit__(self, *exc):
        _X87.prec = self.saved


def X(v):
    """A Double or Long operand promoted to an extended intermediate."""
    return _X87.mpf(v)


def D(v):
    """Store an intermediate to a Double (round to nearest, ties to even)."""
    return float(v)


# --------------------------------------------------------------------------
# ported procedures


def log_gamma(z, lanczos=SOURCE_LANCZOS):
    """PROB_LogGamma, Lanczos branch."""
    if z < 0.5:
        raise NotModelled(f"LogGamma({z!r}): series or reflection branch")
    p = lanczos
    zm1 = D(X(z) - 1)
    x = p[0]
    for k in range(1, 9):
        x = D(X(x) + X(p[k]) / (X(zm1) + k))
    t = D(X(zm1) + X(C["PROB_LANCZOS_G"]) + X(0.5))
    return D(X(C["PROB_HALF_LOG_TWO_PI"]) + (X(zm1) + X(0.5)) * X(math.log(t))
             - X(t) + X(math.log(x)))


def log_gamma_half_diff(z, lanczos=SOURCE_LANCZOS):
    """PROB_LogGammaHalfDiff."""
    if z < C["PROB_HALF_DIFF_CUTOFF"]:
        return D(X(log_gamma(D(X(z) + X(0.5)), lanczos)) - X(log_gamma(z, lanczos)))
    w = D(1 / X(z))
    w2 = D(X(w) * X(w))
    w, w2 = X(w), X(w2)
    return D(X(0.5) * X(math.log(z)) - w / 8 + w * w2 / 192 - w * w2 * w2 / 640
             + 17 * w * w2 * w2 * w2 / 14336)


def log1p(x):
    """PROB_Log1p."""
    u = D(1 + X(x))
    if u == 1.0:
        return x
    return D(X(math.log(u)) * X(x) / (X(u) - 1))


def _stirling_tables():
    """The local constants and the half-integer table of PROB_StirlingError."""
    text = "\n".join(_logical_lines(procedure_text("special", "PROB_StirlingError")))
    series = {m.group(1): _eval_const(m.group(2), {}, "Double") for m in
              re.finditer(r"Const (S\d)\s+As Double = ([^'\n]+)", text)}
    table = {int(m.group(1)): _eval_const(m.group(2), {}, "Double") for m in
             re.finditer(r"Case (\d+): PROB_StirlingError = ([^'\n]+)", text)}
    if sorted(series) != [f"S{k}" for k in range(5)] or sorted(table) != list(range(31)):
        raise RuntimeError("PROB_StirlingError tables not found as expected")
    return series, table


_SE_SERIES, _SE_TABLE = _stirling_tables()


def stirling_error(n):
    """PROB_StirlingError."""
    s = {k: X(v) for k, v in _SE_SERIES.items()}
    if n < 0.5:
        if n <= 0:
            return 0.0
        return D(X(stirling_error(D(X(n) + 1))) + (X(n) + X(0.5))
                 * (X(log1p(n)) - X(math.log(n))) - 1)
    if n <= 15:
        two_n = D(2 * X(n))
        if two_n == math.floor(two_n) and int(two_n) in _SE_TABLE:
            return _SE_TABLE[int(two_n)]
        return D(X(stirling_error(D(X(n) + 1))) + (X(n) + X(0.5))
                 * X(log1p(D(1 / X(n)))) - 1)
    nn = D(X(n) * X(n))
    q = X(nn)
    if n > 500:
        return D((s["S0"] - s["S1"] / q) / X(n))
    if n > 80:
        return D((s["S0"] - (s["S1"] - s["S2"] / q) / q) / X(n))
    if n > 35:
        return D((s["S0"] - (s["S1"] - (s["S2"] - s["S3"] / q) / q) / q) / X(n))
    return D((s["S0"] - (s["S1"] - (s["S2"] - (s["S3"] - s["S4"] / q) / q) / q) / q) / X(n))


def deviance_part(x, mean):
    """PROB_TryDeviancePart (Loader's bd0)."""
    if x <= 0:
        return mean
    diff = D(X(x) - X(mean))
    total = D(X(x) + X(mean))
    if abs(X(diff)) < X(0.1) * X(total):
        v = D(X(diff) / X(total))
        v2 = D(X(v) * X(v))
        s = D(X(diff) * X(v))
        ej = D(2 * X(x) * X(v))
        for i in range(1, C["PROB_BD0_MAX_ITER"] + 1):
            ej = D(X(ej) * X(v2))
            term = D(X(ej) / (2 * X(i) + 1))
            new = D(X(s) + X(term))
            if new == s:
                return new
            scale = max(abs(new), 1.0)
            if abs(X(term)) <= X(C["PROB_MACH_EPS"]) * X(scale):
                return new
            s = new
        raise NotModelled("deviance series did not converge")
    r = D(X(x) * (X(math.log(x)) - X(math.log(mean))) + X(mean) - X(x))
    if r < 0:
        if abs(X(r)) <= X(C["PROB_MACH_EPS"]) * (X(x) + X(mean)):
            return 0.0
        raise NotModelled("negative deviance")
    return r


def beta_log_pdf(x, y, a, b, log_x, log_y):
    """PROB_TryBetaLogPdf (Loader decomposition of the Beta log-density)."""
    n = D(X(a) + X(b))
    mx = D(X(n) * X(x))
    my = D(X(n) * X(y))
    if mx < C["PROB_BLP_TINY_PRODUCT"] or my < C["PROB_BLP_TINY_PRODUCT"]:
        raise NotModelled("tiny deviance argument")
    dev_a = deviance_part(a, mx)
    dev_b = deviance_part(b, my)
    return D(-(X(dev_a) + X(dev_b))
             + X(0.5) * (X(math.log(a)) + X(math.log(b)) - X(math.log(n)))
             - X(log_x) - X(log_y) - X(C["PROB_HALF_LOG_TWO_PI"])
             - X(stirling_error(a)) - X(stirling_error(b)) + X(stirling_error(n)))


def log_beta(a, b, lanczos=SOURCE_LANCZOS):
    """PROB_LogBeta: half-integer shortcut or the general identity."""
    if b == 0.5 and a >= 1:
        return D(X(C["PROB_HALF_LOG_PI"]) - X(log_gamma_half_diff(a, lanczos)))
    if a == 0.5 and b >= 1:
        return D(X(C["PROB_HALF_LOG_PI"]) - X(log_gamma_half_diff(b, lanczos)))
    large, small = (a, b) if a >= b else (b, a)
    # At a decimal ratio of exactly 0.1, such as (10, 1), the quotient falls
    # on either side of the Double 0.1 depending on intermediate precision,
    # and no committed observation of the current source settles which. The
    # model refuses the seam rather than guess a branch.
    edge = C["PROB_LOGBETA_STABLE_RATIO"]
    ratio = X(small) / X(large)
    if large >= 1:
        if D(ratio) == edge and ratio != X(edge):
            raise NotModelled(f"LogBeta({a!r}, {b!r}): on the 0.1 ratio seam")
        if ratio < X(edge):
            raise NotModelled(f"LogBeta({a!r}, {b!r}): PROB_LogGammaDelta branch")
    return D(X(log_gamma(a, lanczos)) + X(log_gamma(b, lanczos))
             - X(log_gamma(D(X(a) + X(b)), lanczos)))


def beta_continued_fraction(a, b, x):
    """PROB_TryBetaContinuedFraction; returns (value, iterations)."""
    fpmin, eps = C["PROB_FPMIN"], C["PROB_NUM_EPS"]
    qab = D(X(a) + X(b))
    qap = D(X(a) + 1)
    qam = D(X(a) - 1)
    c = 1.0
    d = D(1 - X(qab) * X(x) / X(qap))
    if abs(d) < fpmin:
        d = fpmin
    d = D(1 / X(d))
    h = d
    for m in range(1, C["PROB_BETA_MAX_ITER"] + 1):
        m2 = 2 * m
        aa = D(X(m) * (X(b) - m) * X(x) / ((X(qam) + m2) * (X(a) + m2)))
        d = D(1 + X(aa) * X(d))
        if abs(d) < fpmin:
            d = fpmin
        c = D(1 + X(aa) / X(c))
        if abs(c) < fpmin:
            c = fpmin
        d = D(1 / X(d))
        h = D(X(h) * X(d) * X(c))
        aa = D(-(X(a) + m) * (X(qab) + m) * X(x) / ((X(a) + m2) * (X(qap) + m2)))
        d = D(1 + X(aa) * X(d))
        if abs(d) < fpmin:
            d = fpmin
        c = D(1 + X(aa) / X(c))
        if abs(c) < fpmin:
            c = fpmin
        d = D(1 / X(d))
        de = D(X(d) * X(c))
        h = D(X(h) * X(de))
        if abs(X(de) - 1) <= X(eps):
            return h, m
    raise NotModelled("continued fraction did not converge")


def exact_log_beta(a, b):
    """Log(Beta(a, b)) correctly rounded to a Double (component substitution)."""
    with mp.workdps(40):
        return float(mp.log(mp.beta(mp.mpf(a), mp.mpf(b))))


def beta_regularized(x, y, a, b, trace=None, lanczos=SOURCE_LANCZOS,
                     exact_lb=False, prefactor=None):
    """PROB_TryBetaRegularized; trace, when a dict, receives the parts.

    Substitutions, for attributing the error (the defaults are the source):
      exact_lb=True       PROB_LogBeta replaced by the correctly rounded
                          Log(Beta); every other statement unchanged.
      lanczos=coeffs      PROB_LogGamma with another coefficient set.
      prefactor="loader"  the Loader branch (PROB_TryBetaLogPdf) at every
                          A + B, as if PROB_IBETA_LOADER_MIN_SHAPE were 0.
    """
    if x <= 0:
        return 0.0
    if y <= 0:
        return 1.0
    loader = X(a) + X(b) >= X(C["PROB_IBETA_LOADER_MIN_SHAPE"])
    if prefactor == "loader":
        loader = True
    elif prefactor is not None:
        raise ValueError(f"unknown prefactor {prefactor!r}")
    if loader:
        lb = None
        log_bt = D(X(beta_log_pdf(x, y, a, b, math.log(x), math.log(y)))
                   + X(math.log(x)) + X(math.log(y)))
    else:
        lb = exact_log_beta(a, b) if exact_lb else log_beta(a, b, lanczos)
        log_bt = D(X(a) * X(math.log(x)) + X(b) * X(math.log(y)) - X(lb))
    try:
        bt = math.exp(log_bt)
    except OverflowError:
        raise NotModelled("prefactor overflow") from None
    if X(x) < (X(a) + 1) / (X(a) + X(b) + 2):
        h, m = beta_continued_fraction(a, b, x)
        value = D(X(bt) * X(h) / X(a))
        branch = "direct"
    else:
        h, m = beta_continued_fraction(b, a, y)
        value = D(1 - X(bt) * X(h) / X(b))
        branch = "reflected"
    value = min(max(value, 0.0), 1.0)
    if trace is not None:
        trace.update(log_beta=lb, log_bt=log_bt, bt=bt, cf=h, iterations=m,
                     branch=branch, prefactor="loader" if loader else "literal")
    return value


def beta_survival(x, a, b, trace=None, **variant):
    """K_STATS_Beta_Survival on valid arguments."""
    if x <= 0:
        return 1.0
    if x >= 1:
        return 0.0
    return beta_regularized(D(1 - X(x)), x, b, a, trace, **variant)


def beta_cumulative(x, a, b, trace=None, **variant):
    """K_STATS_Beta_Cumulative on valid arguments."""
    if x <= 0:
        return 0.0
    if x >= 1:
        return 1.0
    return beta_regularized(x, D(1 - X(x)), a, b, trace, **variant)


def beta_density(x, a, b):
    """K_STATS_Beta_Density for 0 < x < 1: the Loader log-density.

    Ported only to validate the Loader branch against the density
    observations, which reach PROB_TryBetaLogPdf at every shape."""
    if not 0 < x < 1:
        raise NotModelled("density support edge")
    if max(a, b) > C["PROB_DENSITY_SHAPE_MAX"]:
        raise NotModelled("density shape envelope")
    if min(a, b) < 0.5:
        raise NotModelled("shape below 0.5: PROB_StirlingError small-argument recursion")
    one_minus_x = D(1 - X(x))
    log_pdf = beta_log_pdf(x, one_minus_x, a, b, math.log(x), log1p(-x))
    try:
        return math.exp(log_pdf)
    except OverflowError:
        raise NotModelled("density overflow") from None
