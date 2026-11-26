"""
#175: binary64 model of the candidate incomplete-beta prefactor.

The rule of PREREGISTRATION.md has two parts.

  Loader at every shape.  PROB_TryBetaRegularized takes the Loader form of the
      prefactor (PROB_TryBetaLogPdf) whenever both shapes are at least 1/2,
      not only from A + B = 1000. Below 1/2 the literal form stays.

  The far deviance.  PROB_TryDeviancePart, away from the series, computes
      X * Log(X / Mean) - (X - Mean) with one logarithm of a ratio, through
      Log1p((X - Mean) / Mean) while X is within a factor of two of Mean,
      where X - Mean is exact. Today it subtracts two logarithms.

Since the rule shipped, beta_balanced_survival/beta_route_model.py ports it:
`candidate` is that port as it stands, and `current` restores the dispatch and
the far deviance of before #175, so the comparison stays reproducible.

Two ports of the source evaluate both, because the semantics of VBA's
intermediates are not settled for every branch:

  "x87"     the port as it stands: every intermediate in extended precision,
            rounded on each store. It reproduced Excel on the balanced route
            before the rule (1.30E-13 at I_x(50, 100), x = 0.333333, where
            Excel measured 1.3E-13) and refuses what it does not port:
            unbalanced shapes through PROB_LogGammaDelta.
  "double"  the same port with Double intermediates and the unbalanced branch,
            as the #35 study extended it (stagnation_model).

Neither is bit-faithful to Excel everywhere; their figures bound the rule,
they do not predict observations.
"""
import importlib.util
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
BENCH = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(BENCH, "student_t_large_df_study"))
import above8_model as _above8  # noqa: E402

MIN_SHAPE = 0.5


def _load_x87():
    path = os.path.join(BENCH, "beta_balanced_survival", "beta_route_model.py")
    spec = importlib.util.spec_from_file_location("f175_forward_x87", path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


PORTS = {"double": _above8.forward, "x87": _load_x87()}
NotModelled = PORTS["x87"].NotModelled
LOADER_MIN_SHAPE = PORTS["x87"].C["PROB_IBETA_LOADER_MIN_SHAPE"]


def rule_applies(a, b):
    """Whether the candidate changes the prefactor form at shapes (a, b)."""
    return min(a, b) >= MIN_SHAPE and a + b < LOADER_MIN_SHAPE


def _deviance_before(port):
    X, D, now = port.X, port.D, port.deviance_part

    def deviance(x, mean):
        """PROB_TryDeviancePart as it was before #175: two logarithms in the far branch."""
        if x <= 0:
            return mean
        diff = D(X(x) - X(mean))
        total = D(X(x) + X(mean))
        if abs(X(diff)) < X(0.1) * X(total):
            return now(x, mean)                      # the series, unchanged
        r = D(X(x) * (X(math.log(x)) - X(math.log(mean))) + X(mean) - X(x))
        if r < 0:
            return 0.0 if abs(X(r)) <= X(port.C["PROB_MACH_EPS"]) * (X(x) + X(mean)) else r
        return r
    return deviance


def _value(result):
    return result[0] if isinstance(result, tuple) else result


def current(x, y, a, b, semantics="x87"):
    """I_x(a, b) as the source before #175 computed it."""
    port = PORTS[semantics]
    now = port.deviance_part
    port.deviance_part = _deviance_before(port)
    try:
        return _value(port.beta_regularized(x, y, a, b, prefactor="sum"))
    finally:
        port.deviance_part = now


def candidate(x, y, a, b, semantics="x87"):
    """I_x(a, b) under the rule: the ported source as it stands."""
    return _value(PORTS[semantics].beta_regularized(x, y, a, b))


def candidate_deviance(x, mean, semantics="x87"):
    return PORTS[semantics].deviance_part(x, mean)
