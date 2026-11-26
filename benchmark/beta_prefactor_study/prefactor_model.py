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

Two ports of the source evaluate it, because the semantics of VBA's
intermediates are not settled for every branch:

  "x87"     beta_balanced_survival/beta_route_model.py as it stands: every
            intermediate in extended precision, rounded on each store. It
            reproduces Excel on the balanced route (1.30E-13 at I_x(50, 100),
            x = 0.333333, where Excel measures 1.3E-13) and refuses what it
            does not port: unbalanced shapes through PROB_LogGammaDelta.
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


def _candidate_deviance(port):
    X, D, literal = port.X, port.D, port.deviance_part

    def deviance(x, mean):
        """PROB_TryDeviancePart with the candidate's far branch."""
        if x <= 0:
            return mean
        diff = D(X(x) - X(mean))
        total = D(X(x) + X(mean))
        if abs(X(diff)) < X(0.1) * X(total):
            return literal(x, mean)
        if 0.5 * mean <= x <= 2.0 * mean:
            log_ratio = port.log1p(D(X(diff) / X(mean)))
        else:
            log_ratio = math.log(D(X(x) / X(mean)))
        return D(X(x) * X(log_ratio) - X(diff))
    return deviance


def _value(result):
    return result[0] if isinstance(result, tuple) else result


def current(x, y, a, b, semantics="x87"):
    """I_x(a, b) as the source of today computes it."""
    return _value(PORTS[semantics].beta_regularized(x, y, a, b))


def candidate(x, y, a, b, semantics="x87"):
    """I_x(a, b) under the rule."""
    port = PORTS[semantics]
    literal = port.deviance_part
    port.deviance_part = _candidate_deviance(port)
    try:
        if min(a, b) >= MIN_SHAPE:
            return _value(port.beta_regularized(x, y, a, b, prefactor="loader"))
        return _value(port.beta_regularized(x, y, a, b))
    finally:
        port.deviance_part = literal


def candidate_deviance(x, mean, semantics="x87"):
    return _candidate_deviance(PORTS[semantics])(x, mean)
