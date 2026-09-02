"""
Faithful binary64 model of K_STATS_F_InverseCumulative with per-component
precision substitution, for the #35 isolation (PREREGISTRATION.md section 4).

Each of seven components can run in binary64 (Python float, matching VBA
Double semantics) or at 50 dps (mpmath). The substitution sweeps in section 5
flip one component at a time; the criteria in section 5 decide which
component, if any, is the responsible term. This module does the computing;
run_isolation.py does the deciding.

Components:
  C1 branch     solve for x if p <= 1-p, else for y            (VBA rule)
  C2 seed       starting point for the solver
  C3 forward    regularized incomplete beta I_u(Sa, Sb)
  C4 density    beta log-density for the Newton step
  C5 newton     the iteration itself
  C6 pair       the unsolved member, formed as 1 - U
  C7 transform  F = (d2/d1) * X / Y, in log space

Reference: _ibeta.beta_invcdf at 50 dps, transformed exactly.
"""
import math
import os
import sys

import mpmath as mp
from mpmath import mpf

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _ibeta import ibeta, beta_invcdf                                # noqa: E402

DPS = 50
PROB_EPS = 2.0 ** -52                     # relative Newton stop, as VBA
MAX_NEWTON = 200
COMPONENTS = ("C1", "C2", "C3", "C4", "C5", "C6", "C7")


def _I_hp(u, a, b):
    mp.mp.dps = DPS
    return ibeta(mpf(u), mpf(a), mpf(b))


def _I_f64(u, a, b):
    """Correctly rounded binary64 forward.

    The VBA continued fraction with its Loader prefactor was measured at
    <= 3E-16 relative on every row in the design (PREREGISTRATION.md section 1),
    i.e. at the rounding floor. Modelling it as the correctly rounded value is
    therefore faithful for THESE rows and keeps C3 from masquerading as a
    culprit through modelling noise.
    """
    return float(_I_hp(u, a, b))


def _logbeta(a, b):
    mp.mp.dps = DPS
    return mp.loggamma(mpf(a)) + mp.loggamma(mpf(b)) - mp.loggamma(mpf(a) + mpf(b))


def _logpdf_f64(u, a, b, lb):
    return (a - 1.0) * math.log(u) + (b - 1.0) * math.log1p(-u) - lb


def _logpdf_hp(u, a, b, lb):
    mp.mp.dps = DPS
    u = mpf(u)
    return (mpf(a) - 1) * mp.log(u) + (mpf(b) - 1) * mp.log1p(-u) - lb


def _seed(target, Sa, Sb, forward):
    """Bracketed bisection in the working precision of `forward`, to a loose
    tolerance; Newton finishes it. The seed cannot change the converged root,
    which is why C2 is expected to be inert."""
    lo, hi = 0.0, 1.0
    for _ in range(40):
        mid = (lo + hi) / 2
        if forward(mid, Sa, Sb) < target:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def f_inverse(p, d1, d2, prec):
    """Compute the F quantile under a per-component precision map.

    prec: dict C1..C7 -> "f64" | "hp". Returns (F, X, Y) as mpf at 50 dps so
    the caller can compare against the reference without further rounding.
    """
    hp = {c: prec[c] == "hp" for c in COMPONENTS}
    a, b = d1 / 2.0, d2 / 2.0

    # --- C1 branch (the rule is the hypothesis; both forms apply the SAME
    #     rule - "hp" here only means the comparison is done at 50 dps).
    if hp["C1"]:
        mp.mp.dps = DPS
        P, Q = mpf(p), 1 - mpf(p)
        solve_direct = P <= Q
    else:
        P, Q = p, 1.0 - p
        solve_direct = P <= Q
    if solve_direct:
        Sa, Sb, target = a, b, (P if hp["C1"] else float(P))
    else:
        Sa, Sb, target = b, a, (Q if hp["C1"] else float(Q))
    lb_hp = _logbeta(Sa, Sb)
    lb_f64 = float(lb_hp)

    # --- C3 / C4 forms
    forward = _I_hp if hp["C3"] else _I_f64
    logpdf = (lambda u: _logpdf_hp(u, Sa, Sb, lb_hp)) if hp["C4"] \
        else (lambda u: _logpdf_f64(float(u), Sa, Sb, lb_f64))

    # --- C2 seed
    if hp["C2"]:
        mp.mp.dps = DPS
        u = _seed(mpf(target), Sa, Sb, _I_hp)
    else:
        u = _seed(float(target), Sa, Sb, _I_f64)

    # --- C5 Newton
    if hp["C5"]:
        mp.mp.dps = DPS
        u = mpf(u); tgt = mpf(target)
        for _ in range(MAX_NEWTON):
            fu = forward(u, Sa, Sb) - tgt
            step = fu / mp.e ** logpdf(u)
            u_new = u - step
            if u_new <= 0 or u_new >= 1:
                u_new = u - step / 2
            if abs(u_new - u) <= mpf(PROB_EPS) * abs(u_new):
                u = u_new
                break
            u = u_new
    else:
        u = float(u); tgt = float(target)
        for _ in range(MAX_NEWTON):
            fu = float(forward(u, Sa, Sb)) - tgt
            step = fu / math.exp(float(logpdf(u)))
            u_new = u - step
            if u_new <= 0.0 or u_new >= 1.0:
                u_new = u - step / 2
            if abs(u_new - u) <= PROB_EPS * abs(u_new):
                u = u_new
                break
            u = u_new

    # --- C6 pair
    if hp["C6"]:
        mp.mp.dps = DPS
        U = mpf(u)
        other = 1 - U
    else:
        U = float(u)
        other = 1.0 - U
    X, Y = (U, other) if solve_direct else (other, U)

    # --- C7 transform
    if hp["C7"]:
        mp.mp.dps = DPS
        F = mp.e ** (mp.log(mpf(d2)) - mp.log(mpf(d1)) + mp.log(mpf(X)) - mp.log(mpf(Y)))
    else:
        F = math.exp(math.log(d2) - math.log(d1) + math.log(float(X)) - math.log(float(Y)))
    mp.mp.dps = DPS
    return mpf(F), mpf(X), mpf(Y)


def reference(p, d1, d2):
    mp.mp.dps = DPS
    x = beta_invcdf(mpf(p), mpf(d1) / 2, mpf(d2) / 2)
    y = 1 - x
    return mpf(d2) / mpf(d1) * x / y, x, y


def all_f64():
    return {c: "f64" for c in COMPONENTS}


def all_hp():
    return {c: "hp" for c in COMPONENTS}
