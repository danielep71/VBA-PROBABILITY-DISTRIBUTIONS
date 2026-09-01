"""
Validate the derived Student-t large-df tail expansion against the frozen
design in PREREGISTRATION.md, and apply its crossover decision rule.

This script does not amend the design. It evaluates:

  * E_K(df, t)  relative truncation error of the order-K expansion vs the
                100-dps primary oracle, for K = 1..8, at both 60 and 100 dps
  * C(df, t)    relative error of a faithful binary64 Lentz continued
                fraction with the VBA tolerance (PROB_NUM_EPS = 3E-14)
  * oracle stability (60 vs 100 dps), and the algorithm-independent
                quadrature leg
  * the decision rule in PREREGISTRATION.md section 8, mechanically
  * holdout (forward and inverse) ONCE, after (K*, DF_MIN, T_MAX) are fixed
  * every seam in section 7, below / at / above

The 559-row benchmark/holdout/holdout_grid.csv is never read.

Run: python3 validate_expansion.py     -> results.json
"""
import json
import os
import platform
import sys
import time

import mpmath
from mpmath import mp, mpf

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
from _ibeta import t_cdf, ibeta                                   # noqa: E402

# ---------------------------------------------------------------------------
# Frozen design (PREREGISTRATION.md). Values are exact binary64.
# ---------------------------------------------------------------------------
FIT_DF = (1e3, 1e4, 1e5, 1e6, 1e7, 1e8)
FIT_T = (0.5, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0, 6.0, 8.0)
HOLD_DF = (3e3, 3e4, 3e5, 3e6, 3e7)
HOLD_T = (0.75, 1.25, 1.75, 2.5, 3.5, 4.5, 7.0)

INV_DF = (1e6, 1e7, 1e8)
INV_FIT_OFFSETS = (2.0 ** -53, 2.0 ** -43, 2.0 ** -33, 2.0 ** -23)
INV_FIT_BRIDGE = (0.625, 0.75, 0.875)
INV_FIT_ENVELOPE = (0.9, 0.99)
INV_HOLD_OFFSETS = (2.0 ** -48, 2.0 ** -38, 2.0 ** -28, 2.0 ** -24)
INV_HOLD_BRIDGE = (0.5625, 0.6875, 0.8125, 0.9375)

DPS_LOW, DPS_HIGH = 60, 100
BOUND = mpf(2) ** -50
MAX_ORDER = 8
DF_MIN_CANDIDATES = (1e5, 1e6)
T_MAX_CANDIDATES = (4.0, 6.0, 8.0)
SEAM_OFFSET = mpf("1e-9")
STABILITY_DIGITS = 50

# ---------------------------------------------------------------------------
# Coefficients: loaded from the derivation output, never from this file.
# ---------------------------------------------------------------------------
with open(os.path.join(HERE, "coefficients.json"), encoding="utf-8") as f:
    COEF = json.load(f)


def g_poly(k, t):
    """Evaluate g_k(t) from the exact rational terms."""
    s = mpf(0)
    for n, p, q in COEF["g"][str(k)]["terms"]:
        s += mpf(p) / mpf(q) * t ** n
    return s


def expansion_tail(t, df, K):
    """S(t; df) = Q(t) + phi(t) * sum_{k=1..K} g_k(t) / df^k, upper tail."""
    t, df = mpf(t), mpf(df)
    Q = mp.erfc(t / mp.sqrt(2)) / 2
    phi = mp.e ** (-t * t / 2) / mp.sqrt(2 * mp.pi)
    s = mpf(0)
    for k in range(1, K + 1):
        s += g_poly(k, t) / df ** k
    return Q + phi * s


def oracle_tail(t, df):
    """Primary: upper tail DIRECTLY from the incomplete beta.

    S(t) = I_z(df/2, 1/2) / 2 with z = df/(df + t^2), for t >= 0. This is the
    preregistered oracle. The first run built it as 1 - t_cdf(t), which forms
    a ~1E-15 tail by subtracting two numbers near 1 and cost ~6 digits of
    stability at t = 8 - the same complement cancellation the Chi-square study
    caught. Computing the tail directly is the correct realisation of the
    stated oracle, not a change of oracle.
    """
    t, df = mpf(t), mpf(df)
    z = df / (df + t * t)
    return ibeta(z, df / 2, mpf(1) / 2) / 2


def quadrature_tail(t, df):
    """Independent: tanh-sinh quadrature of the density on [t, inf)."""
    t, df = mpf(t), mpf(df)
    logC = mp.loggamma((df + 1) / 2) - mp.loggamma(df / 2) - mp.log(df * mp.pi) / 2

    def f(u):
        return mp.e ** (logC - (df + 1) / 2 * mp.log1p(u * u / df))
    return mp.quad(f, [t, t + 10, mp.inf])


# ---------------------------------------------------------------------------
# Faithful binary64 continued fraction with the VBA tolerance and test.
# ---------------------------------------------------------------------------
NUM_EPS = 3e-14
FPMIN = 1e-300


def betacf_double(a, b, x, maxit=100000):
    qab = a + b; qap = a + 1.0; qam = a - 1.0; c = 1.0
    d = 1.0 - qab * x / qap
    if abs(d) < FPMIN:
        d = FPMIN
    d = 1.0 / d; h = d
    for m in range(1, maxit + 1):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2)); d = 1.0 + aa * d
        if abs(d) < FPMIN:
            d = FPMIN
        c = 1.0 + aa / c
        if abs(c) < FPMIN:
            c = FPMIN
        d = 1.0 / d; h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2)); d = 1.0 + aa * d
        if abs(d) < FPMIN:
            d = FPMIN
        c = 1.0 + aa / c
        if abs(c) < FPMIN:
            c = FPMIN
        d = 1.0 / d; dl = d * c; h *= dl
        if abs(dl - 1.0) <= NUM_EPS:
            return h, m
    return None, maxit


def cf_double_tail(t, df):
    """binary64 tail via the CF on the route the VBA selects, exact prefactor.

    The prefactor is exact so the figure isolates the CF's own error, which is
    what #34 identifies. Returns (value, route, iterations) or (None, ..).
    """
    a = df / 2.0; b = 0.5
    x = df / (df + t * t); y = t * t / (df + t * t)
    xs = (a + 1.0) / (a + b + 2.0)
    mp.dps = 50
    T, DF = mpf(t), mpf(df)
    X = DF / (DF + T * T); Y = T * T / (DF + T * T)
    lb = (mp.loggamma(mpf(a) + b) - mp.loggamma(a) - mp.loggamma(b)
          + a * mp.log(X) + b * mp.log(Y))
    if x < xs:
        h, it = betacf_double(a, b, x)
        if h is None:
            return None, "direct", it
        return mp.e ** lb * h / a / 2, "direct", it
    h, it = betacf_double(b, a, y)
    if h is None:
        return None, "complement", it
    return (1 - mp.e ** lb * h / b) / 2, "complement", it


def t_switch(df):
    a = mpf(df) / 2; b = mpf("0.5")
    return mp.sqrt(mpf(df) * (b + 1) / (a + 1))


def rel(x, y):
    return abs(x - y) / abs(y)


def digits(x, y):
    d = rel(x, y)
    return None if d == 0 else float(-mp.log10(d))


# ---------------------------------------------------------------------------
def evaluate_point(t, df):
    """All measurements at one (t, df). Oracle at 100 dps is the reference."""
    rec = {"t": t, "df": df}
    mp.dps = DPS_HIGH
    ref = oracle_tail(t, df)
    mp.dps = DPS_LOW
    ref_low = oracle_tail(t, df)
    mp.dps = DPS_HIGH
    rec["oracle_stability_digits"] = digits(mpf(ref_low), ref)
    rec["E"] = {}
    rec["E_low"] = {}
    for K in range(1, MAX_ORDER + 1):
        mp.dps = DPS_HIGH
        e_hi = rel(expansion_tail(t, df, K), ref)
        mp.dps = DPS_LOW
        e_lo = rel(expansion_tail(t, df, K), oracle_tail(t, df))
        mp.dps = DPS_HIGH
        rec["E"][K] = e_hi
        rec["E_low"][K] = mpf(e_lo)
    v, route, it = cf_double_tail(t, df)
    mp.dps = DPS_HIGH
    rec["C"] = None if v is None else rel(v, ref)
    rec["cf_route"] = route
    rec["cf_iterations"] = it
    return rec


def fmt(x, n=4):
    return None if x is None else mp.nstr(mpf(x), n)


def main():
    started = time.time()
    problems = []

    # ---- forward fitting ---------------------------------------------------
    fit = []
    for df in FIT_DF:
        for t in FIT_T:
            fit.append(evaluate_point(t, df))
    print(f"forward fitting: {len(fit)} points", flush=True)

    # Oracle stability and precision-pair agreement (sections 3, 4).
    for r in fit:
        d = r["oracle_stability_digits"]
        if d is not None and d < STABILITY_DIGITS:
            problems.append(f"oracle unstable at t={r['t']} df={r['df']:.0e}: {d:.1f} digits")
        # Amendment 1: the precision-pair agreement requirement applies only
        # where the truncation error exceeds the oracle's own floor. Below it
        # both precisions report noise, and demanding they agree measures the
        # oracle, not the expansion. The floor is derived from the measured
        # 60-vs-100 dps oracle stability at that point.
        floor = mpf(10) ** (-(r["oracle_stability_digits"] or STABILITY_DIGITS) + 2)
        for K in range(1, MAX_ORDER + 1):
            e_hi, e_lo = r["E"][K], r["E_low"][K]
            if e_hi > floor and e_lo > floor:
                if abs(mp.log10(e_hi) - mp.log10(e_lo)) > 0.05:   # ~2 sig digits
                    problems.append(
                        f"precision pair disagrees at t={r['t']} df={r['df']:.0e} K={K}: "
                        f"{fmt(e_hi)} vs {fmt(e_lo)}")

    # Monotonic improvement in K where df >= 1E6, |t| <= 4 (section 9).
    for r in fit:
        if r["df"] >= 1e6 and r["t"] <= 4:
            for K in range(1, MAX_ORDER):
                if r["E"][K + 1] > r["E"][K] and r["E"][K] > BOUND:
                    problems.append(
                        f"E not monotone at t={r['t']} df={r['df']:.0e}: "
                        f"E_{K}={fmt(r['E'][K])} E_{K+1}={fmt(r['E'][K+1])}")

    # ---- decision rule (section 8) ----------------------------------------
    def region_ok(df_min, t_max, K):
        for r in fit:
            if r["df"] >= df_min and r["t"] <= t_max and r["E"][K] > BOUND:
                return False
        return True

    adopted = None
    trace = []
    for df_min in DF_MIN_CANDIDATES:
        for t_max in reversed(T_MAX_CANDIDATES):
            for K in range(1, MAX_ORDER + 1):
                ok = region_ok(df_min, t_max, K)
                trace.append({"df_min": df_min, "t_max": t_max, "K": K, "meets_bound": ok})
                if ok:
                    break
            if any(x["meets_bound"] for x in trace if x["df_min"] == df_min and x["t_max"] == t_max):
                K_star = min(x["K"] for x in trace
                             if x["df_min"] == df_min and x["t_max"] == t_max and x["meets_bound"])
                adopted = {"K": K_star, "DF_MIN": df_min, "T_MAX": t_max}
                break
        if adopted:
            break

    improvement_violations = []
    if adopted:
        K = adopted["K"]
        for r in fit:
            if r["df"] >= adopted["DF_MIN"] and r["t"] <= adopted["T_MAX"]:
                if r["C"] is not None and r["C"] > BOUND and not (r["E"][K] < r["C"]):
                    improvement_violations.append(
                        f"expansion not better than CF at t={r['t']} df={r['df']:.0e}: "
                        f"E={fmt(r['E'][K])} C={fmt(r['C'])}")
        problems.extend(improvement_violations)
    else:
        problems.append("no (K, DF_MIN, T_MAX) meets the bound")

    # ---- holdout, evaluated ONCE after adoption (section 8) ---------------
    hold, hold_violations = [], []
    inv_hold, inv_hold_violations = [], []
    inv_fit = []
    if adopted:
        K = adopted["K"]
        for df in HOLD_DF:
            for t in HOLD_T:
                r = evaluate_point(t, df)
                inside = df >= adopted["DF_MIN"] and t <= adopted["T_MAX"]
                r["inside_region"] = inside
                hold.append(r)
                if inside and r["E"][K] > BOUND:
                    hold_violations.append(f"holdout t={t} df={df:.0e}: E_{K}={fmt(r['E'][K])}")
        print(f"forward holdout: {len(hold)} points", flush=True)

        # Inverse round-trip: oracle inverts p -> q at 100 dps; the expansion's
        # tail at q must recover 1-p. Median excluded (kernel returns 0 exactly).
        def inverse_points(offsets, bridge, envelope):
            ps = []
            for d in offsets:
                ps += [0.5 + d, 0.5 - d]
            for b in bridge:
                ps += [b, 1 - b]
            ps += list(envelope)
            return ps

        def round_trip(p, df):
            mp.dps = DPS_HIGH
            P = mpf(p); DF = mpf(df)
            target_tail = 1 - P if P > mpf("0.5") else P
            # invert on the upper tail by symmetry
            lo, hi = mpf(0), mpf(60)
            for _ in range(400):
                mid = (lo + hi) / 2
                if oracle_tail(mid, DF) > target_tail:
                    lo = mid
                else:
                    hi = mid
            q = (lo + hi) / 2
            if q > adopted["T_MAX"] or df < adopted["DF_MIN"]:
                return {"p": p, "df": df, "q": fmt(q, 20), "inside_region": False}
            got = expansion_tail(q, DF, K)
            resid = abs(got - target_tail) / min(P, 1 - P)
            return {"p": p, "df": df, "q": fmt(q, 20), "inside_region": True,
                    "tail_residual": resid}

        for df in INV_DF:
            for p in inverse_points(INV_FIT_OFFSETS, INV_FIT_BRIDGE, INV_FIT_ENVELOPE):
                inv_fit.append(round_trip(p, df))
            for p in inverse_points(INV_HOLD_OFFSETS, INV_HOLD_BRIDGE, ()):
                r = round_trip(p, df)
                inv_hold.append(r)
                if r["inside_region"] and r["tail_residual"] > BOUND:
                    inv_hold_violations.append(
                        f"inverse holdout p={p!r} df={df:.0e}: {fmt(r['tail_residual'])}")
        print(f"inverse: {len(inv_fit)} fitting, {len(inv_hold)} holdout", flush=True)
        problems.extend(hold_violations)
        problems.extend(inv_hold_violations)

    # ---- seams (section 7) -------------------------------------------------
    seams = []
    if adopted:
        K = adopted["K"]

        def probe(label, t, df):
            r = evaluate_point(float(t), float(df))
            return {"seam": label, "t": fmt(t, 15), "df": df,
                    "E_Kstar": fmt(r["E"][K]), "C": fmt(r["C"]),
                    "cf_route": r["cf_route"]}
        for df in FIT_DF:
            ts = t_switch(df)
            for lab, tt in (("below", ts * (1 - SEAM_OFFSET)), ("at", ts),
                            ("above", ts * (1 + SEAM_OFFSET))):
                seams.append(probe(f"cf_switch_{lab}", tt, df))
        dfm = mpf(adopted["DF_MIN"])
        for t in (1.0, 2.0, 4.0):
            for lab, dd in (("below", dfm * (1 - SEAM_OFFSET)), ("at", dfm),
                            ("above", dfm * (1 + SEAM_OFFSET))):
                seams.append(probe(f"df_min_{lab}", mpf(t), float(dd)))
        tm = mpf(adopted["T_MAX"])
        for df in (1e6, 1e8):
            for lab, tt in (("below", tm * (1 - SEAM_OFFSET)), ("at", tm),
                            ("above", tm * (1 + SEAM_OFFSET))):
                seams.append(probe(f"t_max_{lab}", tt, df))
        print(f"seams: {len(seams)} probes", flush=True)

    # ---- quadrature leg on a subset (section 3) ---------------------------
    quad = []
    for df in (1e6, 1e8):
        for t in (1.0, 2.0, 4.0):
            mp.dps = DPS_HIGH
            a = oracle_tail(t, df); q = quadrature_tail(t, df)
            quad.append({"t": t, "df": df, "agreement_digits": digits(q, a)})

    # ---- serialise ---------------------------------------------------------
    def ser(r):
        out = {"t": r["t"], "df": r["df"],
               "oracle_stability_digits": r["oracle_stability_digits"],
               "cf_route": r["cf_route"], "cf_iterations": r["cf_iterations"],
               "C": fmt(r["C"]),
               "E": {str(K): fmt(r["E"][K]) for K in r["E"]},
               "E_low_dps": {str(K): fmt(r["E_low"][K]) for K in r["E_low"]}}
        if "inside_region" in r:
            out["inside_region"] = r["inside_region"]
        return out

    worst_by_K = {}
    for K in range(1, MAX_ORDER + 1):
        w = max(fit, key=lambda r: r["E"][K])
        worst_by_K[str(K)] = {"E": fmt(w["E"][K]), "t": w["t"], "df": w["df"]}

    def ser_inv(r):
        o = dict(r)
        if "tail_residual" in o:
            o["tail_residual"] = fmt(o["tail_residual"])
        return o

    result = {
        "study": "#34 Student-t large-df tail expansion",
        "preregistration": "PREREGISTRATION.md",
        "status": "FAIL" if problems else "PASS",
        "problems": problems,
        "generated_utc": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "runtime_s": round(time.time() - started, 3),
        "versions": {"python": platform.python_version(),
                     "mpmath": mpmath.__version__,
                     "sympy_at_derivation": COEF["sympy_version"]},
        "design": {"precision_pair_dps": [DPS_LOW, DPS_HIGH],
                   "bound": "2^-50", "max_order": MAX_ORDER,
                   "df_min_candidates": list(DF_MIN_CANDIDATES),
                   "t_max_candidates": list(T_MAX_CANDIDATES)},
        "adopted": adopted,
        "decision_trace": trace,
        "worst_point_by_order": worst_by_K,
        "forward_fitting": [ser(r) for r in fit],
        "forward_holdout": [ser(r) for r in hold],
        "inverse_fitting": [ser_inv(r) for r in inv_fit],
        "inverse_holdout": [ser_inv(r) for r in inv_hold],
        "seams": seams,
        "quadrature_leg": quad,
        "holdout_559_inspected": False,
    }
    with open(os.path.join(HERE, "results.json"), "w", newline="\n") as f:
        json.dump(result, f, indent=2)
        f.write("\n")

    print()
    print(f"status: {result['status']}")
    if adopted:
        print(f"adopted: K*={adopted['K']}, DF_MIN={adopted['DF_MIN']:.0e}, "
              f"T_MAX={adopted['T_MAX']}")
    for K in range(1, MAX_ORDER + 1):
        w = worst_by_K[str(K)]
        print(f"  K={K}: worst E={w['E']:>10} at t={w['t']} df={w['df']:.0e}")
    for p in problems[:15]:
        print("  - " + p)
    print(f"runtime {result['runtime_s']} s")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
