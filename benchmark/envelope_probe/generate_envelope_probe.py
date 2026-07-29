"""
Build benchmark/envelope_probe/envelope_probe_grid.csv.

WHY THIS STUDY EXISTS
    The StudentT, ChiSquare and F envelopes (1E6, 1E6, 1E5) were set before
    CR-P1-02 repaired the incomplete-gamma and incomplete-beta prefactors. The
    P2-02 chi divergence at df = 1E16 is plausibly that same defect, in which
    case the caps measure a bug rather than an intrinsic limit.

    The public functions REJECT df above their cap, so they cannot be used to
    find out. This grid therefore probes the underlying PUBLIC kernels directly,
    at exactly the parameters each family would pass them:

        ChiSquare_Cumulative(X, df) -> PROB_TryGammaRegularizedP(df/2, X/2)
        StudentT_Cumulative(t, df)  -> PROB_TryBetaRegularized(x, y, df/2, 1/2)
                                       with x = df / (df + t*t)
        F_Cumulative(f, d1, d2)     -> PROB_TryBetaRegularized(x, y, d1/2, d2/2)
                                       with x = d1*f / (d1*f + d2)

    No source change is needed to run it, and the result is direct evidence for
    where each cap can honestly sit.

REFERENCES
    mpmath at 30 digits by the converging route, revalidated at 45 digits.
    Anything that does not survive revalidation is dropped, not trusted.
"""
import csv
import mpmath as mp

mp.mp.dps = 30
EPS = mp.mpf(10) ** -26
TOL = mp.mpf(10) ** -22

CHI_DF = [1e6, 1e7, 1e8, 1e9, 1e10, 1e12, 1e14, 1e16]
T_DF = [1e6, 1e7, 1e8, 1e10, 1e12, 1e14, 1e16]
F_DF = [(1e5, 1e5), (1e6, 1e6), (1e8, 1e8), (1e10, 1e10), (1e12, 1e12),
        (1e6, 1e2), (1e2, 1e6), (1e10, 1e4), (1e4, 1e10)]
Z = [-1, 0, 1]


def gamma_P_series(A, X):
    A = mp.mpf(A); X = mp.mpf(X)
    Ap = A; S = 1 / A; Del = S; n = 0
    while n < 3000000:
        Ap += 1; Del *= X / Ap; S += Del; n += 1
        if abs(Del) <= abs(S) * EPS:
            break
    return S * mp.e ** (-X + A * mp.log(X) - mp.loggamma(A))


def gamma_Q_cf(A, X):
    A = mp.mpf(A); X = mp.mpf(X)
    tiny = mp.mpf(10) ** -300
    b = X + 1 - A; c = 1 / tiny; d = 1 / b; h = d
    for i in range(1, 200000):
        an = -mp.mpf(i) * (mp.mpf(i) - A)
        b += 2
        d = an * d + b
        if abs(d) < tiny: d = tiny
        c = b + an / c
        if abs(c) < tiny: c = tiny
        d = 1 / d; de = d * c; h *= de
        if abs(de - 1) <= EPS:
            break
    return h * mp.e ** (-X + A * mp.log(X) - mp.loggamma(A))


def gamma_P(A, X, dps):
    prev = mp.mp.dps
    mp.mp.dps = dps
    try:
        A = mp.mpf(A); X = mp.mpf(X)
        return +(gamma_P_series(A, X) if X < A + 1 else 1 - gamma_Q_cf(A, X))
    finally:
        mp.mp.dps = prev


def beta_cf(A, B, X):
    A = mp.mpf(A); B = mp.mpf(B); X = mp.mpf(X)
    tiny = mp.mpf(10) ** -300
    qab = A + B; qap = A + 1; qam = A - 1
    c = mp.mpf(1); d = 1 - qab * X / qap
    if abs(d) < tiny: d = tiny
    d = 1 / d; h = d
    for m in range(1, 200000):
        m2 = 2 * m
        aa = m * (B - m) * X / ((qam + m2) * (A + m2))
        d = 1 + aa * d
        if abs(d) < tiny: d = tiny
        c = 1 + aa / c
        if abs(c) < tiny: c = tiny
        d = 1 / d; h *= d * c
        aa = -(A + m) * (qab + m) * X / ((A + m2) * (qap + m2))
        d = 1 + aa * d
        if abs(d) < tiny: d = tiny
        c = 1 + aa / c
        if abs(c) < tiny: c = tiny
        d = 1 / d; de = d * c; h *= de
        if abs(de - 1) <= EPS:
            break
    return h


def beta_I(A, B, X, dps):
    prev = mp.mp.dps
    mp.mp.dps = dps
    try:
        A = mp.mpf(A); B = mp.mpf(B); X = mp.mpf(X)
        if X <= 0: return mp.mpf(0)
        if X >= 1: return mp.mpf(1)
        lf = (A * mp.log(X) + B * mp.log1p(-X)
              - (mp.loggamma(A) + mp.loggamma(B) - mp.loggamma(A + B)))
        front = mp.e ** lf
        if X < (A + 1) / (A + B + 2):
            return +(front * beta_cf(A, B, X) / A)
        return +(1 - front * beta_cf(B, A, 1 - X) / B)
    finally:
        mp.mp.dps = prev


def build():
    rows = []
    dropped = 0
    worst = mp.mpf(0)

    def add(fn, args, ref, regime):
        a = list(args) + [None] * (4 - len(args))
        rows.append({
            "function": fn, "vba_kernel": fn, "claim": "characterization",
            "metric": "rel",
            "arg1": mp.nstr(mp.mpf(a[0]), 17),
            "arg2": mp.nstr(mp.mpf(a[1]), 17),
            "arg3": ("" if a[2] is None else mp.nstr(mp.mpf(a[2]), 17)),
            "arg4": ("" if a[3] is None else mp.nstr(mp.mpf(a[3]), 17)),
            "reference": mp.nstr(ref, 34), "observed_vba": "",
            "regime": regime, "evidence_set": "envelope_probe"})

    def keep(v, v2):
        nonlocal dropped, worst
        d = abs(v - v2)
        if d > TOL or not (0 < v < 1):
            dropped += 1
            return False
        if d > worst:
            worst = d
        return True

    # Chi-square through the incomplete gamma
    for df in CHI_DF:
        A = mp.mpf(df) / 2
        sd = mp.sqrt(2 * mp.mpf(df))
        for z in Z:
            X = mp.mpf(float(mp.mpf(df) + z * sd)) / 2
            if X <= 0: continue
            v = gamma_P(A, X, 30); v2 = gamma_P(A, X, 45)
            if keep(v, v2):
                add("PROB_TryGammaRegularizedP", [A, X], v, f"chi_df{df:.0e}_z{z:+d}")

    # Student t through the incomplete beta (b = 1/2)
    for df in T_DF:
        A = mp.mpf(df) / 2
        for z in Z:
            t = mp.mpf(float(1 + z))
            if t <= 0: continue
            X = mp.mpf(df) / (mp.mpf(df) + t * t)
            Y = 1 - X
            v = beta_I(A, mp.mpf("0.5"), X, 30); v2 = beta_I(A, mp.mpf("0.5"), X, 45)
            if keep(v, v2):
                add("PROB_TryBetaRegularized", [X, Y, A, mp.mpf("0.5")], v,
                    f"t_df{df:.0e}_t{float(t):g}")

    # F through the incomplete beta
    for d1, d2 in F_DF:
        A = mp.mpf(d1) / 2; B = mp.mpf(d2) / 2
        for z in Z:
            f = mp.mpf(float(1 + z * mp.mpf("0.1")))
            if f <= 0: continue
            X = mp.mpf(d1) * f / (mp.mpf(d1) * f + mp.mpf(d2))
            Y = 1 - X
            v = beta_I(A, B, X, 30); v2 = beta_I(A, B, X, 45)
            if keep(v, v2):
                add("PROB_TryBetaRegularized", [X, Y, A, B], v,
                    f"f_df{d1:.0e}_{d2:.0e}_f{float(f):g}")

    return rows, dropped, worst


if __name__ == "__main__":
    rows, dropped, worst = build()
    print(f"self-check (30 vs 45 digits) worst disagreement: {mp.nstr(worst, 6)}")
    if worst > TOL:
        raise SystemExit("ABORT: reference self-check failed; no grid written")
    fields = ["function", "vba_kernel", "claim", "metric", "arg1", "arg2", "arg3",
              "arg4", "reference", "observed_vba", "regime", "evidence_set"]
    with open("envelope_probe_grid.csv", "w", newline="\n", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields, lineterminator="\n")
        w.writeheader()
        for r in rows:
            w.writerow(r)
    print(f"wrote envelope_probe_grid.csv: {len(rows)} rows, {dropped} dropped")
    from collections import Counter
    fam = Counter(r["regime"].split("_")[0] for r in rows)
    for k, v in sorted(fam.items()):
        print(f"  {k}: {v}")
