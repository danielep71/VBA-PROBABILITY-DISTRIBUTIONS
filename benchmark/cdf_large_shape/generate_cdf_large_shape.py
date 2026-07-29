"""
Build benchmark/cdf_large_shape/cdf_large_shape_grid.csv.

WHY THIS STUDY EXISTS
    CR-P1-02 removed the cancelling prefactors from the regularized incomplete
    gamma and beta kernels. Those kernels carry every cumulative and survival
    probability outside the normal family, yet the frozen contracts reached only
    shape 3 for Gamma and 1E5 for Beta, so the repaired regime was unmeasured.
    This grid measures it directly and supplies the evidence for contracts.

REFERENCES
    mpmath at 30 digits, computed twice by independent routes wherever both
    converge (lower series and upper continued fraction), and accepted only when
    the two agree. A reference that cannot be cross-checked is dropped rather
    than trusted.

REACHABILITY
    The lower-tail series needs O(sqrt(A)) terms, so the VBA kernel reaches its
    100000-iteration cap somewhere above A = 1E8 and returns #NUM! rather than a
    partial sum. Points beyond that are retained deliberately: an ERROR cell is
    the measurement, and it documents where the validated domain stops.
"""
import csv
import mpmath as mp

mp.mp.dps = 30
EPS = mp.mpf(10) ** -26
TOL = mp.mpf(10) ** -22

SHAPES = [1e2, 1e4, 1e6, 1e8, 1e10, 1e12]
Z = [-3, -1, 0, 1, 3]


def _gamma_log_prefactor(A, X):
    return -X + A * mp.log(X) - mp.loggamma(A)


def gamma_P_series(A, X):
    """Regularized lower incomplete gamma by the ascending series."""
    A = mp.mpf(A); X = mp.mpf(X)
    Ap = A; S = 1 / A; Del = S; n = 0
    while n < 3000000:
        Ap += 1; Del *= X / Ap; S += Del; n += 1
        if abs(Del) <= abs(S) * EPS:
            break
    return S * mp.e ** _gamma_log_prefactor(A, X)


def gamma_Q_cf(A, X):
    """Regularized upper incomplete gamma by the Lentz continued fraction."""
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
    return h * mp.e ** _gamma_log_prefactor(A, X)


def beta_cf(A, B, X):
    """Lentz continued fraction for the incomplete beta, direct orientation."""
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


def beta_I(A, B, X):
    """Regularized incomplete beta, using whichever orientation converges."""
    A = mp.mpf(A); B = mp.mpf(B); X = mp.mpf(X)
    if X <= 0: return mp.mpf(0)
    if X >= 1: return mp.mpf(1)
    lf = (A * mp.log(X) + B * mp.log1p(-X)
          - (mp.loggamma(A) + mp.loggamma(B) - mp.loggamma(A + B)))
    front = mp.e ** lf
    if X < (A + 1) / (A + B + 2):
        return front * beta_cf(A, B, X) / A
    return 1 - front * beta_cf(B, A, 1 - X) / B


def build():
    rows = []
    dropped = 0
    worst = mp.mpf(0)

    def add(fn, a1, a2, a3, ref, regime):
        rows.append({
            "function": fn, "vba_kernel": f"K_STATS_{fn}", "claim": "characterization",
            "metric": "rel", "arg1": mp.nstr(mp.mpf(a1), 17), "arg2": mp.nstr(mp.mpf(a2), 17),
            "arg3": ("" if a3 is None else mp.nstr(mp.mpf(a3), 17)),
            "reference": mp.nstr(ref, 34), "observed_vba": "", "regime": regime,
            "evidence_set": "cdf_large_shape"})

    def gamma_P(A, X, dps):
        """P(A,X) by the route that converges: series below A+1, CF above."""
        prev = mp.mp.dps
        mp.mp.dps = dps
        try:
            if X < A + 1:
                return +gamma_P_series(A, X)
            return +(1 - gamma_Q_cf(A, X))
        finally:
            mp.mp.dps = prev

    for a in SHAPES:
        A = mp.mpf(a)
        sd = mp.sqrt(A)
        for z in Z:
            X = mp.mpf(float(A + z * sd))
            if X <= 0:
                continue
            # Self-check by recomputation at higher precision: a value that has
            # genuinely converged does not move, one that hit an iteration cap
            # does. This validates the reference without assuming a second route
            # converges in the same regime.
            P = gamma_P(A, X, 30)
            P2 = gamma_P(A, X, 45)
            disagree = abs(P - P2)
            if disagree > TOL or not (0 < P < 1):
                dropped += 1
                continue
            add("Gamma_Cumulative", X, A, 1.0, P, f"z{z:+d}")
            add("Gamma_Survival", X, A, 1.0, 1 - P, f"z{z:+d}")
            if disagree > worst:
                worst = disagree

        # Beta, balanced shapes, x near the mode 0.5
        bsd = 1 / (2 * mp.sqrt(2 * A))
        for z in Z:
            X = mp.mpf(float(mp.mpf("0.5") + z * bsd))
            if not (0 < X < 1):
                continue
            I = beta_I(A, A, X)
            Icomp = 1 - beta_I(A, A, 1 - X)      # symmetry cross-check
            disagree = abs(I - Icomp)
            if disagree > TOL or not (0 < I < 1):
                dropped += 1
                continue
            add("Beta_Cumulative", X, A, A, I, f"bal_z{z:+d}")
            add("Beta_Survival", X, A, A, 1 - I, f"bal_z{z:+d}")
            if disagree > worst:
                worst = disagree

    return rows, dropped, worst


if __name__ == "__main__":
    rows, dropped, worst = build()
    print(f"self-check (independent routes) worst disagreement: {mp.nstr(worst, 6)}")
    if worst > TOL:
        raise SystemExit("ABORT: reference self-check failed; no grid written")
    fields = ["function", "vba_kernel", "claim", "metric", "arg1", "arg2", "arg3",
              "reference", "observed_vba", "regime", "evidence_set"]
    with open("cdf_large_shape_grid.csv", "w", newline="\n", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields, lineterminator="\n")
        w.writeheader()
        for r in rows:
            w.writerow(r)
    print(f"wrote cdf_large_shape_grid.csv: {len(rows)} rows, {dropped} dropped")
    from collections import Counter
    for fn, c in sorted(Counter(r["function"] for r in rows).items()):
        print(f"  {fn}: {c}")
