"""
Build benchmark/inverse_probe/inverse_probe_grid.csv.

WHY THIS STUDY EXISTS
    The forward envelopes were raised to df 1E8 (StudentT, ChiSquare) and 1E10
    (F) on the evidence of benchmark/envelope_probe. That study measured the
    FORWARD kernels only. The inverses add safeguarded Newton plus bisection
    with their own iteration budget, and K_STATS_F_InverseCumulative is measured
    to refuse at df (1E6, 3) - well inside the raised forward cap - so the
    forward evidence does not transfer. The inverses therefore kept their
    previously validated caps (F 1E5, StudentT/ChiSquare 1E6) pending this
    study.

    As in envelope_probe, the public inverses reject beyond their caps, so this
    grid probes the PUBLIC kernels directly:

        ChiSquare_InverseCumulative(p, df)  -> PROB_TryGammaInvP(p, q, df/2)
        F_InverseCumulative(p, d1, d2)      -> PROB_TryBetaInvRegularized(
                                                   p, q, d1/2, d2/2)
        StudentT_InverseCumulative(p, df)   -> the same beta inverse with
                                               B = 1/2

    Both take the probability and its complement as a pair, so no tail is
    reconstructed by subtraction.

WHAT IS MEASURED
    The reference is the quantile itself, obtained by bisection on the
    high-precision CDF and cross-checked by feeding it back through that CDF.
    A refusal (#NUM!) is a measurement too: it marks where the iteration gives
    up rather than returning an unconverged root.

DESIGN NOTE
    Degree-of-freedom RATIO is varied as well as magnitude. The known failure is
    at ratio 1E6:3, not at large balanced df, so a magnitude-only grid would
    have missed the very case that motivated the study.
"""
import csv
import mpmath as mp

mp.mp.dps = 30
EPS = mp.mpf(10) ** -26
TOL = mp.mpf(10) ** -20

CHI_DF = [1e6, 1e7, 1e8]
# F degree-of-freedom PAIRS. These are df, converted to beta shapes as
# A = d1/2, B = d2/2 exactly as K_STATS_F_InverseCumulative does. An earlier
# version of this generator passed the pairs straight through as beta shapes,
# which silently probed shapes twice the intended size and therefore missed the
# very configuration that motivated the study.
F_DF_BALANCED = [(2e5, 2e5), (2e6, 2e6), (2e8, 2e8), (2e10, 2e10)]
F_DF_UNBALANCED = [(1e6, 3.0), (2.0, 1e6),          # the measured refusals
                   (1e6, 4.0), (1e8, 10.0), (10.0, 1e8),
                   (1e10, 100.0), (100.0, 1e10)]
T_DF = [1e6, 1e7, 1e8]
# 0.95 and 0.99 are included because the measured refusals occur there too, and
# an inverse can converge at the median while failing further into the tail.
PROBS = [mp.mpf("0.1"), mp.mpf("0.5"), mp.mpf("0.9"),
         mp.mpf("0.95"), mp.mpf("0.99")]


def gser(A, X):
    A = mp.mpf(A); X = mp.mpf(X); Ap = A; S = 1 / A; D = S; n = 0
    while n < 3000000:
        Ap += 1; D *= X / Ap; S += D; n += 1
        if abs(D) <= abs(S) * EPS:
            break
    return S * mp.e ** (-X + A * mp.log(X) - mp.loggamma(A))


def gcf(A, X):
    A = mp.mpf(A); X = mp.mpf(X); tiny = mp.mpf(10) ** -300
    b = X + 1 - A; c = 1 / tiny; d = 1 / b; h = d
    for i in range(1, 200000):
        an = -mp.mpf(i) * (mp.mpf(i) - A); b += 2
        d = an * d + b
        if abs(d) < tiny: d = tiny
        c = b + an / c
        if abs(c) < tiny: c = tiny
        d = 1 / d; de = d * c; h *= de
        if abs(de - 1) <= EPS:
            break
    return h * mp.e ** (-X + A * mp.log(X) - mp.loggamma(A))


def gamma_P(A, X):
    A = mp.mpf(A); X = mp.mpf(X)
    return gser(A, X) if X < A + 1 else 1 - gcf(A, X)


def bcf(A, B, X):
    A = mp.mpf(A); B = mp.mpf(B); X = mp.mpf(X); tiny = mp.mpf(10) ** -300
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
    A = mp.mpf(A); B = mp.mpf(B); X = mp.mpf(X)
    if X <= 0: return mp.mpf(0)
    if X >= 1: return mp.mpf(1)
    lf = (A * mp.log(X) + B * mp.log1p(-X)
          - (mp.loggamma(A) + mp.loggamma(B) - mp.loggamma(A + B)))
    fr = mp.e ** lf
    if X < (A + 1) / (A + B + 2):
        return fr * bcf(A, B, X) / A
    return 1 - fr * bcf(B, A, 1 - X) / B


def gamma_pdf(A, X):
    A = mp.mpf(A); X = mp.mpf(X)
    return mp.e ** ((A - 1) * mp.log(X) - X - mp.loggamma(A))


def beta_pdf(A, B, X):
    A = mp.mpf(A); B = mp.mpf(B); X = mp.mpf(X)
    return mp.e ** ((A - 1) * mp.log(X) + (B - 1) * mp.log1p(-X)
                    - (mp.loggamma(A) + mp.loggamma(B) - mp.loggamma(A + B)))


def newton(cdf, pdf, x0, target, lo, hi, iters=60):
    """
    Safeguarded Newton: a Newton step when it stays inside the bracket, a
    bisection step otherwise. Each CDF evaluation is expensive at large shape
    (the ascending series needs O(sqrt(A)) terms), so plain bisection would need
    hundreds of them; this needs a handful while keeping bisection's guarantee
    that the bracket always shrinks.
    """
    x = mp.mpf(x0); lo = mp.mpf(lo); hi = mp.mpf(hi)
    for _ in range(iters):
        f = cdf(x) - target
        if f > 0:
            hi = x
        else:
            lo = x
        d = pdf(x)
        step = f / d if d != 0 else mp.mpf(0)
        nx = x - step
        if not (lo < nx < hi):
            nx = (lo + hi) / 2
        if abs(nx - x) <= abs(x) * mp.mpf(10) ** -32:
            return nx
        x = nx
    return x


def wilson_hilferty(A, p):
    """Cube-root normal approximation, a good starting point for the gamma."""
    z = mp.sqrt(2) * mp.erfinv(2 * mp.mpf(p) - 1)
    t = 1 - 2 / (9 * A) + z * mp.sqrt(2 / (9 * A))
    return A * t ** 3


def build():
    rows = []
    dropped = 0
    worst = mp.mpf(0)

    def add(fn, args, ref, regime):
        a = list(args) + [None] * (4 - len(args))
        rows.append({
            "function": fn, "vba_kernel": fn, "claim": "characterization",
            "metric": "rel",
            "arg1": mp.nstr(a[0], 17), "arg2": mp.nstr(a[1], 17),
            "arg3": ("" if a[2] is None else mp.nstr(mp.mpf(a[2]), 17)),
            "arg4": ("" if a[3] is None else mp.nstr(mp.mpf(a[3]), 17)),
            "reference": mp.nstr(ref, 34), "observed_vba": "",
            "regime": regime, "evidence_set": "inverse_probe"})

    def check(fwd, root, p):
        """Round-trip the reference back through the CDF before trusting it."""
        nonlocal dropped, worst
        d = abs(fwd(root) - p)
        if d > TOL:
            dropped += 1
            return False
        if d > worst:
            worst = d
        return True

    # Gamma inverse (chi-square)
    for df in CHI_DF:
        A = mp.mpf(df) / 2
        for p in PROBS:
            sd = mp.sqrt(A)
            root = newton(lambda x: gamma_P(A, x), lambda x: gamma_pdf(A, x),
                          wilson_hilferty(A, p), p,
                          max(mp.mpf("1e-30"), A - 60 * sd), A + 60 * sd)
            if check(lambda x: gamma_P(A, x), root, p):
                add("PROB_TryGammaInvP", [p, 1 - p, A], root, f"chi_df{df:.0e}_p{float(p)}")

    # Beta inverse: balanced, unbalanced, and the t shape (B = 1/2)
    groups = ([(mp.mpf(d1) / 2, mp.mpf(d2) / 2, "bal") for d1, d2 in F_DF_BALANCED]
              + [(mp.mpf(d1) / 2, mp.mpf(d2) / 2, "unb") for d1, d2 in F_DF_UNBALANCED]
              + [(mp.mpf(df) / 2, mp.mpf("0.5"), "t") for df in T_DF])
    for A, B, tag in groups:
        A = mp.mpf(A); B = mp.mpf(B)
        for p in PROBS:
            m = A / (A + B)
            root = newton(lambda x: beta_I(A, B, x), lambda x: beta_pdf(A, B, x),
                          m, p, mp.mpf(0), mp.mpf(1))
            if not (0 < root < 1):
                dropped += 1
                continue
            if check(lambda x: beta_I(A, B, x), root, p):
                add("PROB_TryBetaInvRegularized", [p, 1 - p, A, B], root,
                    f"{tag}_a{float(A):.0e}_b{float(B):.0e}_p{float(p)}")

    return rows, dropped, worst


if __name__ == "__main__":
    rows, dropped, worst = build()
    print(f"self-check (reference fed back through the CDF) worst residual: {mp.nstr(worst, 6)}")
    if worst > TOL:
        raise SystemExit("ABORT: reference self-check failed; no grid written")
    fields = ["function", "vba_kernel", "claim", "metric", "arg1", "arg2", "arg3",
              "arg4", "reference", "observed_vba", "regime", "evidence_set"]
    with open("inverse_probe_grid.csv", "w", newline="\n", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields, lineterminator="\n")
        w.writeheader()
        for r in rows:
            w.writerow(r)
    print(f"wrote inverse_probe_grid.csv: {len(rows)} rows, {dropped} dropped")
    from collections import Counter
    for k, v in sorted(Counter(r["regime"].split("_")[0] for r in rows).items()):
        print(f"  {k}: {v}")
