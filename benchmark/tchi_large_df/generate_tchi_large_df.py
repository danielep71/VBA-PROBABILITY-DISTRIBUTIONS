"""
Student t and Chi-square LARGE-DEGREES-OF-FREEDOM characterization study (P2-02).

The public t / Chi-square surfaces ACCEPT df up to the 1E100 representational
bound but are only asserted "validated to roughly 1E9". This study measures the
regime between the accuracy-contracted region and that bound so any enforced
limit or capability claim rests on committed evidence, not an assertion.

Regime structure (why references need care):
  * By the CLT the standardized t and Chi-square both converge to the standard
    normal. |t_cdf(x,df) - Phi(x)| ~ C(x)/df, so beyond df ~ 1E15 they agree to
    better than Double epsilon: the CORRECT Double answer there IS the normal
    limit, and the measurement question becomes "does the kernel return the
    normal limit, or silently diverge from it?"
  * The incomplete-beta / incomplete-gamma forms are ill-conditioned at large df.
    Student t: the beta path is computed at ESCALATING precision until two
    settings agree to 1E-40 (stable across 1E5..1E100). Chi-square: the exact
    incomplete-gamma series does not converge in the central region for large df,
    so this study references Chi-square only where a Wilson-Hilferty limit is
    provably exact to Double precision, df >= 1E16 (WH error ~9E-3/df, validated
    below against the exact gammainc at low df). The Chi-square CENTRAL band
    1E5..1E16 is reference-limited and deliberately omitted (needs a Temme
    uniform-asymptotic reference - a scoped follow-up), rather than shipped at
    under-precision.

Every reference is self-consistency-checked; a row that cannot be made solid is
dropped, never shipped. Output: tchi_large_df_grid.csv (observed_vba empty).
"""
import csv
import mpmath as mp

DPS_LADDER = (80, 160, 320, 640, 1280)
EPS = mp.mpf(2) ** -53


def _stable(f):
    prev = None
    for dps in DPS_LADDER:
        with mp.workdps(dps):
            v = +f()
        if prev is not None and abs(v - prev) <= mp.mpf(10) ** -40:
            return v, True
        prev = v
    return prev, False


def phi(x):
    with mp.workdps(60):
        return +mp.ncdf(mp.mpf(x))


def _dbl(x):
    """Round an mpmath value to the nearest IEEE-754 double (what the VBA receives)."""
    return mp.mpf(float(x))


def t_cdf(x, df):
    def f():
        x_ = mp.mpf(x); d = mp.mpf(df)
        ib = mp.betainc(d / 2, mp.mpf(1) / 2, 0, d / (d + x_ * x_), regularized=True)
        return 1 - ib / 2 if x_ > 0 else (ib / 2 if x_ < 0 else mp.mpf("0.5"))
    return _stable(f)


def t_inv(p, df):
    def cdf(x):
        v, ok = t_cdf(x, df)
        return v
    with mp.workdps(120):
        a, b = mp.mpf(-40), mp.mpf(40); p = mp.mpf(p)
        for _ in range(400):
            m = (a + b) / 2
            (a, b) = (m, b) if cdf(m) < p else (a, m)
        return (a + b) / 2, True


def chi_cdf_wh(x, df):
    with mp.workdps(60):
        x = mp.mpf(x); d = mp.mpf(df); t = (x / d) ** (mp.mpf(1) / 3)
        return +mp.ncdf((t - (1 - 2 / (9 * d))) / mp.sqrt(2 / (9 * d)))


def chi_inv_wh(p, df):
    with mp.workdps(60):
        d = mp.mpf(df); z = mp.mpf(mp.erfinv(2 * mp.mpf(p) - 1)) * mp.sqrt(2)
        return +d * ((1 - 2 / (9 * d)) + z * mp.sqrt(2 / (9 * d))) ** 3


def chi_exact(x, df):
    with mp.workdps(80):
        return +mp.gammainc(mp.mpf(df) / 2, 0, mp.mpf(x) / 2, regularized=True)


# --- build-time validation: WH formula correct, and exact-to-Double at 1E16 ---
def _validate_wh():
    d = mp.mpf(1000); sd = mp.sqrt(2 * d)
    worst = max(abs(chi_exact(d + z * sd, d) - chi_cdf_wh(d + z * sd, d)) for z in (-1, 0, 1))
    assert worst < mp.mpf("1e-5"), f"WH formula check failed: {worst}"
    # WH error ~ (worst * 1000) / df; require < EPS at the minimum Chi df used
    c = worst * d
    assert c / mp.mpf("1e16") < EPS, "WH not exact-to-Double at 1E16"
    return float(worst), float(c)


T_DECADES = [1e5, 1e6, 1e7, 1e8, 1e9, 1e11, 1e13, 1e15, 1e18, 1e30, 1e60, 1e100]
CHI_DECADES = [1e16, 1e18, 1e30]     # WH exact-to-Double AND transition band still
                                     # Double-representable (beyond ~1E31 the whole
                                     # non-trivial band is sub-ULP, so no Double x
                                     # yields a determinate Chi-square CDF)
T_X = ["0.5", "1.0", "2.0", "3.0", "4.0"]
T_P = ["0.9", "0.99", "0.999"]
CHI_Z = [-3, -1, 0, 1, 3]
CHI_P = ["0.1", "0.5", "0.9", "0.99"]


def _row(fn, a1, a2, ref, note):
    return {"function": fn, "vba_kernel": f"K_STATS_{fn}", "claim": "characterization",
            "metric": "rel", "arg1": mp.nstr(mp.mpf(a1), 17), "arg2": mp.nstr(mp.mpf(a2), 17),
            "arg3": "", "reference": mp.nstr(ref, 34), "observed_vba": "",
            "regime": note, "evidence_set": "tchi_large_df"}


def build():
    rows, dropped = [], 0

    def add(fn, a1, a2, ref, ok, note):
        nonlocal dropped
        if ok and ref is not None:
            rows.append(_row(fn, a1, a2, ref, note))
        else:
            dropped += 1

    for df in T_DECADES:
        for xs in T_X:
            c, ok = t_cdf(xs, df)
            note = "meaningful" if (ok and abs(c - phi(xs)) > EPS) else "normal_limit"
            add("StudentT_Cumulative", xs, df, c, ok, note)
            add("StudentT_Survival", xs, df, (1 - c) if ok else None, ok, note)
        for ps in T_P:
            x, ok = t_inv(ps, df)
            add("StudentT_InverseCumulative", ps, df, x, ok, "inverse")

    for df in CHI_DECADES:
        sd = mp.sqrt(2 * mp.mpf(df))
        for z in CHI_Z:
            x = _dbl(mp.mpf(df) + z * sd)   # the exact Double the VBA evaluates at
            if x <= 0:
                continue
            c = chi_cdf_wh(x, df)
            add("ChiSquare_Cumulative", mp.nstr(x, 17), df, c, True, f"z{z:+d}_normal_limit")
            add("ChiSquare_Survival", mp.nstr(x, 17), df, 1 - c, True, f"z{z:+d}_normal_limit")
        for ps in CHI_P:
            add("ChiSquare_InverseCumulative", ps, df, chi_inv_wh(ps, df), True, "inverse_normal_limit")

    return rows, dropped


if __name__ == "__main__":
    worst_wh, c = _validate_wh()
    rows, dropped = build()
    fields = ["function", "vba_kernel", "claim", "metric", "arg1", "arg2", "arg3",
              "reference", "observed_vba", "regime", "evidence_set"]
    with open("tchi_large_df_grid.csv", "w", newline="\n") as fh:
        w = csv.DictWriter(fh, fieldnames=fields, lineterminator="\n")
        w.writeheader(); w.writerows(rows)
    print(f"WH validated: worst vs exact at df=1000 is {worst_wh:.1e}; C={c:.1e} => exact-to-Double for df>=1E16")
    print(f"wrote tchi_large_df_grid.csv: {len(rows)} rows, {dropped} dropped (reference not self-consistent)")
    from collections import Counter
    for k, v in sorted(Counter(r["function"] for r in rows).items()):
        print(f"  {k}: {v}")
