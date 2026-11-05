"""
Student t and Chi-square LARGE-DEGREES-OF-FREEDOM characterization study (P2-02).

When this study was opened, the public t / Chi-square surfaces accepted df up
to the 1E100 representational bound but were only asserted "validated to roughly
1E9". This study measures the regime between the accuracy-contracted region and
that bound so any enforced limit or capability claim rests on committed
evidence, not an assertion. (Since then the CDF, survival and inverse reject df
above PROB_T_MAX_DF / PROB_CHI_MAX_DF = 1E8; see benchmark/envelope_probe.)

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

Precision and survival (#159). The module now sets its precision explicitly. Before,
it never set mp.mp.dps: each CDF was computed under workdps, but the survival was
then formed as 1 - CDF at mpmath's default 53 bits, so all 75 survival references
were accurate only to about 16 digits although 34 were printed. The module-level
precision is now stated (the 15-digit design precision, which keeps the argument
tokens), and each survival is computed directly, inside a self-checked
evaluation at its own working precision:

  * Student t: I_z(df/2, 1/2) / 2 for x > 0, from the same escalating-precision
    beta evaluation as the CDF.
  * Chi-square: Temme's uniform asymptotic expansion of the regularized
    incomplete gamma (DLMF 8.12), which gives Q and P each directly, in place
    of the Wilson-Hilferty limit. Wilson-Hilferty is exact to Double precision
    at df >= 1E16 but only to 16-29 digits, below the 25-digit reference floor.
    With the c0 and c1 terms the expansion is exact to better than 1E-38
    relative at df >= 1E16; it is validated at build time against the exact
    incomplete gamma at df = 2E8, where the omitted terms are still visible.
    The inverse rows keep the Wilson-Hilferty quantile; they are not survival
    references.

The committed grid holds Excel observations. Re-reference it, changing
reference cells only, with

    python generate_tchi_large_df.py --merge-references          # report
    python generate_tchi_large_df.py --merge-references --write  # patch
"""
import argparse
import csv
import os
import sys

import mpmath as mp

# Design precision, set explicitly: the test points and their argument tokens
# are formed at mpmath's 15-digit default, as when the grid was built, so they
# do not move. No reference is computed at this precision: each one runs under
# its own workdps (the _stable ladder, phi, chi_temme, chi_inv_wh).
mp.mp.dps = 15
HERE = os.path.dirname(os.path.abspath(__file__))
DPS_LADDER = (80, 160, 320, 640, 1280)
EPS = mp.mpf(2) ** -53
TEMME_DPS = 200          # c1 cancels about 3*log10(df) digits near the centre


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


def t_sf(x, df):
    """Upper tail, direct: I_z(df/2, 1/2) / 2 for x > 0, never 1 - CDF."""
    def f():
        x_ = mp.mpf(x); d = mp.mpf(df)
        ib = mp.betainc(d / 2, mp.mpf(1) / 2, 0, d / (d + x_ * x_), regularized=True)
        return ib / 2 if x_ > 0 else (1 - ib / 2 if x_ < 0 else mp.mpf("0.5"))
    return _stable(f)


def chi_temme(x, df):
    """(P, Q) of Chi-square(df) at x by Temme's uniform expansion (DLMF 8.12).

    With a = df/2, z = x/2, lambda = z/a and
    eta = sign(lambda - 1) * sqrt(2 (lambda - 1 - ln lambda)):
        Q = erfc(eta sqrt(a/2)) / 2 + R,   P = erfc(-eta sqrt(a/2)) / 2 - R,
        R = exp(-a eta^2 / 2) / sqrt(2 pi a) * (c0(eta) + c1(eta) / a + ...),
        c0 = 1/(lambda - 1) - 1/eta,
        c1 = 1/eta^3 - 1/(lambda - 1)^3 - 1/(lambda - 1)^2 - 1/(12 (lambda - 1)),
    with c0(0) = -1/3 and c1(0) = -1/540. Both tails come out directly. The
    first omitted term is of order a^-2.5 / 240, below 1E-38 relative to either
    tail at a >= 5E15 within the five-sigma band this study uses."""
    with mp.workdps(TEMME_DPS):
        a = mp.mpf(df) / 2
        z = mp.mpf(x) / 2
        d = (z - a) / a                      # lambda - 1, exact for Doubles
        if d == 0:
            eta = mp.mpf(0)
            c0, c1 = mp.mpf(-1) / 3, mp.mpf(-1) / 540
        else:
            eta = mp.sign(d) * mp.sqrt(2 * (d - mp.log1p(d)))
            c0 = 1 / d - 1 / eta
            c1 = 1 / eta ** 3 - 1 / d ** 3 - 1 / d ** 2 - 1 / (12 * d)
        r = mp.e ** (-a * eta * eta / 2) / mp.sqrt(2 * mp.pi * a) * (c0 + c1 / a)
        w = eta * mp.sqrt(a / 2)
        return +(mp.erfc(-w) / 2 - r), +(mp.erfc(w) / 2 + r)


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


def chi_exact_q(x, df):
    with mp.workdps(80):
        return +mp.gammainc(mp.mpf(df) / 2, mp.mpf(x) / 2, mp.inf, regularized=True)


def _validate_temme():
    """Temme against the exact incomplete gamma at df = 2E8, both tails.

    At a = 1E8 the c1 term is still about 1E-15 of the tail, so a wrong c0 or c1
    would show at that size; the omitted a^-2 term is about 1E-23. The bound
    below sits between the two."""
    with mp.workdps(80):
        d = mp.mpf("2e8"); sd = mp.sqrt(2 * d)
        worst = mp.mpf(0)
        for z in (-3, -1, 0, 1, 3):
            x = mp.mpf(float(d + z * sd))
            p, q = chi_temme(x, d)
            qe = chi_exact_q(x, d)
            worst = max(worst, abs(q - qe) / qe, abs(p - (1 - qe)) / (1 - qe))
        assert worst < mp.mpf("1e-19"), f"Temme expansion check failed: {worst}"
        return float(worst)


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
            s, ok_s = t_sf(xs, df)
            note = "meaningful" if (ok and abs(c - phi(xs)) > EPS) else "normal_limit"
            add("StudentT_Cumulative", xs, df, c, ok, note)
            add("StudentT_Survival", xs, df, s if ok_s else None, ok_s, note)
        for ps in T_P:
            x, ok = t_inv(ps, df)
            add("StudentT_InverseCumulative", ps, df, x, ok, "inverse")

    for df in CHI_DECADES:
        sd = mp.sqrt(2 * mp.mpf(df))
        for z in CHI_Z:
            x = _dbl(mp.mpf(df) + z * sd)   # the exact Double the VBA evaluates at
            if x <= 0:
                continue
            c, s = chi_temme(x, df)
            add("ChiSquare_Cumulative", mp.nstr(x, 17), df, c, True, f"z{z:+d}_normal_limit")
            add("ChiSquare_Survival", mp.nstr(x, 17), df, s, True, f"z{z:+d}_normal_limit")
        for ps in CHI_P:
            add("ChiSquare_InverseCumulative", ps, df, chi_inv_wh(ps, df), True, "inverse_normal_limit")

    return rows, dropped


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--merge-references", action="store_true",
                    help="re-reference the committed grid, keeping every observation")
    ap.add_argument("--write", action="store_true",
                    help="with --merge-references: patch the committed grid")
    args = ap.parse_args()
    worst_wh, c = _validate_wh()
    worst_temme = _validate_temme()
    rows, dropped = build()
    print(f"WH validated: worst vs exact at df=1000 is {worst_wh:.1e}; C={c:.1e} => exact-to-Double for df>=1E16")
    print(f"Temme validated: worst vs exact at df=2E8 is {worst_temme:.1e} (omitted terms ~a^-2)")
    grid = os.path.join(HERE, "tchi_large_df_grid.csv")
    if args.merge_references:
        sys.path.insert(0, os.path.dirname(HERE))
        from _study_merge import merge_references, print_report, study_describe, study_key
        changes, held = merge_references(grid, rows, study_key, {"observed_vba"},
                                         write=args.write, describe=study_describe)
        print_report("tchi_large_df_grid.csv", changes, held, args.write)
        return
    fields = ["function", "vba_kernel", "claim", "metric", "arg1", "arg2", "arg3",
              "reference", "observed_vba", "regime", "evidence_set"]
    with open(grid, "w", newline="\n") as fh:
        w = csv.DictWriter(fh, fieldnames=fields, lineterminator="\n")
        w.writeheader(); w.writerows(rows)
    print(f"wrote tchi_large_df_grid.csv: {len(rows)} rows, {dropped} dropped (reference not self-consistent)")
    from collections import Counter
    for k, v in sorted(Counter(r["function"] for r in rows).items()):
        print(f"  {k}: {v}")


if __name__ == "__main__":
    main()
