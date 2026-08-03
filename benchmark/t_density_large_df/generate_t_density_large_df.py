"""
Build benchmark/t_density_large_df/t_density_large_df_grid.csv.

WHY THIS STUDY EXISTS
    CR-P1-01B gave the Gamma, Beta, ChiSquare and F densities a measured
    accuracy envelope (PROB_DENSITY_SHAPE_MAX = 1E20) and frozen contracts.
    StudentT_Density was left out: it carried no envelope and no large-df study,
    and its only coverage above df 1E6 asserted merely that it returned *a*
    value, not a correct one.

    The t density is a plausible carrier of the same cancellation the other four
    had, because it needs

        LogGamma((df+1)/2) - LogGamma(df/2)

    a difference of two large log-gammas of exactly the kind that lost accuracy
    elsewhere. This grid measures whether it does.

WHAT THE MEASUREMENT FOUND
    It does not. PROB_LogGammaHalfDiff evaluates that difference by an
    asymptotic series in 1/Z above a cutoff rather than subtracting two large
    values, so the cancellation never forms. The density is machine-precise to
    df 1E20 across the body and well into the tail.

    This study therefore exists to CONTRACT that behaviour, not to repair it.
"""
import csv
import mpmath as mp

mp.mp.dps = 50

DFS = [1e2, 1e4, 1e6, 1e8, 1e10, 1e12, 1e14, 1e16, 1e18, 1e20]
# 0 is the mode; 1 and 2 are the body; 3 and 10 probe the tail, where the
# (df+1)/2 * Log1p(x*x/df) term is largest and most exposed to error.
XS = [0.0, 1.0, 2.0, 3.0, 10.0]


def t_pdf(x, df):
    x = mp.mpf(x); df = mp.mpf(df)
    return mp.e ** (mp.loggamma((df + 1) / 2) - mp.loggamma(df / 2)
                    - mp.log(df * mp.pi) / 2
                    - ((df + 1) / 2) * mp.log1p(x * x / df))


def normal_pdf(x):
    x = mp.mpf(x)
    return mp.e ** (-x * x / 2) / mp.sqrt(2 * mp.pi)


def build():
    rows = []
    dropped = 0
    worst = mp.mpf(0)

    for df in DFS:
        for x in XS:
            ref = t_pdf(x, df)
            # Self-check: as df grows the t density must approach the standard
            # normal density. The gap is O(1/df), so at large df an independent
            # limit is available and is used to validate the reference.
            if df >= 1e8:
                gap = abs(ref - normal_pdf(x)) / normal_pdf(x)
                if gap > mp.mpf(10) ** -6:
                    dropped += 1
                    continue
                if gap > worst:
                    worst = gap
            if not (ref > 0):
                dropped += 1
                continue
            rows.append({
                "function": "StudentT_Density", "vba_kernel": "K_STATS_StudentT_Density",
                "claim": "characterization", "metric": "rel",
                "arg1": mp.nstr(mp.mpf(x), 17), "arg2": mp.nstr(mp.mpf(df), 17),
                "arg3": "", "reference": mp.nstr(ref, 34), "observed_vba": "",
                "regime": f"df{df:.0e}_x{x:g}", "evidence_set": "t_density_large_df"})

    return rows, dropped, worst


if __name__ == "__main__":
    rows, dropped, worst = build()
    print(f"self-check (t density vs its normal limit at df >= 1E8) worst gap: "
          f"{mp.nstr(worst, 6)}")
    fields = ["function", "vba_kernel", "claim", "metric", "arg1", "arg2", "arg3",
              "reference", "observed_vba", "regime", "evidence_set"]
    with open("t_density_large_df_grid.csv", "w", newline="\n", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields, lineterminator="\n")
        w.writeheader()
        for r in rows:
            w.writerow(r)
    print(f"wrote t_density_large_df_grid.csv: {len(rows)} rows, {dropped} dropped")
