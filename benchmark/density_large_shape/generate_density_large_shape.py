"""
Large-shape density characterization study (CR-P1-01).

The continuous density kernels evaluate the log-density in the cancellation-prone
normalization form, e.g. Gamma:

    (a-1) log(y) - y - log Gamma(a),   y = X/Scale

At large shape each term is ~ a log a while the true log-density near the mode is
only ~ log a, so the leading terms cancel and expose absolute error from the
separately rounded components. This study measures that error across shape/df
from 1E2 to 1E20 for Gamma, Chi-square, Beta and F densities, so the boundary
where the public densities silently lose accuracy is committed evidence.

References are self-checked TWO ways, which also validates the reviewer's Loader
deviance formulas (15.2 / 15.3): the direct high-precision log-density and the
Loader form (stable deviance + StirlingError) must agree to 1E-40. A row whose
two forms disagree, or whose density is not representable, is dropped rather than
shipped. The stored reference is the density value; the analyzer forms the
absolute log-density error and the relative density error against it.
"""
import csv
import mpmath as mp

mp.mp.dps = 60
TOL = mp.mpf("1e-40")


def stirlerr(a):
    a = mp.mpf(a)
    return mp.loggamma(a) - (a - mp.mpf("0.5")) * mp.log(a) + a - mp.mpf("0.5") * mp.log(2 * mp.pi)


def bd0(a, y):                      # deviance D(a,y) = a log(a/y) + y - a
    a = mp.mpf(a); y = mp.mpf(y)
    return a * mp.log(a / y) + y - a


# ---- Gamma (Chi-square reuses this with a=df/2, scale=2) --------------------
def gamma_logpdf_direct(x, a, scale):
    x = mp.mpf(x); a = mp.mpf(a); scale = mp.mpf(scale)
    return (a - 1) * mp.log(x) - x / scale - a * mp.log(scale) - mp.loggamma(a)


def gamma_logpdf_loader(x, a, scale):
    x = mp.mpf(x); a = mp.mpf(a); scale = mp.mpf(scale); y = x / scale
    return (-bd0(a, y) + mp.log(a) / 2 - mp.log(y) - mp.log(2 * mp.pi) / 2
            - stirlerr(a) - mp.log(scale))


# ---- Beta (reviewer 15.3) --------------------------------------------------
def beta_logpdf_direct(x, a, b):
    x = mp.mpf(x); a = mp.mpf(a); b = mp.mpf(b)
    return (a - 1) * mp.log(x) + (b - 1) * mp.log1p(-x) - (mp.loggamma(a) + mp.loggamma(b) - mp.loggamma(a + b))


def beta_logpdf_loader(x, a, b):
    x = mp.mpf(x); a = mp.mpf(a); b = mp.mpf(b); n = a + b; y = 1 - x
    D = a * mp.log(a / (n * x)) + b * mp.log(b / (n * y))
    return (-D + (mp.log(a) + mp.log(b) - mp.log(n)) / 2 - mp.log(x) - mp.log(y)
            - mp.log(2 * mp.pi) / 2 - stirlerr(a) - stirlerr(b) + stirlerr(n))


# ---- F via Beta transform (reviewer 15.3) ----------------------------------
def f_logpdf_direct(x, d1, d2):
    x = mp.mpf(x); d1 = mp.mpf(d1); d2 = mp.mpf(d2)
    return (d1 / 2 * mp.log(d1) + d2 / 2 * mp.log(d2) + (d1 / 2 - 1) * mp.log(x)
            - (d1 + d2) / 2 * mp.log(d2 + d1 * x)
            - (mp.loggamma(d1 / 2) + mp.loggamma(d2 / 2) - mp.loggamma((d1 + d2) / 2)))


def f_logpdf_loader(x, d1, d2):
    x = mp.mpf(x); d1 = mp.mpf(d1); d2 = mp.mpf(d2)
    r = x * d1 / d2; u = r / (1 + r); v = 1 / (1 + r)
    return beta_logpdf_loader(u, d1 / 2, d2 / 2) + mp.log(u) + mp.log(v) - mp.log(x)


def _dbl(x):
    return mp.mpf(float(x))


SHAPES = [1e2, 1e4, 1e6, 1e8, 1e10, 1e12, 1e16, 1e20]
Z = [-3, -1, 0, 1, 3]

# Extreme degree-ratio F pairs (CR-P1-01B). Near the mode the transformed Beta
# variate U = r / (1 + r) is comfortably interior, so the grid above never
# reaches the regime where U rounds to 1 (or 0) while the density stays finite
# and material - exactly where the pre-fix kernel returned a silent zero. Both
# orientations of every pair are generated: the original defect was asymmetric
# (df1 >> df2 failed while df2 >> df1 was correct), so a one-sided grid would
# have hidden it. All pairs stay inside the validated 1E20 density envelope.
EXTREME_F = [(1e16, 1.0), (1e18, 1.0), (1e20, 1.0), (1e20, 1e2), (1e16, 1e2)]
X_EXTREME = [0.25, 1.0, 4.0]


def _density(direct, loader, *args):
    """Return (density_ref, ok): ok requires direct==loader to 1E-40 and representable."""
    ld = direct(*args); ll = loader(*args)
    if abs(ld - ll) > TOL * max(1, abs(ld)):
        return None, False, abs(ld - ll)
    if ld < -742 or ld > 709:                 # exp under/overflows Double -> not representable
        return None, False, mp.mpf(0)
    return mp.e ** ld, True, abs(ld - ll)


def build():
    rows = []
    dropped = 0
    worst_selfcheck = mp.mpf(0)

    def add(fn, a1, a2, a3, direct, loader, dargs, note):
        nonlocal dropped, worst_selfcheck
        ref, ok, sc = _density(direct, loader, *dargs)
        worst_selfcheck = max(worst_selfcheck, sc)
        if not ok:
            dropped += 1
            return
        rows.append({"function": fn, "vba_kernel": f"K_STATS_{fn}", "claim": "characterization",
                     "metric": "abs_log", "arg1": mp.nstr(mp.mpf(a1), 17),
                     "arg2": mp.nstr(mp.mpf(a2), 17), "arg3": ("" if a3 is None else mp.nstr(mp.mpf(a3), 17)),
                     "reference": mp.nstr(ref, 34), "observed_vba": "", "regime": note,
                     "evidence_set": "density_large_shape"})

    for a in SHAPES:
        sd = mp.sqrt(mp.mpf(a))
        # Gamma (scale=1): X/Scale near the mode
        for z in Z:
            x = _dbl(mp.mpf(a) + z * sd)
            if x <= 0:
                continue
            add("Gamma_Density", x, a, 1.0, gamma_logpdf_direct, gamma_logpdf_loader,
                (x, a, 1.0), f"z{z:+d}")
        # Chi-square: X near df (reuses Gamma with a=df/2, scale=2)
        sdc = mp.sqrt(2 * mp.mpf(a))
        for z in Z:
            x = _dbl(mp.mpf(a) + z * sdc)
            if x <= 0:
                continue
            add("ChiSquare_Density", x, a, None, lambda X, df, s=2: gamma_logpdf_direct(X, mp.mpf(df) / 2, 2),
                lambda X, df, s=2: gamma_logpdf_loader(X, mp.mpf(df) / 2, 2), (x, a), f"z{z:+d}")
        # Beta balanced a=b, x near 0.5
        bsd = 1 / (2 * mp.sqrt(2 * mp.mpf(a)))
        for z in Z:
            x = _dbl(mp.mpf("0.5") + z * bsd)
            if 0 < x < 1:
                add("Beta_Density", x, a, a, beta_logpdf_direct, beta_logpdf_loader, (x, a, a), f"bal_z{z:+d}")
        # Beta unbalanced a=shape, b=shape/100, x near mode a/(a+b)
        b = mp.mpf(a) / 100
        if b >= 1:
            n = mp.mpf(a) + b; mode = (mp.mpf(a) - 1) / (n - 2)
            ubsd = mp.sqrt(mp.mpf(a) * b / (n * n * (n + 1)))
            for z in Z:
                x = _dbl(mode + z * ubsd)
                if 0 < x < 1:
                    add("Beta_Density", x, a, b, beta_logpdf_direct, beta_logpdf_loader, (x, a, b), f"unb_z{z:+d}")
        # F balanced df1=df2, x near 1
        for z in Z:
            x = _dbl(1 + z / sd)              # F(d,d) concentrates near 1 with width ~1/sqrt(d)
            if x > 0:
                add("F_Density", x, a, a, f_logpdf_direct, f_logpdf_loader, (x, a, a), f"bal_z{z:+d}")
        # F unbalanced df1=shape, df2=shape/100, x near the density MODE (r>1 branch).
        # An interior mode needs the second Beta shape df2/2 > 1; the balanced case
        # above already exercises the r=1 boundary, and F_Survival/F CDF the r<1 side.
        d2 = mp.mpf(a) / 100
        if d2 / 2 > 1:
            a1 = mp.mpf(a) / 2; b1 = d2 / 2; nb = a1 + b1
            umode = (a1 - 1) / (nb - 2)
            ustd = mp.sqrt(a1 * b1 / (nb * nb * (nb + 1)))
            for z in Z:
                u = umode + z * ustd
                if 0 < u < 1:
                    r = u / (1 - u)
                    x = _dbl(r * d2 / mp.mpf(a))
                    if x > 0:
                        add("F_Density", x, a, d2, f_logpdf_direct, f_logpdf_loader, (x, a, d2), f"unb_z{z:+d}")

    # Extreme degree ratios, both orientations (CR-P1-01B regression regime).
    for d1, d2 in EXTREME_F:
        for da, db in ((d1, d2), (d2, d1)):
            for x in X_EXTREME:
                add("F_Density", x, da, db, f_logpdf_direct, f_logpdf_loader,
                    (x, da, db), "extreme_ratio")

    return rows, dropped, worst_selfcheck


if __name__ == "__main__":
    rows, dropped, worst = build()
    fields = ["function", "vba_kernel", "claim", "metric", "arg1", "arg2", "arg3",
              "reference", "observed_vba", "regime", "evidence_set"]
    with open("density_large_shape_grid.csv", "w", newline="\n") as fh:
        w = csv.DictWriter(fh, fieldnames=fields, lineterminator="\n")
        w.writeheader(); w.writerows(rows)
    print(f"self-check (direct vs Loader) worst disagreement: {mp.nstr(worst, 3)}  (validates 15.2/15.3)")
    print(f"wrote density_large_shape_grid.csv: {len(rows)} rows, {dropped} dropped")
    from collections import Counter
    for k, v in sorted(Counter(r["function"] for r in rows).items()):
        print(f"  {k}: {v}")
