"""
Score the large-shape density study and show the before/after for CR-P1-01.

For each (function, shape) it reports, once the VBA observations are exported:
  * the worst absolute log-density error of the current kernel (naive form);
  * the worst relative density error where representable;
  * for comparison, the error a stable Loader-in-Double implementation achieves,
    so the improvement the fix would deliver is visible next to the defect.

Run with no observations to see the stable-target column alone (the naive column
needs the exported VBA values).
"""
import csv
import math
import os
from collections import defaultdict
from decimal import Decimal, InvalidOperation

import mpmath as mp
mp.mp.dps = 50


def parse_obs(s):
    s = (s or "").strip()
    if not s or s.upper() == "ERROR":
        return None
    try:
        return float(sum(Decimal(p.strip()) for p in s.split(";")))
    except (InvalidOperation, ValueError):
        return "BAD"


# stable stirling error (stands in for PROB_StirlingError, accurate ~1E-13+)
def stirlerr(a):
    return float(mp.loggamma(a) - (mp.mpf(a) - mp.mpf('0.5')) * mp.log(a) + mp.mpf(a) - mp.log(2 * mp.pi) / 2)


def bd0(k, m):
    """Loader's stable deviance k*log(k/m)+m-k, cancellation-free near k==m."""
    u = (k - m) / m
    return m * ((1 + u) * math.log1p(u) - u)


def gamma_loader(x, a, scale=1.0):
    y = x / scale
    return -bd0(a, y) + 0.5 * math.log(a) - math.log(y) - 0.5 * math.log(2 * math.pi) - stirlerr(a) - math.log(scale)


def beta_loader(x, a, b):
    n = a + b; y = 1 - x
    D = bd0(a, n * x) + bd0(b, n * y)          # stable deviance (not the raw D)
    return (-D + 0.5 * (math.log(a) + math.log(b) - math.log(n)) - math.log(x) - math.log(y)
            - 0.5 * math.log(2 * math.pi) - stirlerr(a) - stirlerr(b) + stirlerr(n))


def f_loader(x, d1, d2):
    r = x * d1 / d2; u = r / (1 + r); v = 1 / (1 + r)
    return beta_loader(u, d1 / 2, d2 / 2) + math.log(u) + math.log(v) - math.log(x)


def loader_logpdf(fn, a1, a2, a3):
    if fn == "Gamma_Density":
        return gamma_loader(a1, a2, a3)
    if fn == "ChiSquare_Density":
        return gamma_loader(a1, a2 / 2.0, 2.0)
    if fn == "Beta_Density":
        return beta_loader(a1, a2, a3)
    if fn == "F_Density":
        return f_loader(a1, a2, a3)
    raise ValueError(fn)


def main():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "density_large_shape_grid.csv")
    rows = list(csv.DictReader(open(path)))
    any_obs = any((r["observed_vba"] or "").strip() for r in rows)
    by = defaultdict(list)
    for r in rows:
        by[(r["function"], r["arg2"])].append(r)

    print(f"{'function':<20}{'shape/df':>10}{'n':>4}{'naive absLog':>14}{'naive relDens':>15}{'loader absLog':>15}")
    worst_naive = 0.0
    for (fn, a2), rs in sorted(by.items(), key=lambda k: (k[0][0], float(k[0][1]))):
        wn = wl = -1.0; wrel = -1.0; miss = 0
        for r in rs:
            ref = float(mp.mpf(r["reference"]))
            a1 = float(r["arg1"]); a2v = float(r["arg2"]); a3v = float(r["arg3"]) if r["arg3"].strip() else None
            try:
                ll = loader_logpdf(fn, a1, a2v, a3v); wl = max(wl, abs(ll - math.log(ref)))
            except Exception:
                pass
            if any_obs:
                o = parse_obs(r["observed_vba"])
                if o is None or o == "BAD" or o <= 0:
                    miss += 1; continue
                wn = max(wn, abs(math.log(o) - math.log(ref)))
                wrel = max(wrel, abs(o - ref) / ref)
        worst_naive = max(worst_naive, wn)
        nn = f"{wn:.2e}" if wn >= 0 else ("n/a" if any_obs else "-")
        nr = f"{wrel:.2e}" if wrel >= 0 else ("n/a" if any_obs else "-")
        nl = f"{wl:.2e}" if wl >= 0 else "n/a"
        flag = "  <-- DEGRADED" if (wn > 1e-6) else ""
        m = f" [{miss} bad]" if miss else ""
        print(f"{fn:<20}{float(a2):>10.0e}{len(rs):>4}{nn:>14}{nr:>15}{nl:>15}{flag}{m}")

    if any_obs:
        print(f"\nworst naive absolute log-density error: {worst_naive:.2e}")
        print("The loader column is what a stable Loader-in-Double kernel achieves on the")
        print("same points (StirlingError + stable deviance); the gap is the CR-P1-01 defect.")
    else:
        print("\nNo observations yet - showing the stable Loader target only. Export the VBA")
        print("densities into observed_vba and re-run to see the current-kernel error beside it.")


if __name__ == "__main__":
    main()
