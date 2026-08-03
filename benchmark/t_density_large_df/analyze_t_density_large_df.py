"""
Report the measured accuracy of K_STATS_StudentT_Density at large df.

Run after Export_TDensityLargeDf has filled observed_vba.
"""
import csv
import re
from collections import defaultdict
from decimal import Decimal


def observed(text):
    text = text.strip()
    if not text or text.upper() == "ERROR":
        return None
    return float(sum(Decimal(part) for part in text.split(";")))


def main():
    rows = list(csv.DictReader(open("t_density_large_df_grid.csv", encoding="utf-8")))
    worst_df = defaultdict(float)
    worst_x = defaultdict(float)
    counts = defaultdict(lambda: [0, 0])

    for r in rows:
        df = float(r["arg2"]); x = float(r["arg1"])
        obs = observed(r["observed_vba"])
        if obs is None:
            counts[df][1] += 1
            continue
        ref = float(r["reference"])
        counts[df][0] += 1
        if ref != 0:
            rel = abs(obs - ref) / abs(ref)
            worst_df[df] = max(worst_df[df], rel)
            worst_x[x] = max(worst_x[x], rel)

    print(f"{'df':>10}{'worst rel':>13}{'n':>4}{'#NUM!':>7}")
    for df in sorted(counts):
        ok, bad = counts[df]
        w = f"{worst_df[df]:.2e}" if ok else "n/a"
        print(f"{df:>10.0e}{w:>13}{ok:>4}{bad:>7}")

    print(f"\n{'x':>10}{'worst rel':>13}   (tail exposure: the (df+1)/2 * Log1p term "
          f"grows with x)")
    for x in sorted(worst_x):
        print(f"{x:>10g}{worst_x[x]:>13.2e}")

    overall = max(worst_df.values()) if worst_df else 0.0
    err = sum(c[1] for c in counts.values())
    print(f"\nworst relative error across all measured points: {overall:.2e}")
    if err:
        print(f"{err} point(s) returned #NUM!")
    else:
        print("no refusals: the density is evaluable across the whole probed range")
if __name__ == "__main__":
    main()
