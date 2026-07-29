"""
Report the measured accuracy of the large-shape cumulative and survival
probabilities, per family and shape, from cdf_large_shape_grid.csv.

Run after Export_CdfLargeShape has filled observed_vba.
"""
import csv
from collections import defaultdict
from decimal import Decimal


def observed(text):
    text = text.strip()
    if not text or text.upper() == "ERROR":
        return None
    return float(sum(Decimal(part) for part in text.split(";")))


def main():
    rows = list(csv.DictReader(open("cdf_large_shape_grid.csv", encoding="utf-8")))
    worst = defaultdict(float)
    counts = defaultdict(lambda: [0, 0])       # [measured, errored]
    for r in rows:
        key = (r["function"], float(r["arg2"]))
        obs = observed(r["observed_vba"])
        if obs is None:
            counts[key][1] += 1
            continue
        ref = float(r["reference"])
        counts[key][0] += 1
        if ref != 0:
            worst[key] = max(worst[key], abs(obs - ref) / abs(ref))

    print(f"{'function':<20}{'shape':>10}{'worst rel':>13}{'n':>5}{'#NUM!':>7}")
    for key in sorted(counts):
        fn, shape = key
        measured, errored = counts[key]
        w = f"{worst[key]:.2e}" if measured else "n/a"
        print(f"{fn:<20}{shape:>10.0e}{w:>13}{measured:>5}{errored:>7}")

    overall = max(worst.values()) if worst else 0.0
    err = sum(c[1] for c in counts.values())
    print(f"\nworst relative error across all measured points: {overall:.2e}")
    if err:
        print(f"{err} point(s) returned #NUM! - the kernel refused rather than "
              f"returning an unconverged value; that is the validated-domain edge.")


if __name__ == "__main__":
    main()
