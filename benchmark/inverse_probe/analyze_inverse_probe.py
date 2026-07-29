"""
Report where each inverse kernel converges and how accurately, to set the
StudentT, ChiSquare and F INVERSE envelopes from measurement.

Run after Export_InverseProbe has filled observed_vba.
"""
import csv
import re
from collections import defaultdict
from decimal import Decimal

CAPS = {"chi": 1e6, "t": 1e6, "bal": 1e5, "unb": 1e5}      # current inverse caps
LABEL = {"chi": "ChiSquare (gamma inv)", "t": "StudentT (beta inv, B=1/2)",
         "bal": "F balanced", "unb": "F unbalanced"}


def observed(text):
    text = text.strip()
    if not text or text.upper() == "ERROR":
        return None
    return float(sum(Decimal(part) for part in text.split(";")))


def main():
    rows = list(csv.DictReader(open("inverse_probe_grid.csv", encoding="utf-8")))
    worst = defaultdict(float)
    counts = defaultdict(lambda: [0, 0])
    shapes = defaultdict(set)

    for r in rows:
        tag = r["regime"].split("_")[0]
        a = float(r["arg3"]) if r["arg3"].strip() else 0.0
        b = float(r["arg4"]) if r["arg4"].strip() else 0.0
        key = (tag, a, b)
        shapes[tag].add((a, b))
        obs = observed(r["observed_vba"])
        if obs is None:
            counts[key][1] += 1
            continue
        ref = float(r["reference"])
        counts[key][0] += 1
        if ref != 0:
            worst[key] = max(worst[key], abs(obs - ref) / abs(ref))

    print(f"{'group':<26}{'A':>10}{'B':>10}{'worst rel':>13}{'n':>4}{'refused':>9}")
    for key in sorted(counts, key=lambda k: (k[0], k[1], k[2])):
        tag, a, b = key
        ok, bad = counts[key]
        w = f"{worst[key]:.2e}" if ok else "n/a"
        print(f"{LABEL[tag]:<26}{a:>10.0e}{b:>10.0e}{w:>13}{ok:>4}{bad:>9}")

    print("\nPer group: highest shape with NO refusals and every point measured")
    for tag in sorted(shapes):
        clean = [(a, b) for (a, b) in shapes[tag]
                 if counts[(tag, a, b)][1] == 0 and counts[(tag, a, b)][0] > 0]
        if clean:
            best = max(clean, key=lambda ab: max(ab))
            w = max(worst[(tag, a, b)] for a, b in clean)
            print(f"  {LABEL[tag]:<26} clean to A={best[0]:.0e} B={best[1]:.0e}"
                  f"  (worst {w:.2e})")
        else:
            print(f"  {LABEL[tag]:<26} no fully clean shape")

    refused = sum(c[1] for c in counts.values())
    if refused:
        print(f"\n{refused} point(s) refused. A refusal is a clean #NUM!, never a wrong "
              f"value, but it caps the usable domain regardless of accuracy elsewhere.")
    print("\nAn inverse cap may only be raised to a shape that was MEASURED clean here.")


if __name__ == "__main__":
    main()
