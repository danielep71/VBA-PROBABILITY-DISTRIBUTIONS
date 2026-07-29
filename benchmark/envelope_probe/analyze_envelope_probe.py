"""
Report where each family's kernel stops being accurate, to set the StudentT,
ChiSquare and F envelopes from measurement.

Run after Export_EnvelopeProbe has filled observed_vba.
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


def family_and_df(regime):
    fam = regime.split("_")[0]
    m = re.search(r"df([0-9.e+]+)", regime)
    return fam, (float(m.group(1)) if m else float("nan"))


def main():
    rows = list(csv.DictReader(open("envelope_probe_grid.csv", encoding="utf-8")))
    worst = defaultdict(float)
    counts = defaultdict(lambda: [0, 0])
    for r in rows:
        fam, df = family_and_df(r["regime"])
        key = (fam, df)
        obs = observed(r["observed_vba"])
        if obs is None:
            counts[key][1] += 1
            continue
        ref = float(r["reference"])
        counts[key][0] += 1
        if ref != 0:
            worst[key] = max(worst[key], abs(obs - ref) / abs(ref))

    caps = {"chi": 1e6, "t": 1e6, "f": 1e5}
    print(f"{'family':<8}{'df':>10}{'worst rel':>13}{'n':>4}{'refused':>9}  vs current cap")
    for key in sorted(counts, key=lambda k: (k[0], k[1])):
        fam, df = key
        ok, bad = counts[key]
        w = f"{worst[key]:.2e}" if ok else "n/a"
        mark = "ABOVE cap" if df > caps.get(fam, 0) else "at/below cap"
        print(f"{fam:<8}{df:>10.0e}{w:>13}{ok:>4}{bad:>9}  {mark}")

    print("\nHighest df measured accurately per family (candidate cap):")
    for fam in sorted(caps):
        good = [df for (f, df), w in worst.items()
                if f == fam and counts[(f, df)][1] == 0]
        if good:
            print(f"  {fam:<5} current cap {caps[fam]:.0e} -> measured clean to {max(good):.0e}"
                  f"  (worst {max(w for (f, d), w in worst.items() if f == fam):.2e})")
        else:
            print(f"  {fam:<5} no clean df measured")
    print("\nA cap may only be raised to a df that was MEASURED here, never beyond.")


if __name__ == "__main__":
    main()
