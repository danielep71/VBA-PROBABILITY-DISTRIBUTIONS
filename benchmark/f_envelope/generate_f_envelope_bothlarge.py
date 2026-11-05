"""
F envelope study - BOTH-LARGE orientation.

The one-large sweep (one df large, the other ~1) is not the worst case. When BOTH
incomplete-beta parameters are large the continued fraction degrades earlier (an
existing test already used a loose 1E-6 tolerance for F_cdf(1, 1E5, 1E5)). This
sweep measures d1 = d2 = df and d1 = df, d2 = 3*df across a df range, placing x at
fixed CDF levels via the (fast) beta inverse so no reference is degenerate.
The strict F envelope must be the MINIMUM boundary over all orientations.

References are the 50-digit continued-fraction incomplete beta at the exact
binary64 argument each row stores, with the survival computed directly (#159;
see _fenv_common.py). Before #159 they were taken at 40 digits and at the
17-digit decimal of x rather than at the Double it parses to.

The design points come from SciPy's betaincinv. The committed grid holds Excel
observations; re-reference it, changing reference cells only, with

    python generate_f_envelope_bothlarge.py --merge-references          # report
    python generate_f_envelope_bothlarge.py --merge-references --write  # patch

When SciPy is installed the merge also proves that the design reproduces the
committed argument tokens; without it the merge re-references the committed
tokens, which are the arguments the observations were measured at.
"""
import argparse, csv, os
import mpmath as mp
from _fenv_common import FIELDS, HERE, merge, rows_for
mp.mp.dps = 50

DF_VALUES = [100, 300, 1e3, 3e3, 1e4, 3e4, 1e5, 3e5, 1e6]
PROBS = [0.1, 0.25, 0.5, 0.75, 0.9]

def design():
    """(x token, d1, d2, tag) for every point, from SciPy's beta inverse."""
    from scipy.special import betaincinv
    pts = []
    for df in DF_VALUES:
        for d1, d2, tag in ((df, df, "botheq"), (df, 3 * df, "bothbig")):
            a, b = d1 / 2, d2 / 2
            for p in PROBS:
                y = betaincinv(a, b, p)
                x = d2 * y / (d1 * (1 - y))
                pts.append((mp.nstr(mp.mpf(x), 17), d1, d2, tag))
    return pts

def committed_design(path):
    """The same points, read back from the committed grid."""
    pts = []
    with open(path, newline="", encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            if r["function"] == "F_Cumulative":
                pts.append((r["arg1"], float(r["arg2"]), float(r["arg3"]),
                            r["evidence_set"].split(":", 1)[1]))
    return pts

def build(points):
    rows = []
    for xs, d1, d2, tag in points:
        rows += rows_for(xs, d1, d2, tag)
    return rows

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--out", default="f_envelope_bothlarge_grid.csv")
    ap.add_argument("--merge-references", action="store_true",
                    help="re-reference the committed grid, keeping every observation")
    ap.add_argument("--write", action="store_true",
                    help="with --merge-references: patch the committed grid")
    a = ap.parse_args()
    grid = os.path.join(HERE, "f_envelope_bothlarge_grid.csv")
    if a.merge_references:
        try:
            points = design()
            print("design points regenerated with SciPy")
        except ImportError:
            points = committed_design(grid)
            print("SciPy not installed: re-referencing the committed argument tokens")
        merge("f_envelope_bothlarge_grid.csv", build(points), a.write)
        return
    rows = build(design())
    with open(a.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS, lineterminator="\n"); w.writeheader(); w.writerows(rows)
    print(f"wrote {a.out}: {len(rows)} rows")

if __name__ == "__main__":
    main()
