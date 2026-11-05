"""
Generator for survival_boundary_grid.csv (#159).

The grid locates where the standard-normal survival core stops meeting the
tight relative contracts (survival_boundary_summary.md). It was built in July
2026 without a committed generator; this script reproduces its rows and
evaluates every reference directly as an upper tail, erfc(z / sqrt 2) / 2, at
the exact binary64 argument each row stores:

  NormalStandard_Survival(z)      z on the boundary ladder 1.5 .. 6
  Normal_Survival(x, 0, 1)        x = 2.5, 3.5, 5
  Lognormal_Survival(x, 0, 1)     x = the Double nearest e^2.5 and e^3.5

The two Lognormal references were first taken at the design value z = 2.5 and
3.5, not at log(x) for the Double x the exporter passes, which kept only about
16 of their 30 digits. This generator evaluates them at log(x).

The committed grid holds Excel observations, and the #72 normal-tail record
(normal_tail_study/results.json) binds it and records each row's error against
its committed reference. Re-reference it, changing reference cells only, with

    python generate_survival_boundary.py --merge-references          # report
    python generate_survival_boundary.py --merge-references --write  # patch

The merge holds back every row that record reads, which is every row of this
grid, so it reports what would change and writes nothing until that study's
record is reconciled.
"""
import argparse
import csv
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
from _binary64 import exact_binary64_mpf as D  # noqa: E402

mp.mp.dps = 50
FIELDS = ["function", "vba_kernel", "claim", "metric", "arg1", "arg2", "arg3",
          "reference", "observed_vba", "regime", "evidence_set"]

Z_LADDER = [1.5, 2.0, 2.25, 2.5, 2.75, 3.0, 3.25, 3.5, 3.75, 4.0, 4.5, 5.0, 5.5, 6.0]
NORMAL_X = [2.5, 3.5, 5.0]
LOGNORMAL_Z = [2.5, 3.5]


def q(z):
    """Standard-normal upper tail, direct."""
    return mp.erfc(z / mp.sqrt(2)) / 2


def row(fn, args, ref):
    """args are tokens or numbers; numbers are written with 17 digits."""
    toks = [a if isinstance(a, str) else mp.nstr(mp.mpf(a), 17) for a in args]
    toks += [""] * (3 - len(toks))
    return {"function": fn, "vba_kernel": f"K_STATS_{fn}", "claim": "boundary",
            "metric": "rel", "arg1": toks[0], "arg2": toks[1], "arg3": toks[2],
            "reference": mp.nstr(ref, 30), "observed_vba": "", "regime": "all",
            "evidence_set": "survival_boundary"}


def build():
    rows = []
    for z in Z_LADDER:
        rows.append(row("NormalStandard_Survival", [z], q(D(mp.nstr(mp.mpf(z), 17)))))
    for x in NORMAL_X:
        X, mu, sd = D(mp.nstr(mp.mpf(x), 17)), D("0.0"), D("1.0")
        rows.append(row("Normal_Survival", [x, 0.0, 1.0], q((X - mu) / sd)))
    for z in LOGNORMAL_Z:
        # The token is e^z written to 17 digits; the exporter passes the Double
        # it parses to, so the reference is taken at log of that Double.
        tok = mp.nstr(mp.e ** mp.mpf(z), 17)
        rows.append(row("Lognormal_Survival", [tok, 0.0, 1.0], q(mp.log(D(tok)))))
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(HERE, "survival_boundary_grid.csv"))
    ap.add_argument("--merge-references", action="store_true",
                    help="re-reference the committed grid, keeping every observation")
    ap.add_argument("--write", action="store_true",
                    help="with --merge-references: patch the committed grid")
    a = ap.parse_args()
    rows = build()
    if a.merge_references:
        from _study_merge import merge_references, print_report, study_describe, study_key

        def hold(r):
            return ("read by the #72 normal-tail record "
                    "(normal_tail_study/results.json binds this grid)")

        changes, held = merge_references(os.path.join(HERE, "survival_boundary_grid.csv"),
                                         rows, study_key, {"observed_vba"}, hold=hold,
                                         write=a.write, describe=study_describe)
        print_report("survival_boundary_grid.csv", changes, held, a.write)
        return
    with open(a.out, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=FIELDS, lineterminator="\n")
        w.writeheader()
        w.writerows(rows)
    print(f"wrote {a.out}: {len(rows)} rows")


if __name__ == "__main__":
    main()
