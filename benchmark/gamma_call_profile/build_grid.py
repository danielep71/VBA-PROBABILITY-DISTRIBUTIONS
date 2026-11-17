"""
#348: the rows of the Gamma call profile (README.md).

One file, gamma_call_profile.csv. Excel fills its last three columns with
M_STATS_PROBDIST_GCP.bas; the first five are the design.

  call      one part of a public Gamma call, timed alone at one of the eight
            points of the cost benchmark of #298;
  fixed     a part whose cost does not depend on the point, timed once;
  inverse   the public inverse and the solver at five shapes and two
            probabilities; the column x holds the probability.

Columns (8): point_id, kind, part, a, x, calls, us_per_call, note.

Run:
    python3 build_grid.py --write   (only when the file is absent)
    python3 build_grid.py --check   (the committed file is the design)
"""
import argparse
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
FILE = os.path.join(HERE, "gamma_call_profile.csv")
HEADER = ["point_id", "kind", "part", "a", "x", "calls", "us_per_call", "note"]
KEY_COLUMNS = 5

# The eight (shape, argument) points of ci/Run-ExcelVbaTests.ps1 (#298).
POINTS = (("0.0001", "0.5"), ("0.5", "0.25"), ("2.5", "1"), ("2.5", "6"),
          ("100", "90"), ("100", "110"), ("10000", "9900"), ("10000", "10100"))
CALL_PARTS = ("public_cdf", "public_sf", "routed_p", "routed_q", "b64_kernel", "legacy_loop",
              "kernel_p", "kernel_q", "log_pdf", "run_factor", "run_validate")
FIXED_PARTS = ("empty", "set_status", "try_divide", "classify", "is_finite", "run_pow2")
INVERSE_SHAPES = ("0.01", "0.5", "2.5", "100", "10000")
INVERSE_PROBABILITIES = ("0.1", "0.9")
INVERSE_PARTS = ("public_inv", "inv_p", "solver")


def rows():
    out = []
    for a, x in POINTS:
        parts = list(CALL_PARTS)
        # The Stirling error serves shapes above one, the small log-gamma the others.
        parts.append("stirling" if float(a) > 1.0 else "run_small_lg1p")
        out += [("call", part, a, x) for part in parts]
    out += [("fixed", part, "1", "1") for part in FIXED_PARTS]
    out += [("inverse", part, a, p) for a in INVERSE_SHAPES for p in INVERSE_PROBABILITIES
            for part in INVERSE_PARTS]
    return out


def render():
    lines = [",".join(HEADER)]
    counts = {}
    for kind, part, a, x in rows():
        n = counts.get(kind, 0)
        counts[kind] = n + 1
        lines.append(",".join([f"{kind[:2]}{n:03d}", kind, part, a, x, "", "", ""]))
    return "\n".join(lines) + "\n"


def identity(text):
    return [line.split(",")[:KEY_COLUMNS] for line in text.splitlines() if line.strip()]


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    mode = ap.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true")
    mode.add_argument("--check", action="store_true")
    a = ap.parse_args(argv)
    text = render()
    if a.check:
        with open(FILE, encoding="utf-8") as handle:
            if identity(handle.read()) != identity(text):
                print("FAIL: the identity columns of the file differ from a fresh build")
                return 1
    elif os.path.exists(FILE):
        print(f"refusing: {FILE} exists; it may hold measurements. Use --check.")
        return 1
    else:
        with open(FILE, "w", encoding="utf-8", newline="") as handle:
            handle.write(text)
    print(f"{'PASS' if a.check else 'wrote'}: {len(identity(text)) - 1} rows")
    return 0


if __name__ == "__main__":
    sys.exit(main())
