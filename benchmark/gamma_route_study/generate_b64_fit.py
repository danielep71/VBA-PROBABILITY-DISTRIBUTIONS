"""
Build the #307 binary64 route fit grid from the frozen fit design, with references.

Columns (14): point_id, shape, x, ref_p, ref_q, echo_shape, echo_x,
series_status, series_p, cf_status, cf_q, route_us, route_calls, note. The
first five are the design and its references; the exporter
(M_STATS_PROBDIST_GBF.bas) writes the rest. References as for the map
(generate_binary64_map.py): the converging tail directly, the other as its
high-precision complement, at 40 and 60 digits, agreeing to 1E-30 relative.

Run:
    python3 generate_b64_fit.py --write   (only when the grid is absent)
    python3 generate_b64_fit.py --check   (the committed grid is the design)
"""
import argparse
import csv
import io
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from fit_design import rows as design_rows  # noqa: E402
from generate_binary64_map import references  # noqa: E402

GRID = os.path.join(HERE, "b64_fit_grid.csv")
FIELDS = ["point_id", "shape", "x", "ref_p", "ref_q", "echo_shape", "echo_x", "series_status",
          "series_p", "cf_status", "cf_q", "route_us", "route_calls", "note"]
KEY_COLUMNS = 5


def build():
    out = []
    for pid, shape, x in design_rows():
        p, q = references(shape, x)
        out.append({"point_id": pid, "shape": repr(shape), "x": repr(x), "ref_p": p, "ref_q": q,
                    **{k: "" for k in FIELDS[KEY_COLUMNS:]}})
    return out


def render(rows):
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=FIELDS, lineterminator="\n")
    writer.writeheader()
    writer.writerows(rows)
    return buf.getvalue()


def identity(text):
    return [r[:KEY_COLUMNS] for r in csv.reader(io.StringIO(text))]


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    mode = ap.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true")
    mode.add_argument("--check", action="store_true")
    a = ap.parse_args(argv)
    text = render(build())
    if a.check:
        with open(GRID, encoding="utf-8") as f:
            committed = f.read()
        if identity(committed) != identity(text):
            print("FAIL: the fit grid's identity columns differ from a fresh build of the design")
            return 1
        print(f"PASS: {len(identity(text)) - 1} rows match the frozen #307 fit design")
        return 0
    if os.path.exists(GRID):
        print(f"refusing: {GRID} exists; it may hold observations. Use --check.")
        return 1
    with open(GRID, "w", encoding="utf-8", newline="") as f:
        f.write(text)
    print(f"wrote {len(identity(text)) - 1} rows to {GRID}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
