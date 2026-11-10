"""
Build the #307 binary64 route map grid from the frozen design, with references.

Columns (13): point_id, shape, x, ref_p, ref_q, echo_shape, echo_x,
series_status, series_p, cf_status, cf_q, cf_note, series_note. The first five
are the design and its references; the exporter (M_STATS_PROBDIST_GBR.bas)
writes the rest. References come from benchmark/_igamma.py: the tail on the
converging side directly, the other as its high-precision complement, at 40
and 60 digits, agreeing to 1E-30 relative.

Run:
    python3 generate_binary64_map.py --write   (only when the grid is absent)
    python3 generate_binary64_map.py --check   (the committed grid is the design)
"""
import argparse
import csv
import io
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.dirname(HERE))
from design import SHAPES, points  # noqa: E402
import _igamma  # noqa: E402

GRID = os.path.join(HERE, "binary64_map_grid.csv")
FIELDS = ["point_id", "shape", "x", "ref_p", "ref_q", "echo_shape", "echo_x",
          "series_status", "series_p", "cf_status", "cf_q", "cf_note", "series_note"]
KEY_COLUMNS = 5


def tails(shape, x, dps):
    with mp.workdps(dps):
        value, side, _route, _n = _igamma.reg_incomplete(mp.mpf(shape), mp.mpf(x))
        return (value, 1 - value) if side == "P" else (1 - value, value)


def references(shape, x):
    (p40, q40), (p60, q60) = tails(shape, x, 40), tails(shape, x, 60)
    with mp.workdps(60):
        for low, high in ((p40, p60), (q40, q60)):
            if high != 0 and abs(low - high) > mp.mpf("1E-30") * abs(high):
                raise ValueError(f"unstable reference at A={shape!r}, X={x!r}")
        return mp.nstr(p60, 40), mp.nstr(q60, 40)


def build():
    rows = []
    for i, shape in enumerate(SHAPES):
        for j, x in enumerate(points(shape)):
            p, q = references(shape, x)
            rows.append({"point_id": f"s{i:02d}_x{j:02d}", "shape": repr(shape), "x": repr(x),
                         "ref_p": p, "ref_q": q, **{k: "" for k in FIELDS[5:]}})
    return rows


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
            print("FAIL: the map's identity columns differ from a fresh build of the design")
            return 1
        print(f"PASS: {len(identity(text)) - 1} rows match the frozen #307 map design")
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
