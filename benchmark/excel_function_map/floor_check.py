"""
#493: judge the exported Excel floor grid (floor_build.py).

At each point the library's relative error is compared with that of the most
accurate Excel formulation that returned a number. The library is INFERIOR at
a point when Excel's error is smaller and the library's is above FLOOR
(CONTRIBUTING.md, "Excel as the floor"). For the discrete inverse, whose
result is a count, the library is inferior when it misses the reference and
Excel does not.

A point is NOT COMPARED when no Excel formulation returned a number. A library
side that returned ERROR where Excel returned a number is reported apart: it
is not an accuracy comparison, and it says so.

    python floor_check.py            report; write floor_results.json; exit 0 when nowhere inferior
    python floor_check.py --grid G   judge another copy of the grid (an artifact)
"""
import argparse
import csv
import json
import os
import struct
import sys
from collections import defaultdict

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import floor_build as B  # noqa: E402

FLOOR = 1e-14
OUT = os.path.join(HERE, "floor_results.json")


def value(bits):
    """The Double of a 16-hex-digit observation, or None for ERROR or blank."""
    bits = (bits or "").strip()
    if len(bits) != 16:
        return None
    return struct.unpack(">d", bytes.fromhex(bits))[0]


def rel(v, ref):
    with mp.workdps(60):
        ref = mp.mpf(ref)
        return float(abs((mp.mpf(v) - ref) / ref))


def judge(rows):
    """(points, blank) from the grid rows; points holds one verdict per point."""
    groups, blank = defaultdict(list), 0
    for r in rows:
        if not r["observed_kstats"].strip() and not r["observed_excel"].strip():
            blank += 1
        groups[r["kind"]].append(r)
    points = []
    for kind, group in groups.items():
        source, line, fn = kind.split(":")
        ref = group[0]["reference"]
        lib = value(group[0]["observed_kstats"])
        excel = []
        for r in group:
            v = value(r["observed_excel"])
            if v is not None:
                excel.append((r["excel_formula"], v))
        p = {"source": source, "line": int(line), "function": fn,
             "kstats_call": group[0]["kstats_call"], "reference": ref}
        if not excel:
            p["verdict"] = "not compared"
            points.append(p)
            continue
        if fn in B.DISCRETE_INVERSE:
            target = float(ref)
            best = next(((f, v) for f, v in excel if v == target), excel[0])
            p.update(excel_formula=best[0], excel=best[1], library=lib)
            lib_ok, xl_ok = lib == target, best[1] == target
            p["verdict"] = ("library error" if lib is None else
                            "inferior" if (xl_ok and not lib_ok) else "not inferior")
            points.append(p)
            continue
        errors = [(rel(v, ref), f, v) for f, v in excel]
        e_xl, form, v_xl = min(errors)
        p.update(excel_formula=form, excel=v_xl, excel_rel=e_xl, library=lib)
        if lib is None:
            p["verdict"] = "library error"
        else:
            e_lib = rel(lib, ref)
            p["library_rel"] = e_lib
            p["verdict"] = "inferior" if (e_xl < e_lib and e_lib > FLOOR) else "not inferior"
        points.append(p)
    return points, blank


def main(argv=None):
    ap = argparse.ArgumentParser(description=(__doc__ or "").split("\n\n")[0])
    ap.add_argument("--grid", default=B.GRID)
    ap.add_argument("--out", default=OUT)
    a = ap.parse_args(argv)
    with open(a.grid, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    points, blank = judge(rows)
    if blank:
        print(f"FAIL: {blank} row(s) hold no observation: export the grid first")
        return 1
    count = defaultdict(int)
    by_function = defaultdict(lambda: defaultdict(int))
    for p in points:
        count[p["verdict"]] += 1
        by_function[p["function"]][p["verdict"]] += 1
    inferior = sorted((p for p in points if p["verdict"] == "inferior"),
                      key=lambda p: -p.get("library_rel", 1.0))
    result = {"issue": "#493", "floor": FLOOR, "points": len(points), "verdicts": dict(count),
              "by_function": {k: dict(v) for k, v in sorted(by_function.items())},
              "inferior": inferior,
              "library_error": [p for p in points if p["verdict"] == "library error"],
              "not_compared": [p for p in points if p["verdict"] == "not compared"]}
    with open(a.out, "w", encoding="utf-8", newline="\n") as f:
        json.dump(result, f, indent=1)
        f.write("\n")
    print(f"  points                {len(points)}")
    for k in ("not inferior", "inferior", "library error", "not compared"):
        print(f"  {k:<21} {count.get(k, 0)}")
    for fn, v in sorted(by_function.items()):
        if v.get("inferior") or v.get("library error"):
            print(f"    {fn:<38} inferior {v.get('inferior', 0):>3} of "
                  f"{sum(v.values()):>3}; library error {v.get('library error', 0)}")
    if inferior or count.get("library error"):
        print(f"INFERIOR TO EXCEL at {len(inferior)} point(s); "
              f"library error at {count.get('library error', 0)}")
        return 1
    print(f"PASS: the library is nowhere inferior to Excel (floor {FLOOR:g})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
