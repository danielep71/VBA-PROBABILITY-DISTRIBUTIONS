"""
#330: the grid Excel evaluates for the function map (function_map.py).

One file, excel_function_map_grid.csv, one row per example and per probe:

    point_id, kind, excel_formula, kstats_call, reference,
    observed_excel, observed_kstats

Excel fills the last two with M_STATS_PROBDIST_XFM.bas: the first is what
Excel's own function returns, the second what the library returns, each as
the 16 hexadecimal digits of its Double, or ERROR. The formulas are stored
with ";" between arguments, because the file is comma-separated; the
exporter puts the commas back. reference is the library expression from
60-digit arithmetic (references.py), to 25 digits.

Run:
    python3 build_grid.py --write   (only when the file is absent)
    python3 build_grid.py --check   (the committed file is the design)
"""
import argparse
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import function_map  # noqa: E402
import references  # noqa: E402

FILE = os.path.join(HERE, "excel_function_map_grid.csv")
HEADER = ["point_id", "kind", "excel_formula", "kstats_call", "reference",
          "observed_excel", "observed_kstats"]
KEY_COLUMNS = 5


def stored(formula):
    if ";" in formula:
        raise ValueError(f"a formula may not hold a semicolon: {formula}")
    return formula.replace(",", ";")


def rows():
    out = [("example", *row["example"]) for row in function_map.ROWS]
    out += [("probe", probe["excel"], probe["kstats"]) for probe in function_map.PROBES]
    return out


def render():
    lines = [",".join(HEADER)]
    counts = {}
    for kind, excel, kstats in rows():
        n = counts.get(kind, 0)
        counts[kind] = n + 1
        reference = mp.nstr(references.evaluate(kstats), 25)
        lines.append(",".join([f"{kind[:2]}{n:03d}", kind, stored(excel), stored(kstats),
                               reference, "", ""]))
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
        print(f"refusing: {FILE} exists; it may hold observations. Use --check.")
        return 1
    else:
        with open(FILE, "w", encoding="utf-8", newline="") as handle:
            handle.write(text)
    print(f"{'PASS' if a.check else 'wrote'}: {len(identity(text)) - 1} rows")
    return 0


if __name__ == "__main__":
    sys.exit(main())
