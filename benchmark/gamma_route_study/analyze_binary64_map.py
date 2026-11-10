"""
Score the #307 binary64 route map (MAP_DESIGN.md) against the contracts.

At every point it forms what the pre-#242 adapter returned: below X = A + 1,
P from the lower series and Q = 1 - P in binary64; from A + 1, Q from the
continued fraction and P = 1 - Q. It also scores each routine alone wherever
it converged. Each point is classed against the regime that governs it
(all, large_shape, extreme_shape); a reference below 1E-300 is reported as
underflow, not scored. The result is exploratory evidence for a routing rule,
never a validation of one.

Run:
    python3 analyze_binary64_map.py            (print the summary)
    python3 analyze_binary64_map.py --write    (also write map_results.json)
    python3 analyze_binary64_map.py --check    (exit 0 for an unexported grid)
"""
import argparse
import csv
import json
import math
import os
import sys
from fractions import Fraction

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from design import SHAPES  # noqa: E402

GRID = os.path.join(HERE, "binary64_map_grid.csv")
RESULTS = os.path.join(HERE, "map_results.json")
BOUNDS = {"all": (Fraction(3, 10**15), Fraction(5, 10**15)),
          "large_shape": (Fraction(2, 10**13), Fraction(2, 10**13)),
          "extreme_shape": (Fraction(2, 10**12), Fraction(5, 10**13))}
UNDERFLOW = Fraction(1, 10**300)


def double(token):
    token = (token or "").strip()
    if not token or token.upper() in ("ERROR", "FAIL"):
        return None
    hi, lo = token.split(";")
    return float(mp.mpf(hi) + mp.mpf(lo))


def regime(a, x):
    within = abs(x - a) <= math.sqrt(a)
    if 1e2 <= a <= 1e6 and within:
        return "large_shape"
    if a > 1e6 and within:
        return "extreme_shape"
    return "all"


def relative(value, reference):
    if value is None:
        return None
    if reference < UNDERFLOW:
        return "underflow"
    return float(abs(Fraction(value) - reference) / reference)


def score(rows):
    observed = [r for r in rows if r["series_status"].strip()]
    if not observed:
        return None
    problems = []
    if len(observed) != len(rows):
        problems.append(f"partly exported: {len(observed)} of {len(rows)} rows observed")
    points = []
    for r in rows:
        a, x = float(r["shape"]), float(r["x"])
        if (double(r["echo_shape"]), double(r["echo_x"])) != (a, x):
            problems.append(f"{r['point_id']}: echoed inputs differ from the design")
        ref_p, ref_q = Fraction(r["ref_p"]), Fraction(r["ref_q"])
        series = double(r["series_p"]) if r["series_status"] == "OK" else None
        cf = double(r["cf_q"]) if r["cf_status"] == "OK" else None
        lower = x < a + 1.0
        if lower:
            route_p = series
            route_q = None if series is None else min(max(1.0 - series, 0.0), 1.0)
        else:
            route_q = cf
            route_p = None if cf is None else min(max(1.0 - cf, 0.0), 1.0)
        reg = regime(a, x)
        bound_p, bound_q = BOUNDS[reg]
        err_p, err_q = relative(route_p, ref_p), relative(route_q, ref_q)
        meets = all(e == "underflow" or (e is not None and e <= float(b))
                    for e, b in ((err_p, bound_p), (err_q, bound_q)))
        points.append({"point_id": r["point_id"], "shape": a, "x": x, "branch":
                       "series" if lower else "cf", "regime": reg,
                       "route_err_p": err_p, "route_err_q": err_q, "meets": meets,
                       "series_status": r["series_status"], "series_err_p": relative(series, ref_p),
                       "cf_status": r["cf_status"], "cf_err_q": relative(cf, ref_q),
                       "series_note": r["series_note"], "cf_note": r["cf_note"]})
    return {"problems": problems, "points": points}


def summary(result):
    lines = []
    for a in SHAPES:
        pts = [p for p in result["points"] if p["shape"] == a]
        bad = [p for p in pts if not p["meets"]]

        def worst(key):
            vals = [p[key] for p in pts if isinstance(p[key], float)]
            return max(vals) if vals else float("nan")
        failing = ", ".join(format(p["x"], ".6g") for p in bad)
        lines.append(f"  A={a:<8g} {len(pts) - len(bad):2d}/{len(pts):2d} meet   "
                     f"worst P {worst('route_err_p'):9.2e}  worst Q {worst('route_err_q'):9.2e}"
                     + (f"   fail at X={failing}" if bad else ""))
    return lines


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--write", action="store_true")
    ap.add_argument("--check", action="store_true")
    a = ap.parse_args(argv)
    with open(GRID, encoding="utf-8", newline="") as f:
        rows = list(csv.DictReader(f))
    result = score(rows)
    if result is None:
        print("The #307 binary64 route map has no observations yet (not exported)")
        return 0
    meets = sum(p["meets"] for p in result["points"])
    print(f"#307 binary64 route map: {meets} of {len(result['points'])} points meet "
          "their governing contract on the pre-#242 route")
    for line in summary(result):
        print(line)
    for p in result["problems"]:
        print("  PROBLEM: " + p)
    if a.write:
        with open(RESULTS, "w", encoding="utf-8") as f:
            json.dump(result, f, indent=1)
            f.write("\n")
    return 1 if result["problems"] else 0


if __name__ == "__main__":
    sys.exit(main())
