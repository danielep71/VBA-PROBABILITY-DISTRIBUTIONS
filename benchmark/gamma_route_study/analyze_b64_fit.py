"""
Score the #307 binary64 route fit set (FIT_DESIGN.md).

Three results, each from the exported grid alone:

  1. Fidelity. At every point, each routine's status and value equal
     route_model.py bit for bit: the series below X = A + 1, the continued
     fraction from it, and a decline where route_model declines.
  2. Accuracy (exploratory). The tails the routing rule would return: below
     A + 1, P from the series and Q = 1 - P where P <= 0.5; from A + 1, Q from
     the continued fraction and P = 1 - Q. Each against its governing contract.
  3. Cost. Microseconds per call against the iterations route_model counts,
     fitted as c0 + c1 n for each routine; and the shape ceiling A2 that
     FIT_DESIGN.md's frozen rule derives from it.

Run:
    python3 analyze_b64_fit.py            (print the summary)
    python3 analyze_b64_fit.py --write    (also write fit_results.json)
    python3 analyze_b64_fit.py --check    (exit 0 for an unexported grid)
"""
import argparse
import csv
import functools
import json
import math
import os
import sys
from fractions import Fraction

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import route_model as RM  # noqa: E402
from analyze_binary64_map import BOUNDS, UNDERFLOW, regime  # noqa: E402

GRID = os.path.join(HERE, "b64_fit_grid.csv")
RESULTS = os.path.join(HERE, "fit_results.json")
# FIT_DESIGN.md: the ladder of candidate shape ceilings and the cost budget.
LADDER = (1e3, 2e3, 3e3, 5e3, 1e4, 2e4, 3e4, 5e4, 1e5, 2e5, 5e5, 1e6)
BUDGET_US = 100.0


def double(token):
    """A "hi;lo" pair written by FormatFullPrecision, as the Double it encodes."""
    hi, lo = token.split(";")
    return float(hi) + float(lo)


@functools.lru_cache(maxsize=None)
def modelled(a, x):
    """(series, cf): each (value, iterations) or None where it declines."""
    below = RM.below_seam(a, x)
    out = []
    for routine, wanted in ((RM.series_p, below), (RM.cf_q, not below)):
        if not (a > 0 and x > 0 and wanted):
            out.append(None)
            continue
        try:
            out.append(routine(a, x))
        except RM.Declined:
            out.append(None)
    return tuple(out)


def relative(value, reference):
    if reference < UNDERFLOW:
        return "underflow"
    return float(abs(Fraction(value) - reference) / reference)


def tails(series, cf):
    """The routed P and Q from the observed routines; Q None where it would go
    to the double-double kernel (P > 0.5 below A + 1)."""
    if series is not None:
        p = series
        q = RM.sub(1.0, p) if p <= 0.5 else None
        return p, q
    if cf is not None:
        return RM.sub(1.0, cf), cf
    return None, None


def fit_line(points):
    """Least squares c0 + c1 n."""
    n = len(points)
    if n < 2:
        return None
    sx = sum(p[0] for p in points)
    sy = sum(p[1] for p in points)
    sxx = sum(p[0] * p[0] for p in points)
    sxy = sum(p[0] * p[1] for p in points)
    den = n * sxx - sx * sx
    if den == 0:
        return None
    c1 = (n * sxy - sx * sy) / den
    return (sy - c1 * sx) / n, c1


def ceiling(rows, model):
    """FIT_DESIGN.md: the largest ladder value A2 such that every fit row with
    shape <= A2 measured at most BUDGET_US, and the cost model predicts at most
    BUDGET_US at A2's worst arguments (the series at the predecessor of A2 + 1,
    the continued fraction at A2 + 1)."""
    best = None
    for a2 in LADDER:
        measured = [r["route_us"] for r in rows if r["shape"] <= a2 and r["route_us"] is not None]
        if not measured or max(measured) > BUDGET_US:
            break
        s = RM.series_p(a2, math.nextafter(a2 + 1.0, 0.0))[1]
        c = RM.cf_q(a2, a2 + 1.0)[1]
        predicted = max(model["series"][0] + model["series"][1] * s, model["cf"][0] + model["cf"][1] * c)
        if predicted > BUDGET_US:
            break
        best = a2
    return best


def score(rows):
    observed = [r for r in rows if r["series_status"].strip()]
    if not observed:
        return None
    problems = []
    if len(observed) != len(rows):
        problems.append(f"partly exported: {len(observed)} of {len(rows)} rows observed")
    points, timings = [], {"series": [], "cf": []}
    for r in rows:
        a, x = float(r["shape"]), float(r["x"])
        pid = r["point_id"]
        if (double(r["echo_shape"]), double(r["echo_x"])) != (a, x):
            problems.append(f"{pid}: echoed inputs differ from the design")
        expected = modelled(a, x)
        got = []
        for name, status, value, model in (("series", r["series_status"], r["series_p"], expected[0]),
                                           ("cf", r["cf_status"], r["cf_q"], expected[1])):
            if status == "OK":
                v = double(value)
                if model is None or model[0] != v:
                    problems.append(f"{pid}: {name} {v!r} differs from the model "
                                    f"{None if model is None else model[0]!r}")
                got.append(v)
            elif status == "DECLINED":
                if model is not None:
                    problems.append(f"{pid}: {name} declined where the model answers")
                got.append(None)
            else:
                problems.append(f"{pid}: {name} status {status!r}")
                got.append(None)
        p, q = tails(*got)
        bound_p, bound_q = BOUNDS[regime(a, x)]
        ref_p, ref_q = Fraction(r["ref_p"]), Fraction(r["ref_q"])
        err_p = None if p is None else relative(p, ref_p)
        err_q = None if q is None else relative(q, ref_q)
        meets = all(e is None or e == "underflow" or e <= float(b)
                    for e, b in ((err_p, bound_p), (err_q, bound_q)))
        answered = "series" if got[0] is not None else ("cf" if got[1] is not None else None)
        n = None
        if answered:
            n = (expected[0] if answered == "series" else expected[1])
            n = None if n is None else n[1]
        us = float(r["route_us"]) if r["route_us"].strip() else None
        if answered and us is None:
            problems.append(f"{pid}: answered but not timed")
        if answered and us is not None and n is not None:
            timings[answered].append((n, us))
        points.append({"point_id": pid, "shape": a, "x": x, "regime": regime(a, x),
                       "answered": answered, "iterations": n, "route_us": us,
                       "err_p": err_p, "err_q": err_q, "q_to_double_double": p is not None and q is None,
                       "meets": meets})
    model = {k: fit_line(v) for k, v in timings.items()}
    a2 = ceiling(points, model) if all(model.values()) else None
    return {"problems": problems, "points": points, "cost_model": model, "a2": a2}


def summary(result):
    lines = []
    shapes = sorted({p["shape"] for p in result["points"]})
    for a in shapes:
        pts = [p for p in result["points"] if p["shape"] == a]
        bad = [p for p in pts if not p["meets"]]

        def worst(key):
            vals = [p[key] for p in pts if isinstance(p[key], float)]
            return max(vals) if vals else float("nan")
        us = [p["route_us"] for p in pts if p["route_us"] is not None]
        lines.append(f"  A={a:<9g} {len(pts) - len(bad):2d}/{len(pts):2d} meet   worst P {worst('err_p'):9.2e}"
                     f"  worst Q {worst('err_q'):9.2e}   max {max(us) if us else float('nan'):8.1f} us")
    return lines


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--write", action="store_true")
    ap.add_argument("--check", action="store_true", help="the default: print and exit 1 on a problem")
    a = ap.parse_args(argv)
    with open(GRID, encoding="utf-8", newline="") as f:
        rows = list(csv.DictReader(f))
    result = score(rows)
    if result is None:
        print("The #307 binary64 route fit set has no observations yet (not exported)")
        return 0
    fidelity = [p for p in result["problems"] if "model" in p or "declined" in p or "status" in p]
    meets = sum(p["meets"] for p in result["points"])
    print(f"#307 binary64 route fit: {len(result['points'])} points; "
          f"{len(fidelity)} fidelity mismatches; {meets} meet their governing contract")
    for line in summary(result):
        print(line)
    for k, v in result["cost_model"].items():
        if v:
            print(f"  cost {k}: {v[0]:.2f} us + {v[1]:.4f} us per iteration")
    print(f"  shape ceiling A2 by the frozen rule: {result['a2']}")
    for p in result["problems"]:
        print("  PROBLEM: " + p)
    if a.write:
        with open(RESULTS, "w", encoding="utf-8") as f:
            json.dump(result, f, indent=1)
            f.write("\n")
    return 1 if result["problems"] else 0


if __name__ == "__main__":
    sys.exit(main())
