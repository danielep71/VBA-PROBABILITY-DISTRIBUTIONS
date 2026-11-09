"""
Score the #24 Excel evidence against its frozen rules.

Part A, the holdout (gamma_lower_q_grid.csv, HOLDOUT_DESIGN.md):
  H1  every row returns a number at the echoed design inputs, survival within
      5E-15 and cumulative within 3E-15 relative (Chi-square held to the same);
  H2  the Gamma and Chi-square routes return the same Double at every point;
  H3  per shape, over increasing X, the CDF never decreases and the survival
      never increases;
  H4  abs(P + Q - 1) <= 3E-15 * P + 5E-15 * Q, exactly, at every point.

Part B, the fitting ladder (positive_ratio_subnormal/gamma_q_complement.csv):
  L1  every public CDF is within 3E-15 and every public survival within 5E-15
      of mpmath references at the binary64 inputs (Q computed directly);
  L2  the raw binary64 series P, whose source #24 does not change, is bit for
      bit the value of the pre-#24 export at e278318.

An unexported holdout (no observations) is reported as such and passes --check;
a partly exported one fails.

Run: python3 analyze_gamma_lower_q.py [--check]
"""
import argparse
import csv
import io
import os
import subprocess
import sys
from fractions import Fraction

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
sys.path.insert(0, HERE)
from design import SHAPES, points  # noqa: E402

GRID = os.path.join(HERE, "gamma_lower_q_grid.csv")
LADDER_PATH = "benchmark/positive_ratio_subnormal/gamma_q_complement.csv"
LADDER_BASELINE = "e278318"
BOUND = {"Cumulative": Fraction(3, 10**15), "Survival": Fraction(5, 10**15)}


def double(token):
    """The Double an exporter hi;lo token encodes, or None for ERROR or blank."""
    token = (token or "").strip()
    if not token or token.upper() == "ERROR":
        return None
    try:
        hi, lo = token.split(";")
        return float(mp.mpf(hi) + mp.mpf(lo))
    except (ValueError, TypeError):
        return None


def relative(value, reference):
    return abs(Fraction(value) - reference) / reference


def read(path):
    with open(path, encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f))


def part_a(path=GRID):
    rows = read(path)
    observed = [r for r in rows if r["observed_vba"].strip()]
    if not observed:
        return None
    problems = []
    if len(observed) != len(rows):
        problems.append(f"partly exported: {len(observed)} of {len(rows)} rows observed")
    values = {}
    worst = {"Cumulative": Fraction(0), "Survival": Fraction(0)}
    for r in rows:
        kind = "Survival" if r["function"].endswith("Survival") else "Cumulative"
        label = f"{r['function']}({r['arg1']}, {r['arg2']})"
        v = double(r["observed_vba"])
        if v is None:
            problems.append(f"H1 {label}: no numeric result ({r['observed_vba'] or 'blank'})")
            continue
        if (double(r["echo_arg1"]), double(r["echo_arg2"])) != (float(r["arg1"]), float(r["arg2"])):
            problems.append(f"H1 {label}: echoed inputs differ from the design")
        err = relative(v, Fraction(r["reference"]))
        worst[kind] = max(worst[kind], err)
        if err > BOUND[kind]:
            problems.append(f"H1 {label}: relative error {float(err):.3e} exceeds {float(BOUND[kind]):.0e}")
        route = "Gamma" if r["function"].startswith("Gamma") else "ChiSquare"
        x = float(r["arg1"]) if route == "Gamma" else float(r["arg1"]) / 2
        a = float(r["arg2"]) if route == "Gamma" else float(r["arg2"]) / 2
        values[(route, kind, a, x)] = v
    pairs = identical = 0
    for (route, kind, a, x), v in values.items():
        if route != "Gamma":
            continue
        other = values.get(("ChiSquare", kind, a, x))
        if other is None:
            continue
        pairs += 1
        if other == v:
            identical += 1
        else:
            problems.append(f"H2 {kind} at A={a!r}, X={x!r}: Gamma {v!r} != Chi-square {other!r}")
    for route in ("Gamma", "ChiSquare"):
        for a in SHAPES:
            xs = sorted(points(a))
            for kind, sign in (("Cumulative", 1), ("Survival", -1)):
                seq = [values.get((route, kind, a, x)) for x in xs]
                for (x0, v0), (x1, v1) in zip(zip(xs, seq), zip(xs[1:], seq[1:])):
                    if v0 is not None and v1 is not None and sign * (v1 - v0) < 0:
                        problems.append(f"H3 {route} {kind} A={a!r}: {v0!r} at X={x0!r} "
                                        f"then {v1!r} at X={x1!r}")
            for x in xs:
                p = values.get((route, "Cumulative", a, x))
                q = values.get((route, "Survival", a, x))
                if p is None or q is None:
                    continue
                gap = abs(Fraction(p) + Fraction(q) - 1)
                if gap > BOUND["Cumulative"] * Fraction(p) + BOUND["Survival"] * Fraction(q):
                    problems.append(f"H4 {route} A={a!r}, X={x!r}: P + Q - 1 = {float(gap):.3e}")
    return {"pass": not problems, "problems": problems, "rows": len(rows),
            "route_pairs": pairs, "route_pairs_bit_identical": identical,
            "worst_cumulative": float(worst["Cumulative"]),
            "worst_survival": float(worst["Survival"])}


def git_show(rev, path):
    return subprocess.check_output(["git", "show", f"{rev}:{path}"], cwd=ROOT, text=True)


def part_b(text=None, baseline=None):
    text = text if text is not None else open(os.path.join(ROOT, LADDER_PATH), encoding="utf-8").read()
    baseline = baseline if baseline is not None else git_show(LADDER_BASELINE, LADDER_PATH)
    rows = list(csv.DictReader(io.StringIO(text)))
    before = {r["point_id"]: r for r in csv.DictReader(io.StringIO(baseline))}
    problems = []
    worst = {"cdf": Fraction(0), "survival": Fraction(0)}
    for r in rows:
        pid = r["point_id"]
        x, a = double(r["x"]), double(r["shape"])
        with mp.workdps(80):
            ref_p = Fraction(str(mp.gammainc(mp.mpf(a), 0, mp.mpf(x), regularized=True)))
            ref_q = Fraction(str(mp.gammainc(mp.mpf(a), mp.mpf(x), mp.inf, regularized=True)))
        for column, ref, bound in (("cdf", ref_p, BOUND["Cumulative"]),
                                   ("survival", ref_q, BOUND["Survival"])):
            v = double(r[column]) if r[column + "_status"] == "OK" else None
            if v is None:
                problems.append(f"L1 {pid}: {column} has no numeric result")
                continue
            err = relative(v, ref)
            worst[column] = max(worst[column], err)
            if err > bound:
                problems.append(f"L1 {pid}: {column} error {float(err):.3e} exceeds {float(bound):.0e}")
        old = before.get(pid)
        if old is None or double(old["series_p"]) != double(r["series_p"]):
            problems.append(f"L2 {pid}: raw series P differs from the {LADDER_BASELINE} export")
    return {"pass": not problems, "problems": problems, "rows": len(rows),
            "worst_cdf": float(worst["cdf"]), "worst_survival": float(worst["survival"])}


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--check", action="store_true",
                    help="exit 0 for an unexported holdout; otherwise both parts must pass")
    ap.parse_args(argv)
    holdout = part_a()
    if holdout is None:
        print("Part A: the #24 holdout has no observations yet (not exported)")
        return 0
    ladder = part_b()
    for name, result in (("Part A (holdout)", holdout), ("Part B (ladder)", ladder)):
        state = "PASS" if result["pass"] else "FAIL"
        detail = {k: v for k, v in result.items() if k not in ("pass", "problems")}
        print(f"{name}: {state} {detail}")
        for p in result["problems"][:40]:
            print("  - " + p)
    return 0 if holdout["pass"] and ladder["pass"] else 1


if __name__ == "__main__":
    sys.exit(main())
