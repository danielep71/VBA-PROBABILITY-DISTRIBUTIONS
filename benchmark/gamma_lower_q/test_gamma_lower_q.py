"""
Proofs for the #24 holdout grid and its analyzer.

  1. The committed grid is exactly the frozen design: generate --check passes;
     7 shapes x 6 arguments x 4 routes; the seam points are the binary64
     neighbours of A + 1; every Chi-square row is the exact doubling of its
     Gamma row, with the same reference.
  2. The analyzer passes an export whose every result is the correctly rounded
     reference, with exact echoes, and counts all 84 route pairs identical.
  3. It fails a result 1E-14 off, an ERROR, a changed echo, a one-ulp route
     disagreement, a monotonicity reversal and a partial export.
  4. Part B fails the committed pre-#24 ladder on L1 (its survival breached
     5E-15 by up to 1.6E-9), and fails L2 when a raw series P changes.
  5. An unexported grid is reported as such.

Run: python3 test_gamma_lower_q.py   (exit 0 = pass)
"""
import csv
import math
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import analyze_gamma_lower_q as A  # noqa: E402
import generate_gamma_lower_q as G  # noqa: E402
from design import INTERIOR, SHAPES, points  # noqa: E402


def token(x):
    """The exporter's hi;lo encoding of a Double."""
    hi = float(f"{x:.14E}")
    return f"{hi:.14E};{x - hi:.14E}"


def exported(mutate=None):
    rows = []
    for r in G.build():
        r = dict(r)
        r["echo_arg1"], r["echo_arg2"] = token(float(r["arg1"])), token(float(r["arg2"]))
        r["observed_vba"] = token(float(r["reference"]))
        rows.append(r)
    if mutate:
        mutate(rows)
    return rows


def score(rows):
    fd, path = tempfile.mkstemp(suffix=".csv")
    with os.fdopen(fd, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=G.FIELDS, lineterminator="\n")
        w.writeheader()
        w.writerows(rows)
    try:
        return A.part_a(path)
    finally:
        os.remove(path)


def find(rows, function, a, x):
    for r in rows:
        route_x = float(r["arg1"]) / (2 if r["function"].startswith("Chi") else 1)
        route_a = float(r["arg2"]) / (2 if r["function"].startswith("Chi") else 1)
        if r["function"] == function and route_a == a and route_x == x:
            return r
    raise KeyError((function, a, x))


def main():
    fails = []

    def check(condition, message):
        if not condition:
            fails.append(message)

    # 1. The grid is the design.
    check(G.main(["--check"]) == 0, "generate --check fails on the committed grid")
    rows = G.build()
    check(len(rows) == 168, f"{len(rows)} rows, not 168")
    for a in SHAPES:
        lo, mid, hi = points(a)[len(INTERIOR):]
        check(mid == a + 1.0 and math.nextafter(mid, 0.0) == lo and math.nextafter(mid, 2.0) == hi,
              f"seam points of A={a!r} are not the binary64 neighbours of A + 1")
    for g, c in zip(rows[0::2], rows[1::2]):
        check(c["function"] == g["function"].replace("Gamma_", "ChiSquare_"),
              "routes are not paired Gamma, Chi-square")
        check(float(c["arg1"]) == 2 * float(g["arg1"]) and float(c["arg2"]) == 2 * float(g["arg2"]),
              f"{c['arg1']}, {c['arg2']} is not the exact doubling of {g['arg1']}, {g['arg2']}")
        check(c["reference"] == g["reference"] and c["claim"] == g["claim"],
              "paired routes differ in reference or claim")

    # 2. A correctly rounded export passes.
    result = score(exported())
    check(result is not None and result["pass"], f"a correct export fails: {result and result['problems'][:3]}")
    check(result and result["route_pairs"] == 84 and result["route_pairs_bit_identical"] == 84,
          "route pairs not all identical on a correct export")

    # 3. Failures are caught, each by its rule.
    a0, x0 = SHAPES[0], INTERIOR[2]

    def off(rs):
        r = find(rs, "Gamma_Survival", a0, x0)
        r["observed_vba"] = token(float(r["reference"]) * (1 - 1e-14))

    def error(rs):
        find(rs, "ChiSquare_Cumulative", a0, x0)["observed_vba"] = "ERROR"

    def echo(rs):
        r = find(rs, "Gamma_Survival", a0, x0)
        r["echo_arg2"] = token(float(r["arg2"]) * (1 + 2 ** -52))

    def route(rs):
        r = find(rs, "ChiSquare_Cumulative", SHAPES[3], INTERIOR[1])
        r["observed_vba"] = token(math.nextafter(float(r["reference"]), 0.0))

    def reverse(rs):
        first = find(rs, "Gamma_Survival", SHAPES[2], INTERIOR[0])
        second = find(rs, "Gamma_Survival", SHAPES[2], INTERIOR[1])
        second["observed_vba"] = token(math.nextafter(A.double(first["observed_vba"]), 1.0))

    def partial(rs):
        rs[5]["observed_vba"] = ""

    for name, mutate, needle in (("1E-14 off", off, "H1 Gamma_Survival"),
                                 ("ERROR", error, "no numeric"),
                                 ("echo", echo, "echoed inputs"),
                                 ("route", route, "H2 Cumulative"),
                                 ("reversal", reverse, "H3 Gamma Survival"),
                                 ("partial", partial, "partly exported")):
        result = score(exported(mutate))
        check(result is not None and not result["pass"]
              and any(needle in p for p in result["problems"]),
              f"{name}: not caught ({result and result['problems'][:3]})")

    # 4. Part B: the committed pre-#24 ladder breaches L1, and L2 sees a moved series.
    committed = open(os.path.join(A.ROOT, A.LADDER_PATH), encoding="utf-8").read()
    baseline = A.git_show(A.LADDER_BASELINE, A.LADDER_PATH)
    before = A.part_b(committed, baseline)
    check(not before["pass"] and before["worst_survival"] > 1e-9
          and all(p.startswith("L1") for p in before["problems"]),
          "Part B must fail the pre-#24 ladder on L1 only")
    lines = committed.splitlines(keepends=True)
    header = lines[0].rstrip("\n").split(",")
    col = header.index("series_p")
    cols = lines[1].rstrip("\n").split(",")
    hi, lo = cols[col].split(";")
    cols[col] = token(math.nextafter(float(hi) + float(lo), 0.0))
    lines[1] = ",".join(cols) + "\n"
    moved = A.part_b("".join(lines), baseline)
    check(any(p.startswith("L2 ") for p in moved["problems"]), "L2: a moved raw series P is not caught")

    # 5. The committed grid is unexported or fully exported.
    with open(G.GRID, encoding="utf-8") as f:
        committed_rows = list(csv.DictReader(f))
    observed = [r for r in committed_rows if r["observed_vba"].strip()]
    check(len(observed) in (0, len(committed_rows)), "the committed grid is partly observed")
    if not observed:
        check(A.part_a() is None, "an unexported grid is not reported as unexported")

    if fails:
        print("FAIL: #24 holdout grid and analyzer")
        for f in fails:
            print("  - " + f)
        return 1
    print("PASS: #24 holdout grid is the frozen design (168 rows); the analyzer passes a correct "
          "export and catches H1-H3, partial exports, and L1/L2 on the ladder")
    return 0


if __name__ == "__main__":
    sys.exit(main())
