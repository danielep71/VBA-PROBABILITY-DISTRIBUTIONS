"""
Proofs for the #307 binary64 route map and its analyzer.

  1. The committed grid is exactly the frozen design (generate --check); every
     shape has its three binary64 seam points.
  2. On an export whose routines return the correctly rounded references, the
     analyzer forms the pre-#242 complement in binary64: a point where Q is
     small relative to P on the series side, or P on the continued-fraction
     side, may fail through the complement, and a point without that
     amplification passes.
  3. FAIL and ERROR statuses are never scored as values; a changed echo and a
     partial export are reported.
  4. An unexported grid is reported as such.

Run: python3 test_binary64_map.py   (exit 0 = pass)
"""
import csv
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import analyze_binary64_map as A  # noqa: E402
import generate_binary64_map as G  # noqa: E402
from design import SHAPES, seam  # noqa: E402


def token(x):
    hi = float(f"{x:.14E}")
    return f"{hi:.14E};{x - hi:.14E}"


def exported(mutate=None):
    rows = []
    for r in G.build():
        r = dict(r)
        r["echo_shape"], r["echo_x"] = token(float(r["shape"])), token(float(r["x"]))
        r["series_status"], r["series_p"] = "OK", token(float(r["ref_p"]))
        r["cf_status"], r["cf_q"] = "OK", token(float(r["ref_q"]))
        rows.append(r)
    if mutate:
        mutate(rows)
    return rows


def main():
    fails = []

    def check(condition, message):
        if not condition:
            fails.append(message)

    # 1. The design.
    check(G.main(["--check"]) == 0, "generate --check fails on the committed grid")
    rows = G.build()
    for a in SHAPES:
        xs = {float(r["x"]) for r in rows if float(r["shape"]) == a}
        check(set(seam(a)) <= xs, f"A={a!r} lacks its seam points")

    # 2. The complement is formed as the pre-#242 adapter did.
    result = A.score(exported())
    check(result is not None and not result["problems"], "a clean export reports problems")
    by_id = {p["point_id"]: p for p in result["points"]}
    amplified = [p for p in result["points"] if p["branch"] == "series" and p["shape"] < 1e-3
                 and p["x"] == 0.5]
    check(amplified and not amplified[0]["meets"],
          "a small-shape lower point does not lose accuracy through Q = 1 - P")
    plain = [p for p in result["points"] if p["shape"] == 2.0 and p["branch"] == "cf"
             and p["x"] > 4.0]
    check(plain and plain[0]["meets"], "an upper point without amplification fails")

    # 3. Statuses, echoes and partial exports.
    def fail_status(rs):
        rs[0]["series_status"], rs[0]["series_p"] = "FAIL", ""

    def echo(rs):
        rs[1]["echo_x"] = token(float(rs[1]["x"]) * (1 + 2 ** -52))

    def partial(rs):
        for k in ("series_status", "series_p"):
            rs[2][k] = ""

    result = A.score(exported(fail_status))
    first = result["points"][0]
    check(first["series_err_p"] is None and (first["branch"] != "series" or not first["meets"]),
          "a FAIL status is scored as a value")
    check(any("echoed" in p for p in A.score(exported(echo))["problems"]), "a changed echo is missed")
    check(any("partly" in p for p in A.score(exported(partial))["problems"]), "a partial export is missed")
    check(math.isfinite(by_id[rows[0]["point_id"]]["shape"]), "point identity lost")

    # 4. Unexported.
    with open(G.GRID, encoding="utf-8") as f:
        committed = list(csv.DictReader(f))
    if not any(r["series_status"].strip() for r in committed):
        check(A.score(committed) is None, "an unexported grid is not reported as unexported")

    if fails:
        print("FAIL: #307 binary64 route map")
        for f in fails:
            print("  - " + f)
        return 1
    print(f"PASS: #307 binary64 route map is the frozen design ({len(rows)} points); "
          "the analyzer forms the pre-#242 complement and refuses unscored statuses")
    return 0


if __name__ == "__main__":
    sys.exit(main())
