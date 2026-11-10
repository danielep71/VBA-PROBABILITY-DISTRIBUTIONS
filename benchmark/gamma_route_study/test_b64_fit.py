"""
Proofs for the #307 binary64 route fit set and its analyzer.

  1. The committed grid is exactly the frozen fit design (generate --check),
     and no holdout shape occurs in it.
  2. On an export that reproduces route_model.py, the analyzer finds no
     fidelity mismatch; a value one ulp away, a decline where the model
     answers, an answer where it declines, an ERROR and a changed echo are
     each reported.
  3. The cost fit recovers a known line, and the frozen ceiling rule stops at
     the first ladder value whose measured or predicted cost exceeds 0.1 ms.
  4. An unexported grid is reported as such.

Run: python3 test_b64_fit.py   (exit 0 = pass)
"""
import csv
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import analyze_b64_fit as A  # noqa: E402
import generate_b64_fit as G  # noqa: E402
from holdout_design import SHAPES as HOLDOUT  # noqa: E402

LINE = {"series": (4.0, 0.06), "cf": (5.0, 0.1)}


def token(x):
    hi = float(f"{x:.14E}")
    return f"{hi:.14E};{x - hi:.14E}"


def exported(committed, mutate=None, line=LINE):
    rows = []
    for r in committed:
        r = dict(r)
        a, x = float(r["shape"]), float(r["x"])
        r["echo_shape"], r["echo_x"] = token(a), token(x)
        series, cf = A.modelled(a, x)
        r["series_status"], r["series_p"] = ("OK", token(series[0])) if series else ("DECLINED", "")
        r["cf_status"], r["cf_q"] = ("OK", token(cf[0])) if cf else ("DECLINED", "")
        answered = series or cf
        if answered:
            c0, c1 = line["series" if series else "cf"]
            r["route_us"], r["route_calls"] = f"{c0 + c1 * answered[1]:.3f}", "1024"
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
    with open(G.GRID, encoding="utf-8") as f:
        committed = list(csv.DictReader(f))
    shapes = {float(r["shape"]) for r in committed} | {2 * float(r["shape"]) for r in committed}
    check(not shapes & set(HOLDOUT), "a holdout shape is in the fit set")

    # 2. Fidelity.
    clean = A.score(exported(committed))
    check(clean is not None and not clean["problems"], f"a faithful export reports {clean and clean['problems'][:3]}")
    idx = next(i for i, r in enumerate(committed) if r["point_id"].startswith("m10"))

    def ulp(rows):
        r = rows[idx]
        key = "series_p" if r["series_status"] == "OK" else "cf_q"
        v = A.double(r[key])
        r[key] = token(math.nextafter(v, math.inf))

    def declined(rows):
        r = rows[idx]
        key = "series" if r["series_status"] == "OK" else "cf"
        r[key + "_status"], r[key + ("_p" if key == "series" else "_q")] = "DECLINED", ""

    def answered(rows):
        r = rows[idx]
        key = "cf" if r["series_status"] == "OK" else "series"
        r[key + "_status"], r[key + ("_q" if key == "cf" else "_p")] = "OK", token(0.5)

    def error(rows):
        rows[idx]["series_status"] = "ERROR"

    def echo(rows):
        rows[idx]["echo_x"] = token(math.nextafter(float(rows[idx]["x"]), 0.0))

    for name, mutate, needle in (("one ulp", ulp, "differs from the model"),
                                 ("decline", declined, "declined where"),
                                 ("answer", answered, "differs from the model"),
                                 ("error", error, "status 'ERROR'"),
                                 ("echo", echo, "echoed")):
        problems = A.score(exported(committed, mutate))["problems"]
        check(any(needle in p for p in problems), f"{name}: not reported ({problems[:2]})")

    # 3. Cost fit and ceiling.
    fit = clean["cost_model"]
    for k, (c0, c1) in LINE.items():
        check(abs(fit[k][0] - c0) < 1e-2 and abs(fit[k][1] - c1) < 1e-4, f"{k}: fit {fit[k]} is not {LINE[k]}")
    expected = None
    for a2 in A.LADDER:
        measured = [p["route_us"] for p in clean["points"] if p["shape"] <= a2 and p["route_us"] is not None]
        s = A.RM.series_p(a2, math.nextafter(a2 + 1.0, 0.0))[1]
        c = A.RM.cf_q(a2, a2 + 1.0)[1]
        if max(measured) > A.BUDGET_US or max(4.0 + 0.06 * s, 5.0 + 0.1 * c) > A.BUDGET_US:
            break
        expected = a2
    check(clean["a2"] == expected and expected is not None, f"ceiling {clean['a2']} is not {expected}")
    slow = A.score(exported(committed, line={"series": (4.0, 1.0), "cf": (5.0, 1.0)}))
    check(slow["a2"] is None, "a slow route still gets a ceiling")

    # 4. Unexported.
    if not any(r["series_status"].strip() for r in committed):
        check(A.score(committed) is None, "an unexported grid is not reported as unexported")

    if fails:
        print("FAIL: #307 binary64 route fit")
        for f in fails:
            print("  - " + f)
        return 1
    print(f"PASS: #307 binary64 route fit set is the frozen design ({len(committed)} points); the analyzer "
          f"checks fidelity bit for bit and derives the shape ceiling (synthetic A2 = {clean['a2']:g})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
