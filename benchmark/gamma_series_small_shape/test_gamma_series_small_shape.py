"""
Proofs for the #23 study grid and its analyzer.

  1. The committed grid is exactly the frozen design: generate --check passes,
     the holdout is 11 shapes x 2 arguments x 2 routes, and each Chi-square row
     is the exact doubling of its Gamma row, with the same reference.
  2. The analyzer passes an export whose every result is the correctly rounded
     reference, with echoes equal to the inputs, and both routes bit-identical.
  3. It fails a result 1E-14 off, an ERROR result and an echo that differs from
     the reference input, and counts the routes that disagree. The ladder rules
     (B1, B3, B4X) refuse a result lost to ERROR instead of scoring it as zero.
  4. It reports "no observations" for the committed, unexported grid.

Run: python3 test_gamma_series_small_shape.py   (exit 0 = pass)
"""
import csv
import io
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import analyze_gamma_series_small_shape as A  # noqa: E402
import generate_gamma_series_small_shape as G  # noqa: E402


def token(x):
    """The exporter's hi;lo encoding of a Double."""
    hi = float(f"{x:.14E}")
    return f"{hi:.14E};{x - hi:.14E}"


def write_grid(rows):
    fd, path = tempfile.mkstemp(suffix=".csv")
    with os.fdopen(fd, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=G.FIELDS, lineterminator="\n")
        w.writeheader()
        w.writerows(rows)
    return path


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


def main():
    fails = []

    def check(condition, message):
        if not condition:
            fails.append(message)

    # 1. The committed grid is the design.
    check(G.main(["--check"]) == 0, "generate --check fails on the committed grid")
    rows = G.build()
    holdout = [r for r in rows if r["regime"] == "holdout"]
    check(len(holdout) == 44, f"holdout has {len(holdout)} rows, not 44")
    check({float(r["arg2"]) for r in holdout if r["function"] == "Gamma_Cumulative"}
          == set(G.HOLDOUT_SHAPES), "holdout shapes differ from HOLDOUT_DESIGN.md")
    for g, c in zip(rows[0::2], rows[1::2]):
        if g["regime"] == "scope_probe":
            continue
        check(c["function"] == "ChiSquare_Cumulative" and g["function"] == "Gamma_Cumulative",
              "routes are not paired Gamma, Chi-square")
        check(float(c["arg1"]) == 2 * float(g["arg1"]) and float(c["arg2"]) == 2 * float(g["arg2"]),
              f"{c['arg1']}, {c['arg2']} is not the exact doubling of {g['arg1']}, {g['arg2']}")
        check(c["reference"] == g["reference"], "paired routes have different references")
    check(sorted((float(r["arg1"]), float(r["arg2"])) for r in rows if r["regime"] == "scope_probe")
          == sorted(G.SCOPE_PROBES), "scope probes differ from HOLDOUT_DESIGN.md")

    # 2. A correctly rounded export passes.
    path = write_grid(exported())
    try:
        result = A.part_a(path)
    finally:
        os.remove(path)
    check(result is not None and result["pass"], f"a correct export fails: {result and result['problems']}")
    check(result and result["route_pairs"] == 23 and result["route_pairs_bit_identical"] == 23,
          "route pairs not all identical on a correct export")

    # 3. Failures are caught.
    def off(rows):
        rows[0]["observed_vba"] = token(float(rows[0]["reference"]) * (1 - 1e-14))

    def error(rows):
        rows[2]["observed_vba"] = "ERROR"

    def echo(rows):
        rows[4]["echo_arg2"] = token(float(rows[4]["arg2"]) * (1 + 2 ** -52))

    for name, mutate, needle in (("1E-14 off", off, "exceeds"), ("ERROR", error, "no numeric"),
                                 ("echo", echo, "echoed inputs")):
        path = write_grid(exported(mutate))
        try:
            result = A.part_a(path)
        finally:
            os.remove(path)
        check(not result["pass"] and any(needle in p for p in result["problems"]),
              f"{name}: not caught ({result['problems']})")
        if name == "1E-14 off":
            check(result["route_pairs_bit_identical"] == 22, "a route disagreement is not counted")

    # 3b. The ladder rules refuse a missing result instead of scoring it as zero.
    after_text = open(os.path.join(A.ROOT, A.LADDER_PATH), encoding="utf-8").read()
    before_text = A.git_show(A.BASELINE, A.LADDER_PATH)

    def lose(text, pid, surface):
        lines = text.splitlines(keepends=True)
        header = lines[0].rstrip("\n").split(",")
        i = header.index(surface)
        for k, line in enumerate(lines):
            cols = line.rstrip("\n").split(",")
            if cols[0] == pid:
                cols[i - 1], cols[i] = "ERROR", ""
                lines[k] = ",".join(cols) + "\n"
        return "".join(lines)

    for rule, pid, surface in (("B3", "xsub_s0", "cdf"), ("B4X", "x2_s3", "survival"),
                               ("B1", "x05_s2", "cdf"), ("B1", "x05_s4", "series_p")):
        result = A.part_b(lose(after_text, pid, surface), before_text)
        check(not result["pass"] and any(p.startswith(f"{rule} {pid}: {surface} after")
                                         for p in result["problems"]),
              f"{rule}: a lost {surface} at {pid} is not refused ({result['problems'][:3]})")

    # 3c. B1 scores the raw series P against 3E-15, not only its movement.
    def skew(text, pid):
        lines = text.splitlines(keepends=True)
        header = lines[0].rstrip("\n").split(",")
        i = header.index("series_p")
        for k, line in enumerate(lines):
            cols = line.rstrip("\n").split(",")
            if cols[0] == pid:
                hi, lo = cols[i].split(";")
                cols[i] = token((float(hi) + float(lo)) * (1 - 1e-14))
                lines[k] = ",".join(cols) + "\n"
        return "".join(lines)

    result = A.part_b(skew(after_text, "x05_s5"), before_text)
    check(any(p.startswith("B1 x05_s5: series_p error") for p in result["problems"]),
          f"B1: a series P 1E-14 off is not refused ({result['problems'][:3]})")

    # 4. The committed grid has no observations yet, or a full set.
    with open(G.GRID, encoding="utf-8") as f:
        committed = list(csv.DictReader(io.StringIO(f.read())))
    observed = [r for r in committed if r["observed_vba"].strip()]
    check(len(observed) in (0, len(committed)), "the committed grid is partly observed")
    if not observed:
        check(A.part_a() is None, "an unexported grid is not reported as unobserved")

    if fails:
        print("FAIL: #23 study grid and analyzer")
        for f in fails:
            print("  - " + f)
        return 1
    print("PASS: #23 study grid is the frozen design (44 holdout, 3 scope, 2 counterexample rows); "
          "the analyzer passes a correct export and catches a 1E-14 error, ERROR and a changed echo")
    return 0


if __name__ == "__main__":
    sys.exit(main())
