"""
#359: verdict of the cheaper beta inverse against PREREGISTRATION.md.

Reads four files. beta_inverse_cost_baseline.csv and cost_baseline.txt come
from one Excel run on the source of v1.0.0; beta_inverse_cost_grid.csv and
cost_candidate.txt from one Excel run on the candidate. The two grid files
hold the same rows. The two text files hold the COST lines of the regression
log for the inverses that share the beta inverse. Every bound is fixed in
the preregistration; nothing is fitted here.

Accuracy, on the grid:

  R0  numbers     every cell of both files is a Double or a worksheet error:
                  none is blank and none is a runtime fault.
  R1  no regress  where both files hold a number, the candidate's error is at
                  most the baseline's plus FLOOR.
  R2  guarantee   where the baseline is within GUARANTEE of the reference, so
                  is the candidate.
  R3  domain      a row the baseline answers with a number, the candidate
                  answers with a number.
  R4  worst       where the baseline is within WORST of the reference, so is
                  the candidate.

Cost, on the benchmark lines:

  C0  lines       both text files hold every benchmark line, each with ok=1.
  C1  direct      on the lines of DIRECT the candidate takes at most
                  FACTOR_DIRECT times what the routine before #35 took, that
                  time scaled to the baseline run: before * baseline / after,
                  with before and after from ../production/PRODUCTION_RESULT.md.
  C2  pair        on the lines of PAIR the candidate takes at most
                  FACTOR_PAIR times what the baseline run took.

The error is relative to the reference. Reported, not gating: the worst
error of each file by kind, how many rows changed and in which direction,
the Student-t cost lines, and every cost ratio.

Usage: python3 analyze.py [--baseline-only]
--baseline-only checks R0 and C0 on the baseline files alone, for the run
that fills them before the candidate exists.
Exit 0 only when every rule holds.
"""
import argparse
import csv
import os
import re
import struct
import sys
from collections import defaultdict

from mpmath import mp, mpf

HERE = os.path.dirname(os.path.abspath(__file__))
BASELINE = os.path.join(HERE, "beta_inverse_cost_baseline.csv")
CANDIDATE = os.path.join(HERE, "beta_inverse_cost_grid.csv")
COST_BASELINE = os.path.join(HERE, "cost_baseline.txt")
COST_CANDIDATE = os.path.join(HERE, "cost_candidate.txt")
GUARANTEE = mpf("1e-12")      # inverse guarantee of #327
FLOOR = mpf("3e-14")          # the binary64 noise on the residual
WORST = mpf("1e-13")
FACTOR_DIRECT = 5.0
FACTOR_PAIR = 0.5
KEY_COLUMNS = 7

# (route, A, B, X): milliseconds per call before and after #35, from the cost
# table of ../production/PRODUCTION_RESULT.md.
RECORDED = {
    ("beta_inv", "0.5", "0.5", "0.1"): (0.027, 1.686),
    ("beta_inv", "0.5", "0.5", "0.9"): (0.025, 1.503),
    ("beta_inv", "0.7", "25", "0.1"): (0.021, 1.396),
    ("beta_inv", "0.7", "25", "0.9"): (0.049, 10.559),
    ("beta_inv", "2", "5", "0.1"): (0.121, 0.751),
    ("beta_inv", "2", "5", "0.9"): (0.024, 0.591),
    ("beta_inv", "50", "80", "0.1"): (0.049, 28.320),
    ("beta_inv", "50", "80", "0.9"): (0.051, 6.744),
    ("f_inv", "5", "10", "0.1"): (0.055, 2.365),
    ("f_inv", "5", "10", "0.9"): (0.046, 2.991),
    ("f_inv", "50", "100000000", "0.1"): (0.062, 4.944),
    ("f_inv", "50", "100000000", "0.9"): (0.246, 5.463),
    ("f_inv", "100", "100", "0.1"): (0.061, 24.292),
    ("f_inv", "100", "100", "0.9"): (0.062, 7.446),
    ("f_inv", "100000", "2", "0.1"): (0.051, 0.793),
    ("f_inv", "100000", "2", "0.9"): (0.043, 0.759),
}
# The lines on which the Python model of the candidate makes a pair
# evaluation; on the others it makes none.
PAIR = frozenset({("beta_inv", "0.7", "25", "0.9"), ("beta_inv", "50", "80", "0.9"),
                  ("f_inv", "50", "100000000", "0.9"), ("f_inv", "100000", "2", "0.1")})
DIRECT = frozenset(RECORDED) - PAIR
REPORTED = (("t_inv", "3", "", "0.9"), ("t_inv", "10", "", "0.9"), ("t_inv", "30", "", "0.9"))
LINE = re.compile(r"(?:COST )?(?P<route>beta_inv|f_inv|t_inv) A=(?P<a>\S+)(?: B=(?P<b>\S+))? "
                  r"X=(?P<x>\S+) ok=(?P<ok>[01]) calls=\d+ us_per_call=(?P<us>[0-9.]+)\s*$")


def from_bits(text):
    return struct.unpack("<d", struct.pack("<Q", int(text, 16)))[0]


def is_bits(text):
    return len(text) == 16 and all(c in "0123456789ABCDEFabcdef" for c in text)


def error(text, ref):
    value, reference = mpf(from_bits(text)), mpf(ref)
    return abs(value - reference) if reference == 0 else abs((value - reference) / reference)


def read(path):
    with open(path, encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def read_cost(path):
    """{(route, A, B, X): (microseconds per call, ok)} from COST lines."""
    out = {}
    with open(path, encoding="utf-8") as handle:
        for text in handle:
            found = LINE.search(text.strip())
            if found:
                key = (found["route"], found["a"], found["b"] or "", found["x"])
                out[key] = (float(found["us"]), found["ok"] == "1")
    return out


def cell_problem(text):
    if not text.strip():
        return "the cell is blank: the file was not exported"
    if text == "ERROR":
        return "a runtime fault"
    if text != "CVERR" and not is_bits(text):
        return f"not a Double: {text!r}"
    return None


def analyze(base, cand=None):
    """(failures, report) for the grid. cand None checks the baseline alone."""
    mp.dps = 60
    failures = []
    worst = {"baseline": defaultdict(lambda: mpf(0)), "candidate": defaultdict(lambda: mpf(0))}
    counts = defaultdict(int)

    def fail(rule, row, text):
        failures.append(f"{rule} {row['point_id']} ({row['kind']}): {text}")

    if cand is not None and [list(r.values())[:KEY_COLUMNS] for r in base] \
            != [list(r.values())[:KEY_COLUMNS] for r in cand]:
        return ["R0: the two files do not hold the same rows"], {}

    for index, b in enumerate(base):
        counts["rows"] += 1
        problem = cell_problem(b["value"])
        if problem:
            fail("R0", b, f"baseline: {problem}")
            continue
        b_number = is_bits(b["value"])
        if b_number:
            e_base = error(b["value"], b["ref"])
            worst["baseline"][b["kind"]] = max(worst["baseline"][b["kind"]], e_base)
        else:
            counts["baseline_errors"] += 1
        if cand is None:
            continue

        c = cand[index]
        problem = cell_problem(c["value"])
        if problem:
            fail("R0", c, f"candidate: {problem}")
            continue
        if not is_bits(c["value"]):
            counts["candidate_errors"] += 1
            if b_number:
                fail("R3", c, "the baseline returns a number and the candidate an error")
            continue
        e_new = error(c["value"], c["ref"])
        worst["candidate"][c["kind"]] = max(worst["candidate"][c["kind"]], e_new)
        if not b_number:
            continue
        if c["value"].upper() != b["value"].upper():
            counts["changed"] += 1
        if e_new < e_base:
            counts["closer"] += 1
        elif e_new > e_base:
            counts["farther"] += 1
        if e_new > e_base + FLOOR:
            fail("R1", c, f"error {mp.nstr(e_new, 3)}, baseline {mp.nstr(e_base, 3)}")
        if e_base <= GUARANTEE < e_new:
            fail("R2", c, f"error {mp.nstr(e_new, 3)} above {mp.nstr(GUARANTEE, 3)}")
        if e_base <= WORST < e_new:
            fail("R4", c, f"error {mp.nstr(e_new, 3)} above {mp.nstr(WORST, 3)}")

    return failures, {"counts": dict(counts), "worst": worst}


def analyze_cost(base, cand=None):
    """(failures, lines) for the benchmark. cand None checks the baseline alone."""
    failures, lines = [], []

    def name(key):
        return f"{key[0]} ({key[1]}{', ' + key[2] if key[2] else ''}) p={key[3]}"

    for label, costs in (("baseline", base), ("candidate", cand)):
        if costs is None:
            continue
        for key in sorted(RECORDED) + list(REPORTED):
            if key not in costs:
                failures.append(f"C0 {name(key)}: no {label} line")
            elif not costs[key][1]:
                failures.append(f"C0 {name(key)}: the {label} call returned an error")
    if failures or cand is None:
        return failures, lines

    for key in sorted(RECORDED):
        before, after = RECORDED[key]
        b_us, c_us = base[key][0], cand[key][0]
        scaled = before / after * b_us          # the routine before #35, on this run
        if key in PAIR:
            rule, limit = "C2", FACTOR_PAIR * b_us
        else:
            rule, limit = "C1", FACTOR_DIRECT * scaled
        lines.append(f"  {rule} {name(key):34s} baseline {b_us:10.1f} us  candidate {c_us:9.1f} us  "
                     f"limit {limit:9.1f} us  {c_us / scaled:6.1f}x before #35  "
                     f"{b_us / c_us:6.1f}x faster")
        if c_us > limit:
            failures.append(f"{rule} {name(key)}: {c_us:.1f} us above the limit {limit:.1f} us")
    for key in REPORTED:
        lines.append(f"  -- {name(key):34s} baseline {base[key][0]:10.1f} us  "
                     f"candidate {cand[key][0]:9.1f} us")
    return failures, lines


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--baseline-only", action="store_true")
    a = ap.parse_args(argv)
    only = a.baseline_only
    failures, report = analyze(read(BASELINE), None if only else read(CANDIDATE))
    for name in ("baseline",) if only else ("baseline", "candidate"):
        for kind, value in sorted(report.get("worst", {}).get(name, {}).items()):
            print(f"  worst {name:9s} {kind:12s} {mp.nstr(value, 3)}")
    for key, value in sorted(report.get("counts", {}).items()):
        print(f"  {key}: {value}")
    cost_failures, lines = analyze_cost(read_cost(COST_BASELINE),
                                        None if only else read_cost(COST_CANDIDATE))
    for text in lines:
        print(text)
    failures += cost_failures
    for text in failures:
        print(f"FAIL: {text}")
    if failures:
        print(f"REJECTED: {len(failures)} finding(s)")
        return 1
    print("PASS: the baseline files are complete" if only
          else "ACCEPTED: R0 to R4 and C0 to C2 hold")
    return 0


if __name__ == "__main__":
    sys.exit(main())
