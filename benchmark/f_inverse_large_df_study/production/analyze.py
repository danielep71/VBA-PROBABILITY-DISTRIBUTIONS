"""
#35: verdict of the production holdout against PRODUCTION_PREREGISTRATION.md.

Reads f_inverse_pair_baseline.csv (Excel, before the change) and
f_inverse_pair_grid.csv (Excel, the candidate). The two files hold the same
rows. Every bound is fixed in the preregistration; nothing is fitted here.

  R0  numbers      every cell of both files is a Double or a worksheet
                   error: none is blank and none is a runtime fault.
  R1  no regress   where both files hold a number, the candidate's error is
                   at most max(the baseline's, FLOOR).
  R2  guarantee    every number the candidate returns for an F or Beta
                   row is within GUARANTEE of the reference: the inverse
                   guarantee of #327. The Student-t rows are held by R1 and
                   R4 alone (Addendum 1).
  R3  domain       a row the baseline answers with a number, the candidate
                   answers with a number.
  R4  shared path  on the Student-t rows the candidate is the baseline's
                   Double, one of its two neighbours, or closer to the
                   reference than the baseline.

The error is relative to the reference, and absolute where the reference is
zero. Reported, not gating: the worst error of each file by kind, how many
rows changed, the distance of the candidate from 2^-48, and whether the F
and Beta rows are the Doubles of the Python model of the rule.

Usage: python3 analyze.py [--baseline-only]
--baseline-only checks R0 on the baseline file alone, for the wave that
fills it before the candidate exists.
Exit 0 only when every rule holds on every row.
"""
import argparse
import csv
import os
import struct
import sys
from collections import defaultdict

from mpmath import mp, mpf

HERE = os.path.dirname(os.path.abspath(__file__))
BASELINE = os.path.join(HERE, "f_inverse_pair_baseline.csv")
CANDIDATE = os.path.join(HERE, "f_inverse_pair_grid.csv")
GUARANTEE = mpf("1e-12")      # inverse guarantee of #327
FLOOR = mpf("1e-14")          # a change below this is not a regression
BINARY64 = mpf(2) ** -48      # reported only
KEY_COLUMNS = 7


def from_bits(text):
    return struct.unpack("<d", struct.pack("<Q", int(text, 16)))[0]


def is_bits(text):
    return len(text) == 16 and all(c in "0123456789ABCDEFabcdef" for c in text)


def error(text, ref):
    """Relative error of the Double in a cell; absolute where ref is zero."""
    value, reference = mpf(from_bits(text)), mpf(ref)
    return abs(value - reference) if reference == 0 else abs((value - reference) / reference)


def ordinal(text):
    """Doubles in order as integers, so that neighbours differ by one."""
    word = int(text, 16)
    return -(word & (2 ** 63 - 1)) if word >> 63 else word


def read(path):
    with open(path, encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def cell_problem(text):
    if not text.strip():
        return "the cell is blank: the file was not exported"
    if text == "ERROR":
        return "a runtime fault"
    if text != "CVERR" and not is_bits(text):
        return f"not a Double: {text!r}"
    return None


def analyze(base, cand=None, model=None):
    """(failures, report). cand None checks the baseline file alone."""
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
        if c["value"].upper() != b["value"].upper():
            counts["changed"] += 1
        if e_new > BINARY64:
            counts["above_2^-48"] += 1
        if e_new > GUARANTEE and c["kind"] != "t_invariance":
            fail("R2", c, f"error {mp.nstr(e_new, 3)} above {mp.nstr(GUARANTEE, 3)}")
        if b_number:
            if e_new > max(e_base, FLOOR):
                fail("R1", c, f"error {mp.nstr(e_new, 3)}, baseline {mp.nstr(e_base, 3)}")
            if e_new < e_base:
                counts["closer"] += 1
            elif e_new > e_base:
                counts["farther"] += 1
            if c["kind"] == "t_invariance" and e_new > e_base \
                    and abs(ordinal(c["value"]) - ordinal(b["value"])) > 1:
                fail("R4", c, "more than one unit from the baseline and farther from the reference")
        if model is not None and c["func"] in ("F", "Beta"):
            counts["model_rows"] += 1
            if model(c) == c["value"].upper():
                counts["model_equal"] += 1

    return failures, {"counts": dict(counts), "worst": worst}


def model_bits(row):
    """The Double the Python model of the rule returns for an F or Beta row."""
    sys.path.insert(0, os.path.join(os.path.dirname(HERE), "fix_comparison"))
    import tail_aware_model as p4
    p, a, b = (from_bits(row[k]) for k in ("p_bits", "a_bits", "b_bits"))
    try:
        if row["func"] == "F":
            value = p4.evaluate(p, a, b)[0]
        else:
            value = p4.inverse_pair(p, 1.0 - p, a, b)[0]
    except Exception:  # the model declines; reported as a difference
        return None
    return f"{struct.unpack('<Q', struct.pack('<d', value))[0]:016X}"


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--baseline-only", action="store_true")
    a = ap.parse_args(argv)
    base = read(BASELINE)
    cand = None if a.baseline_only else read(CANDIDATE)
    failures, report = analyze(base, cand, None if a.baseline_only else model_bits)
    for name in ("baseline",) if cand is None else ("baseline", "candidate"):
        for kind, value in sorted(report.get("worst", {}).get(name, {}).items()):
            print(f"  worst {name:9s} {kind:13s} {mp.nstr(value, 3)}")
    for key, value in sorted(report.get("counts", {}).items()):
        print(f"  {key}: {value}")
    for text in failures:
        print(f"FAIL: {text}")
    if failures:
        print(f"REJECTED: {len(failures)} finding(s)")
        return 1
    print("PASS: the baseline file is complete" if cand is None
          else "ACCEPTED: R0 to R4 hold on every row")
    return 0


if __name__ == "__main__":
    sys.exit(main())
