"""
Fixture proof for the older-families holdout summary and its staleness check (#164).

Cases 1-8 and 10-11 use a synthetic grid and registry written to a temporary
directory, so they depend on no value in the committed grid. Case 9 uses the
record committed before #164 and asserts only that the check rejects it.

Proves:
  1. a summary just written by analyze_holdout_older.py passes the check, even
     when it records a FAIL - the check is about currency, not verdicts;
  2. the analyzer still exits non-zero while a contract fails;
  3. a failing row is listed with its arguments, contract and tracked note;
  4. a row no thresholded contract claims is listed, not dropped silently,
     with the characterization_only contract that covers it;
  5. NEGATIVE: the 43d2415 shape of drift - a contract tightened after the
     summary was written - makes the committed summary stale, and the check fails;
  6. NEGATIVE: a retired contract (the 1fac483 shape) fails the check;
  7. NEGATIVE: a hand edit that turns FAIL into PASS fails the check;
  8. NEGATIVE: a missing summary fails the check rather than passing;
  9. NEGATIVE: the summary committed in ecf3e45, read with `git show` and
     pinned by SHA-256, fails against the committed grid and registry. That is
     the drift this check exists to catch. No row text is asserted, so a
     re-export or a registry change does not break this case;
 10. NEGATIVE (#171): a current summary fails when a row is neither scored nor
     covered by a characterization_only contract - the state 1fac483 left the
     three LogGamma rows in - and the row is named;
 11. NEGATIVE (#171): a contract without a threshold whose status is not
     characterization_only does not count as a disposition.
Run: python3 test_analyze_holdout_older.py   (exit 0 = pass, nonzero = fail)
"""
import csv
import hashlib
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import analyze_holdout_older as A  # noqa: E402
import check_holdout_older_summary as C  # noqa: E402

GRID_FIELDS = ["function", "vba_kernel", "claim", "metric", "arg1", "arg2", "arg3",
               "reference", "observed_vba", "regime", "evidence_set"]
CONTRACT_FIELDS = ["contract_id", "function", "regime", "measure", "metric", "threshold",
                   "domain", "provenance", "status", "evidence", "notes"]

# Relative errors 1E-15 (Foo) and 4E-15 (Bar); Baz has no thresholded
# contract and is covered by a characterization_only one (CONTRACTS[2]).
GRID = [
    {"function": "Foo_Cumulative", "arg1": "1.5", "arg2": "2.0", "arg3": "",
     "reference": "0.5", "observed_vba": "5.00000000000000E-001;5E-016", "regime": "all"},
    {"function": "Bar_Survival", "arg1": "0.82", "arg2": "2.7", "arg3": "3.4",
     "reference": "0.25", "observed_vba": "2.50000000000001E-001;0E+000", "regime": "balanced"},
    {"function": "Baz", "arg1": "3.7", "arg2": "", "arg3": "",
     "reference": "1.0", "observed_vba": "1.0", "regime": "all"},
]
CONTRACTS = [
    {"contract_id": "Foo_Cumulative.all.output", "function": "Foo_Cumulative", "regime": "all",
     "metric": "relative", "threshold": "5E-15", "provenance": "validated and frozen"},
    {"contract_id": "Bar_Survival.balanced.output_rel", "function": "Bar_Survival",
     "regime": "balanced", "metric": "relative", "threshold": "1E-14",
     "provenance": "validated and frozen"},
    {"contract_id": "Baz.all.characterization", "function": "Baz", "regime": "all",
     "metric": "relative", "threshold": "", "provenance": "characterized",
     "status": "characterization_only"},
]

fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


def write_csv(path, fields, rows):
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in fields})


def analyze(d, contracts):
    grid, regs, out = (os.path.join(d, n) for n in ("g.csv", "c.csv", "s.md"))
    write_csv(grid, GRID_FIELDS, GRID)
    write_csv(regs, CONTRACT_FIELDS, contracts)
    proc = subprocess.run([sys.executable, os.path.join(HERE, "analyze_holdout_older.py"),
                           "--grid", grid, "--contracts", regs, "--out", out],
                          capture_output=True, text=True)
    return proc.returncode, grid, regs, out


def read(path):
    with open(path, encoding="utf-8", newline="") as f:
        return f.read()


with tempfile.TemporaryDirectory(prefix="holdout-older-") as d:
    # 1-4. A current summary that records a failure.
    tight = [dict(CONTRACTS[0]), dict(CONTRACTS[1], threshold="2E-15"), dict(CONTRACTS[2])]
    code, grid, regs, out = analyze(d, tight)
    text = read(out)
    ok, msg = C.check(out, grid, regs)
    check(ok, f"1: a freshly written summary must pass the check: {msg}")
    check("1 pass, 1 fail across 2 contract(s)" in text, "1: the FAIL is recorded in the summary\n" + text)
    check(code == 1, f"2: the analyzer exits non-zero while a contract fails (got {code})")
    check("| Bar_Survival.balanced.output_rel | Bar_Survival(0.82, 2.7, 3.4) | balanced |" in text,
          "3: the failing row is listed with its arguments\n" + text)
    check("| Foo_Cumulative(1.5, 2.0)" not in text.split("## Failing rows")[1].split("## Rows")[0],
          "3: a passing row is not listed as failing")
    A.TRACKED["Bar_Survival.balanced.output_rel"] = "tracked in #999"
    check("tracked in #999 |" in A.render(*A.load(grid, regs))[0],
          "3: a tracked contract's note is printed next to its failing row")
    del A.TRACKED["Bar_Survival.balanced.output_rel"]
    check("| Baz(3.7) | all | characterization only (Baz.all.characterization) |"
          in text.split("## Rows no contract scores")[1],
          "4: an unclaimed row is listed with its characterization contract\n" + text)

    # 5. 43d2415: the registry is tightened after the summary was written.
    code, grid, regs, out = analyze(d, CONTRACTS)
    check(code == 0 and C.check(out, grid, regs)[0], "5: baseline summary is current and passes")
    write_csv(regs, CONTRACT_FIELDS, tight)
    ok, msg = C.check(out, grid, regs)
    check(not ok and "is stale" in msg and "+| Bar_Survival.balanced.output_rel" in msg,
          "5: a contract tightened after the summary was written must fail the check\n" + msg)

    # 6. 1fac483: a contract is retired and its rows become unclaimed.
    write_csv(regs, CONTRACT_FIELDS, [CONTRACTS[0], CONTRACTS[2]])
    ok, msg = C.check(out, grid, regs)
    check(not ok and "+| Bar_Survival(0.82, 2.7, 3.4) | balanced |" in msg,
          "6: a retired contract must fail the check and list its rows as unclaimed\n" + msg)

    # 7. A hand edit that hides a failure.
    code, grid, regs, out = analyze(d, tight)
    with open(out, "w", encoding="utf-8", newline="\n") as f:
        f.write(text.replace("| FAIL |", "| PASS |"))
    ok, msg = C.check(out, grid, regs)
    check(not ok, "7: a summary edited from FAIL to PASS must fail the check")

    # 8. No summary at all.
    os.remove(out)
    ok, msg = C.check(out, grid, regs)
    check(not ok and "cannot read" in msg, "8: a missing summary must fail the check\n" + msg)

    # 10. 1fac483 after #164: the summary is current and lists the row, but
    # nothing scores it and nothing marks it characterization only.
    code, grid, regs, out = analyze(d, CONTRACTS[:2])
    ok, msg = C.check(out, grid, regs)
    check(not ok and "neither scored" in msg and "Baz(3.7) in regime all" in msg,
          "10: a current summary with an undisposed row must fail the check\n" + msg)
    check("| Baz(3.7) | all | none |" in read(out),
          "10: the undisposed row is listed with no disposition\n" + read(out))

    # 11. A thresholdless contract that is not characterization_only is no disposition.
    code, grid, regs, out = analyze(d, CONTRACTS[:2] + [dict(CONTRACTS[2], status="active")])
    ok, msg = C.check(out, grid, regs)
    check(not ok and "neither scored" in msg,
          "11: an active contract without a threshold must not count as a disposition\n" + msg)

# 9. The real record before #164, as committed in ecf3e45. Pinned by digest so
# the case cannot silently test some other file. Nothing about its rows or
# today's render is assumed beyond "they differ", so a re-export or a registry
# change leaves this case valid.
PRE_164_COMMIT = "ecf3e45"
PRE_164_PATH = "benchmark/holdout_older/holdout_older_summary.md"
PRE_164_SHA256 = "f99024b122fa7c24fc697c131fdbe9339d7955f8462db0a337eda3f81a1b265b"
proc = subprocess.run(["git", "-C", HERE, "show", f"{PRE_164_COMMIT}:{PRE_164_PATH}"],
                      capture_output=True)
check(proc.returncode == 0,
      "9: git show of the pre-#164 summary failed (full history is required): "
      + proc.stderr.decode(errors="replace"))
if proc.returncode == 0:
    check(hashlib.sha256(proc.stdout).hexdigest() == PRE_164_SHA256,
          "9: the pre-#164 summary does not have its pinned SHA-256")
    with tempfile.TemporaryDirectory(prefix="holdout-older-real-") as d:
        stale = os.path.join(d, "holdout_older_summary.md")
        with open(stale, "wb") as f:
            f.write(proc.stdout)
        ok, msg = C.check(stale, A.GRID, A.CONTRACTS)
        check(not ok and "is stale" in msg,
              "9: the pre-#164 summary must fail against the committed grid and registry\n" + msg)

if fails:
    print(f"FAIL: {len(fails)} holdout_older summary check(s)")
    for m in fails:
        print("  - " + m)
    raise SystemExit(1)
print("PASS: a current holdout_older summary passes whatever its verdicts; a tightened or "
      "retired contract, a hand-edited verdict, a missing file, the pre-#164 record and a row "
      "neither scored nor marked characterization only fail")
