"""
Negative and positive controls for the two-tier accuracy policy (#327).

  1. derive() applies the rule: forward relative errors get max(target, 1E-13),
     inverses max(target, 1E-12), balanced Beta CDF and survival
     max(target, 2E-13); absolute, scaled, log-absolute and tail-residual
     contracts and contracts without a threshold keep what they have.
  2. The committed accuracy_guarantees.csv is the generator's output; a hand
     edit, a guarantee stricter than its target, a stale file after a registry
     change, and an absent file are refused, and the gate refuses to run on them.
  3. compute_errors.py: a main-grid row above its measured target but inside
     its guarantee PASSES and is reported "target missed"; a row above its
     guarantee FAILS the gate.
  4. holdout/analyze_holdout.py: the same two cases on the holdout grid.
  5. _invariants: a silent wrong value, an exact 1 against a reference below
     one, a reversal beyond the allowance and a P + Q mismatch each fail; a
     reversal inside the allowance is counted but does not fail; a row under a
     documented threshold looser than 1E-6 is not called silent.

Run: python3 test_accuracy_tiers.py   (exit 0 = pass)
"""
import csv
import importlib.util
import io
import os
import subprocess
import sys
import tempfile
from decimal import Decimal

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import accuracy_tiers as T  # noqa: E402
import _invariants as I  # noqa: E402
from test_gate_degradation import clean_manifest_seam  # noqa: E402

fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


def contract(cid, function, measure="output_error", metric="relative", threshold="5E-15"):
    return {"contract_id": cid, "function": function, "regime": "all", "measure": measure,
            "metric": metric, "threshold": threshold, "status": "active"}


# 1. The rule.
cases = [
    (contract("A.all.output", "Normal_Cumulative"), "1E-13"),
    (contract("A.all.loose", "Normal_Cumulative", threshold="2E-11"), "2E-11"),
    (contract("A.all.edge", "Normal_Cumulative", threshold="1E-13"), "1E-13"),
    (contract("B.all.q", "Beta_InverseCumulative", measure="quantile_error"), "1E-12"),
    (contract("B.all.inv", "Normal_InverseCumulative"), "1E-12"),
    (contract("B.all.invloose", "F_InverseCumulative", threshold="2E-10"), "2E-10"),
    (contract("Beta_Survival.balanced.output_rel", "Beta_Survival"), "2E-13"),
    (contract("Beta_Cumulative.balanced.output_rel", "Beta_Cumulative", threshold="2E-14"),
     "2E-13"),
    (contract("Beta_Survival.unbalanced.output_rel", "Beta_Survival", threshold="2E-10"),
     "2E-10"),
    (contract("C.all.abs", "LogGamma", metric="absolute", threshold="5E-15"), "5E-15"),
    (contract("C.all.tail", "Beta_InverseCumulative", measure="tail_probability_residual",
              threshold="1E-14"), "1E-14"),
    (contract("C.all.scaled", "LogGamma1p", measure="scaled_output_error",
              metric="absolute", threshold="5E-16"), "5E-16"),
    (contract("C.all.logabs", "Binomial_LogPMF", measure="log_absolute_error",
              metric="absolute", threshold="1E-13"), "1E-13"),
    (contract("C.all.none", "Gamma_Survival", threshold=""), ""),
]
for c, want in cases:
    got = T.derive(c)[0]
    check(got == want, f"1: {c['contract_id']} guarantee {got!r}, expected {want!r}")
    if want:
        check(Decimal(got) >= Decimal(c["threshold"]),
              f"1: {c['contract_id']} guarantee is stricter than its target")

# 2. The committed file and its refusals.
registry = T.load_registry()
with open(T.GUARANTEES, encoding="utf-8", newline="") as f:
    committed = f.read()
check(T.problems(registry, committed) == [], "2: committed accuracy_guarantees.csv is current")
check(len(T.load_guarantees(contracts=registry)) == len(registry),
      "2: every registry contract has a guarantee row")
rows = list(csv.DictReader(io.StringIO(committed)))
victim = next(r for r in rows if r["guarantee"] == "1E-13" and r["measured_target"] == "5E-15")
edited = committed.replace(f"{victim['contract_id']},5E-15,1E-13,",
                           f"{victim['contract_id']},5E-15,1E-11,", 1)
check(edited != committed and T.problems(registry, edited),
      "2: a hand-loosened guarantee is refused")
stricter = committed.replace(f"{victim['contract_id']},5E-15,1E-13,",
                             f"{victim['contract_id']},5E-15,1E-15,", 1)
check(any("stricter than its measured target" in p for p in T.problems(registry, stricter)),
      "2: a guarantee stricter than its target is refused by name")
retargeted = [dict(c) for c in registry]
next(c for c in retargeted if c["contract_id"] == victim["contract_id"])["threshold"] = "2E-12"
check(T.problems(retargeted, committed), "2: a registry change without regeneration is refused")
with tempfile.TemporaryDirectory(prefix="tiers-") as d:
    bad = os.path.join(d, "g.csv")
    with open(bad, "w", encoding="utf-8", newline="") as f:
        f.write(edited)
    try:
        T.load_guarantees(path=bad, contracts=registry)
        check(False, "2: load_guarantees accepted a hand-edited file")
    except ValueError:
        pass
    try:
        T.load_guarantees(path=os.path.join(d, "absent.csv"), contracts=registry)
        check(False, "2: load_guarantees accepted an absent file")
    except OSError:
        pass


# 3. compute_errors.py end to end, with the synthetic provenance seam of
# test_gate_degradation.py so the live binding state cannot mask the verdicts.
def load_ce():
    spec = importlib.util.spec_from_file_location("ce_tiers", os.path.join(HERE, "compute_errors.py"))
    ce = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(ce)
    return ce


def run_gate(grid_path, guarantees_path=None):
    ce = load_ce()
    out = tempfile.NamedTemporaryFile("w", suffix=".md", delete=False).name
    saved = (sys.argv, sys.modules.get("_manifest"), T.GUARANTEES)
    sys.argv = ["compute_errors.py", "--grid", grid_path, "--out", out]
    sys.modules["_manifest"] = clean_manifest_seam()
    if guarantees_path:
        T.GUARANTEES = guarantees_path
        T.load_guarantees.__defaults__ = (guarantees_path, None)
    code = None
    stdout = io.StringIO()
    old_stdout = sys.stdout
    sys.stdout = stdout
    try:
        ce.main()
    except SystemExit as exc:
        code = exc.code
    finally:
        sys.stdout = old_stdout
        sys.argv = saved[0]
        if saved[1] is None:
            sys.modules.pop("_manifest", None)
        else:
            sys.modules["_manifest"] = saved[1]
        T.GUARANTEES = saved[2]
        T.load_guarantees.__defaults__ = (saved[2], None)
    with open(out, encoding="utf-8") as f:
        summary = f.read()
    os.unlink(out)
    return code, summary, stdout.getvalue()


def scaled(reference, factor):
    """An observation `factor` above the reference, as one exported part."""
    return format(Decimal(reference) * (1 + Decimal(factor)), "E")


def mutated_grid(src, dst, function, factor):
    with open(src, encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        fields, grid = reader.fieldnames, list(reader)
    row = next(r for r in grid if r["function"] == function and r["observed_vba"])
    row["observed_vba"] = scaled(row["reference"], factor)
    with open(dst, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, lineterminator="\n")
        w.writeheader()
        w.writerows(grid)


MAIN = os.path.join(HERE, "probability_accuracy_grid.csv")
CID = "Exponential_Cumulative.all.output"      # target 5E-15, guarantee 1E-13
with tempfile.TemporaryDirectory(prefix="tiers-gate-") as d:
    code, summary, _ = run_gate(MAIN)
    check(code == 0, f"3: the committed main grid passes the gate (exit {code})")
    miss = os.path.join(d, "miss.csv")
    mutated_grid(MAIN, miss, "Exponential_Cumulative", "5E-14")
    code, summary, out = run_gate(miss)
    line = next((ln for ln in summary.splitlines() if ln.startswith(f"| {CID} |")), "")
    check(code == 0, f"3: a target-only miss must not fail the gate (exit {code})\n{out}")
    check("| 1E-13 | 5E-15 |" in line and "PASS · target missed" in line,
          f"3: the contract reports guarantee, target and the miss: {line!r}")
    check("> **Measured target** — 1 passing contract(s) exceed" in summary,
          "3: the measured-target tally counts the miss")
    breach = os.path.join(d, "breach.csv")
    mutated_grid(MAIN, breach, "Exponential_Cumulative", "5E-13")
    code, summary, out = run_gate(breach)
    line = next((ln for ln in summary.splitlines() if ln.startswith(f"| {CID} |")), "")
    check(code == 1 and "FAIL" in line, f"3: a guarantee breach must fail the gate "
                                         f"(exit {code}): {line!r}")
    # The gate refuses a hand-edited guarantees file instead of using it.
    edited_path = os.path.join(d, "edited.csv")
    with open(edited_path, "w", encoding="utf-8", newline="") as f:
        f.write(edited)
    code, _summary, out = run_gate(MAIN, guarantees_path=edited_path)
    check(code == 1 and "GUARANTEES" in out,
          f"2: the gate refuses a hand-edited guarantees file (exit {code})\n{out}")

# 4. The independent holdout analyzer.
HOLDOUT = os.path.join(HERE, "holdout", "holdout_grid.csv")
with open(HOLDOUT, encoding="utf-8", newline="") as f:
    holdout_rows = list(csv.DictReader(f))
by_key = {(c["function"], c["regime"]): c for c in registry
          if c["measure"] == "output_error" and c["metric"] == "relative"}
pick = next(r for r in holdout_rows
            if (r["function"], r["regime"]) in by_key
            and T.derive(by_key[(r["function"], r["regime"])])[0] == "1E-13"
            and Decimal(by_key[(r["function"], r["regime"])]["threshold"]) < Decimal("1E-14")
            and r["observed_vba"])
hcid = by_key[(pick["function"], pick["regime"])]["contract_id"]


def run_holdout(grid_path):
    out = tempfile.NamedTemporaryFile("w", suffix=".md", delete=False).name
    proc = subprocess.run([sys.executable, os.path.join(HERE, "holdout", "analyze_holdout.py"),
                           "--grid", grid_path, "--out", out], capture_output=True, text=True)
    with open(out, encoding="utf-8") as f:
        text = f.read()
    os.unlink(out)
    line = next((ln for ln in text.splitlines() if ln.startswith(f"| {hcid} |")), "")
    return proc.returncode, line, text


with tempfile.TemporaryDirectory(prefix="tiers-holdout-") as d:
    for factor, want_code, want_target, want_verdict in (("5E-14", 0, "missed", "PASS"),
                                                          ("5E-13", 1, "missed", "FAIL")):
        g = os.path.join(d, f"h{factor}.csv")
        with open(HOLDOUT, encoding="utf-8", newline="") as f:
            reader = csv.DictReader(f)
            fields, grid = reader.fieldnames, list(reader)
        row = next(r for r in grid if r["function"] == pick["function"]
                   and r["regime"] == pick["regime"] and r["observed_vba"])
        row["observed_vba"] = scaled(row["reference"], factor)
        with open(g, "w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=fields, lineterminator="\n")
            w.writeheader()
            w.writerows(grid)
        code, line, text = run_holdout(g)
        cells = [c.strip() for c in line.strip("|").split("|")]
        check(code == want_code and len(cells) == 10 and cells[2] == "1E-13"
              and cells[7] == want_target and cells[9] == want_verdict,
              f"4: holdout {hcid} at +{factor}: exit {code}, row {line!r}")

# 5. The always-gating invariants, on synthetic rows.
GOV = {("P_Cumulative", "all"): (Decimal("1E-13"), Decimal("5E-15"), "relative"),
       ("P_Survival", "all"): (Decimal("1E-13"), Decimal("5E-15"), "relative"),
       ("Loose_Density", "all"): (Decimal("5E-5"), Decimal("5E-5"), "relative")}


def row(function, x, ref, obs, arg2="1"):
    return {"function": function, "regime": "all", "metric": "rel", "arg1": x, "arg2": arg2,
            "arg3": "", "arg4": "", "reference": ref, "observed_vba": obs, "expected_error": ""}


clean = [row("P_Cumulative", "1", "0.25", "0.25"), row("P_Cumulative", "2", "0.5", "0.5"),
         row("P_Survival", "1", "0.75", "0.75"), row("P_Survival", "2", "0.5", "0.5"),
         row("Loose_Density", "1", "2.0", "2.00002")]
for name, fn in I.CHECKS:
    res = fn(clean, GOV)
    check(not res.violations, f"5: clean rows violate {name}: {res.violations}")
check(I.silent_wrong(clean, GOV).checked == 5,
      "5: a row under a documented 5E-5 threshold is scored, not called silent")
res = I.silent_wrong(clean + [row("Unc_Density", "1", "1.0", "1.00001")], GOV)
check(len(res.violations) == 1, "5: a 1E-5 error on an uncontracted row is silent wrong")
res = I.silent_wrong([row("P_Cumulative", "9", "0.99999999999999988898", "1")], GOV)
check(len(res.violations) == 1, "5: an exact 1 against a reference below one is silent wrong")
back = [row("P_Cumulative", "1", "0.25", "0.25"), row("P_Cumulative", "2", "0.5", "0.24")]
res = I.monotonicity(back, GOV)
check(len(res.violations) == 1, "5: a CDF running backwards beyond its allowance fails")
near = [row("P_Cumulative", "1", "0.5", "0.5000000000000001"),
        row("P_Cumulative", "2", "0.50000000000000005", "0.5")]
res = I.monotonicity(near, GOV)
check(not res.violations and res.reversals == 1,
      "5: a reversal inside the allowance is counted, not failed")
res = I.p_plus_q([row("P_Cumulative", "1", "0.25", "0.25"),
                  row("P_Survival", "1", "0.75", "0.7501")], GOV)
check(len(res.violations) == 1, "5: P + Q off by 1E-4 fails")

if fails:
    print(f"FAIL: {len(fails)} accuracy-tier check(s)")
    for m in fails:
        print("  - " + m)
    raise SystemExit(1)
print("PASS: guarantees follow the #327 rule and are generated, never hand-edited; a target-only "
      "miss passes and is reported in the gate and the holdout; a guarantee breach fails both; "
      "silent wrong, monotonicity and P + Q violations fail independently of the tiers")
