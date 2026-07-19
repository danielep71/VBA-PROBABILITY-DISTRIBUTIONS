"""
Permanent regression: the accuracy release gate must DEGRADE GRACEFULLY when the
high-precision reference helper (_ibeta / mpmath) is unavailable.

tail_probability_residual contracts require the true CDF at the VBA quantile,
which needs mpmath. When it is absent the gate must NOT crash: those contracts
fall back to CHARACTERIZATION ONLY (they remain verified in their study
directory), while every other contract is still evaluated normally.

Run: python3 test_gate_degradation.py   (exit 0 = pass)
"""
import importlib.util, os, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))

def load_compute_errors():
    spec = importlib.util.spec_from_file_location("ce_under_test",
                                                  os.path.join(HERE, "compute_errors.py"))
    ce = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(ce)
    return ce

def run(ce, allow_known=True):
    out = tempfile.NamedTemporaryFile("w", suffix=".md", delete=False).name
    argv = ["compute_errors.py", "--grid", os.path.join(HERE, "probability_accuracy_grid.csv"),
            "--out", out]
    if allow_known:
        argv.append("--allow-known-limitations")
    old = sys.argv; sys.argv = argv
    code = None
    try:
        ce.main()
    except SystemExit as e:
        code = e.code
    finally:
        sys.argv = old
    return code, open(out, encoding="utf-8").read()

def main():
    failures = []

    # 1. Baseline: helper available -> tail contracts should PASS (if mpmath present).
    ce = load_compute_errors()
    have = ce._HAVE_IBETA

    # 2. Degraded: force the helper unavailable and re-run.
    ce2 = load_compute_errors()
    ce2._HAVE_IBETA = False
    code, summary = run(ce2)

    # Must not crash (a normal SystemExit with an int/None code is fine).
    if code not in (0, 1, 2, None):
        failures.append(f"gate crashed or returned unexpected exit code: {code!r}")

    # tail_probability_residual contracts must be CHARACTERIZATION ONLY, not FAIL.
    tail_lines = [ln for ln in summary.splitlines() if "tail_probability_residual" in ln]
    if not tail_lines:
        failures.append("no tail_probability_residual contracts found in summary")
    for ln in tail_lines:
        if "CHARACTERIZATION ONLY" not in ln:
            failures.append(f"tail contract did not degrade to CHARACTERIZATION ONLY: {ln.strip()}")
        if "FAIL" in ln:
            failures.append(f"tail contract wrongly FAILED without the helper: {ln.strip()}")

    # Non-tail contracts must still be evaluated (at least one PASS present).
    if "PASS" not in summary:
        failures.append("no contracts evaluated with the helper unavailable")

    print(f"reference helper available in this environment: {have}")
    if failures:
        print("FAIL - gate did not degrade gracefully:")
        for f in failures:
            print("  -", f)
        sys.exit(1)
    print("PASS - gate degrades gracefully: tail contracts -> CHARACTERIZATION ONLY, "
          "others still evaluated, no crash.")
    sys.exit(0)

if __name__ == "__main__":
    main()
