"""
Permanent regression for the release gate's behavior when the high-precision
reference helper (_ibeta / mpmath) is unavailable.

The gate must NOT pass green with an ACTIVE contract left unevaluated. When the
helper is missing/corrupt, an active tail_probability_residual contract must be
reported PENDING (evaluator unavailable) and the gate must exit non-zero. Only
contracts explicitly marked status=characterization_only may be CHARACTERIZATION
ONLY. This test locks in the fix for the "active -> non-blocking characterization"
gate defect.

Run: python3 test_gate_degradation.py   (exit 0 = pass)
"""
import importlib.util, os, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))

def load_ce():
    spec = importlib.util.spec_from_file_location("ce_under_test",
                                                  os.path.join(HERE, "compute_errors.py"))
    ce = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(ce)
    return ce

def run(ce, allow_known=False):
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

    # Force the reference helper unavailable and simulate the exact import failure.
    ce = load_ce()
    ce._HAVE_IBETA = False
    ce._IBETA_IMPORT_ERROR = "ModuleNotFoundError: simulated missing _ibeta"

    for allow_known in (False, True):   # must block in BOTH strict and dev mode
        code, summary = run(ce, allow_known=allow_known)
        mode = "dev" if allow_known else "strict"

        # 1. Must exit NON-ZERO - never green with an active contract unevaluated.
        if code in (0, None):
            failures.append(f"[{mode}] gate passed (exit {code!r}) with the helper unavailable")

        # 2. Active tail contracts must be PENDING, not CHARACTERIZATION ONLY.
        tail_lines = [ln for ln in summary.splitlines() if "tail_probability_residual" in ln]
        if not tail_lines:
            failures.append(f"[{mode}] no tail_probability_residual contracts in summary")
        for ln in tail_lines:
            if "CHARACTERIZATION ONLY" in ln:
                failures.append(f"[{mode}] active tail contract wrongly CHARACTERIZATION ONLY: {ln.strip()}")
            if "PENDING" not in ln:
                failures.append(f"[{mode}] active tail contract not PENDING when unevaluable: {ln.strip()}")

        # 3. Non-tail contracts must still be evaluated (gate keeps working).
        if "PASS" not in summary:
            failures.append(f"[{mode}] no other contracts evaluated with the helper unavailable")

    if failures:
        print("FAIL - gate did not block correctly on an unevaluated active contract:")
        for f in failures:
            print("  -", f)
        sys.exit(1)
    print("PASS - active tail contracts -> PENDING and the gate exits non-zero in both "
          "strict and dev mode; other contracts still evaluated; CHARACTERIZATION ONLY not misused.")
    sys.exit(0)

if __name__ == "__main__":
    main()
