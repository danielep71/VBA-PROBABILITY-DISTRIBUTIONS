"""
Permanent regression for the release gate's behavior when the high-precision
reference helper (_ibeta / mpmath) is unavailable.

This is a TOOLING fixture, not release evidence. It deliberately supplies a
synthetic clean provenance seam so the intended degradation path is reachable
even while the real repository is truthfully red for stale Excel observations.
Otherwise Phase-0 stale provenance masks the behavior this test exists to prove.

The gate must NOT pass green with an ACTIVE contract left unevaluated. When the
helper is missing/corrupt, an active tail_probability_residual contract must be
reported PENDING (evaluator unavailable) and the gate must exit non-zero. Only
contracts explicitly marked status=characterization_only may be CHARACTERIZATION
ONLY. This locks in the fix for the "active -> non-blocking characterization"
gate defect.

Run: python3 test_gate_degradation.py   (exit 0 = pass)
"""
import importlib.util
import os
import sys
import tempfile
import types

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)


def load_ce():
    spec = importlib.util.spec_from_file_location("ce_under_test",
                                                  os.path.join(HERE, "compute_errors.py"))
    ce = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(ce)
    return ce


def clean_manifest_seam():
    """Synthetic _manifest module that proves only this fixture's gate path.

    Manifest correctness has its own independent mutation suite in test_manifest.py
    and test_manifest_provenance.py. Requiring today's live stale evidence here
    makes this helper-degradation proof unreachable and conflates two controls.
    """
    fake = types.ModuleType("_manifest")
    fake.MANIFEST_NAME = "observation_manifest.json"
    fake.HOLDOUT_MANIFEST_NAME = "holdout_manifest.json"
    fake.repo_root = lambda: ROOT
    fake.load_manifest = lambda _root: {}
    fake.verify_source_binding = lambda _root, _manifest, _grid, _contracts: []
    fake.load_holdout_manifest = lambda _root: {}
    fake.verify_holdout_binding = lambda _root, _manifest, _grid, _contracts: []
    return fake


def run(ce, allow_known=False):
    out = tempfile.NamedTemporaryFile("w", suffix=".md", delete=False).name
    argv = ["compute_errors.py", "--grid", os.path.join(HERE, "probability_accuracy_grid.csv"),
            "--out", out]
    if allow_known:
        argv.append("--allow-known-limitations")
    old_argv = sys.argv
    old_manifest = sys.modules.get("_manifest")
    sys.argv = argv
    sys.modules["_manifest"] = clean_manifest_seam()
    code = None
    try:
        ce.main()
    except SystemExit as exc:
        code = exc.code
    finally:
        sys.argv = old_argv
        if old_manifest is None:
            sys.modules.pop("_manifest", None)
        else:
            sys.modules["_manifest"] = old_manifest
    with open(out, encoding="utf-8") as f:
        summary = f.read()
    os.unlink(out)
    return code, summary


def main():
    failures = []

    # Force the reference helper unavailable and simulate the exact import failure.
    ce = load_ce()
    ce._HAVE_IBETA = False
    ce._TAIL_CDFS = {}
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
        for line in tail_lines:
            if "CHARACTERIZATION ONLY" in line:
                failures.append(f"[{mode}] active tail contract wrongly CHARACTERIZATION ONLY: {line.strip()}")
            if "PENDING" not in line:
                failures.append(f"[{mode}] active tail contract not PENDING when unevaluable: {line.strip()}")

        # 3. Non-tail contracts must still be evaluated (gate keeps working).
        if "PASS" not in summary:
            failures.append(f"[{mode}] no other contracts evaluated with the helper unavailable")

    if failures:
        print("FAIL - gate did not block correctly on an unevaluated active contract:")
        for failure in failures:
            print("  -", failure)
        sys.exit(1)
    print("PASS - active tail contracts -> PENDING and the gate exits non-zero in both "
          "strict and dev mode; other contracts still evaluated; CHARACTERIZATION ONLY not misused; "
          "live stale provenance is isolated by a synthetic clean provenance seam.")
    sys.exit(0)


if __name__ == "__main__":
    main()
