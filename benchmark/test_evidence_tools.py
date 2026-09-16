"""Run every release-blocking tooling proof before the strict numerical gate.

This shim is deliberately the single hosted execution list for verification-depth
proofs. A red numerical evidence state must not skip tests of the machinery that
produces or protects that evidence.
"""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
failures = []
commands = (
    ("test_contract_eval.py",),
    ("test_grid_coverage.py",),
    ("check_grid_coverage.py", "--mode", "auto", "--check-summary"),
    ("test_root_readme.py",),
    ("check_root_readme.py",),
    ("test_igamma_parity.py",),
    ("student_t_large_df_study/test_coefficients.py",),
    ("test_excel_certification.py",),
    ("test_drift_guard_mutations.py",),
    ("test_public_api.py",),
    ("check_public_api.py",),
    ("test_manifest.py",),
    ("test_manifest_provenance.py",),
    ("check_manifest_provenance.py",),
    ("test_gate_degradation.py",),
    ("test_render_contract_table.py",),
    ("test_source_thresholds.py",),
    ("check_source_thresholds.py",),
    ("holdout/test_analyze_holdout.py",),
    ("check_verification_depth.py",),
)
for command in commands:
    proc = subprocess.run([sys.executable] + list(command), cwd=HERE)
    if proc.returncode:
        failures.append(" ".join(command))
if failures:
    print("FAIL: evidence-tool unit tests: " + ", ".join(failures))
    raise SystemExit(1)
print("PASS: evidence-tool unit tests and verification-depth contract")
