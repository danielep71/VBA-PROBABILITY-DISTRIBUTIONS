"""Run evaluator and grid-coverage fixtures under one stable gate label."""
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
    ("test_manifest_provenance.py",),
    ("check_manifest_provenance.py",),
)
for command in commands:
    proc = subprocess.run([sys.executable] + list(command), cwd=HERE)
    if proc.returncode:
        failures.append(" ".join(command))
if failures:
    print("FAIL: evidence-tool unit tests: " + ", ".join(failures))
    raise SystemExit(1)
print("PASS: evidence-tool unit tests")
