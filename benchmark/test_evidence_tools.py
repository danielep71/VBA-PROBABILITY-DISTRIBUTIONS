"""Run every release-blocking tooling proof before the strict numerical gate.

This shim is deliberately the single hosted execution list for verification-depth
proofs. A red numerical evidence state must not skip selected tests of the machinery
that produces or protects that evidence. Local runs are always full by default;
CI may request live checks only for an audited metadata-only change.
"""
import argparse
import os
import time
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
failures = []
commands = (
    ("test_contract_eval.py",),
    ("test_grid_coverage.py",),
    ("check_grid_coverage.py", "--mode", "auto", "--check-summary"),
    ("test_grid_regeneration.py",),
    ("test_binary64.py",),
    ("check_grid_keys.py",),
    ("test_root_readme.py",),
    ("check_root_readme.py",),
    ("test_render_readme_assurance.py",),
    ("test_igamma_parity.py",),
    ("student_t_large_df_study/test_coefficients.py",),
    ("student_t_large_df_study/test_binary64_route_model.py",),
    ("student_t_large_df_study/test_model_fidelity.py",),
    ("normal_tail_study/test_kernel_model.py",),
    ("normal_tail_study/test_holdout_record.py",),
    ("normal_tail_study/test_conditioning.py",),
    ("normal_tail_study/test_fit_convergence.py",),
    ("normal_tail_study/test_result_binding.py",),
    ("test_excel_certification.py",),
    ("test_drift_guard_mutations.py",),
    ("test_public_api.py",),
    ("check_public_api.py",),
    ("test_module_lists.py",),
    ("check_module_lists.py",),
    ("test_manifest.py",),
    ("test_manifest_provenance.py",),
    ("check_manifest_provenance.py",),
    ("test_gate_degradation.py",),
    ("test_render_contract_table.py",),
    ("test_source_thresholds.py",),
    ("check_source_thresholds.py",),
    ("holdout/test_analyze_holdout.py",),
    ("../tools/check_committed_whitespace.py", "--root", "..", "--self-test"),
    ("../tools/check_vba_jumps.py", "--root", "..", "--self-test"),
    ("../tools/check_release_semantics.py", "--root", "..", "--self-test"),
    ("../tests/test_ci_test_coverage.py",),
    ("../tools/check_ci_test_coverage.py", "--root", ".."),
    ("test_verification_depth.py",),
    ("check_verification_depth.py",),
)
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--checks-only', action='store_true',
                    help='Validate live evidence; omit regression fixtures only for metadata-only CI')
args = parser.parse_args()
for command in commands:
    if args.checks_only and (not os.path.basename(command[0]).startswith('check_')
                             or '--self-test' in command):
        continue
    started = time.monotonic()
    proc = subprocess.run([sys.executable] + list(command), cwd=HERE)
    print(f"TIMING {time.monotonic() - started:.2f}s: {' '.join(command)}", flush=True)
    if proc.returncode:
        failures.append(" ".join(command))
if failures:
    print("FAIL: evidence-tool unit tests: " + ", ".join(failures))
    raise SystemExit(1)
print("PASS: live evidence checks" if args.checks_only else
      "PASS: evidence-tool unit tests and verification-depth contract")
