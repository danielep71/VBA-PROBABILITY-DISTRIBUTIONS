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
    ("f_inverse_large_df_study/fix_comparison/test_pair_cf_screen.py",),
    ("f_inverse_large_df_study/fix_comparison/test_vba_order_model.py",),
    ("f_inverse_large_df_study/fix_comparison/test_residual_diagnostic.py",),
    ("f_inverse_large_df_study/fix_comparison/test_tail_aware.py",),
    ("test_contract_eval.py",),
    ("test_accuracy_tiers.py",),
    ("test_vba_code_identity.py",),
    ("test_study_grid_binding.py",),
    ("check_study_grid_binding.py",),
    ("test_frozen_sources.py",),
    ("gamma_cf_study/coverage/test_production_follows_study.py",),
    ("test_harness_timing.py",),
    ("test_harness_studies.py",),
    ("positive_ratio_subnormal/check_post23_cutoffs.py",),
    ("test_grid_coverage.py",),
    ("test_cross_check_scipy.py",),
    ("check_grid_coverage.py", "--mode", "auto", "--check-summary"),
    ("test_grid_regeneration.py",),
    ("test_binary64.py",),
    ("check_grid_keys.py",),
    ("test_root_readme.py",),
    ("check_root_readme.py",),
    ("test_render_readme_assurance.py",),
    ("test_refresh_evidence.py",),
    ("test_igamma_parity.py",),
    ("student_t_large_df_study/test_coefficients.py",),
    ("student_t_large_df_study/test_binary64_route_model.py",),
    ("student_t_large_df_study/test_model_fidelity.py",),
    ("student_t_large_df_study/test_production_route.py",),
    ("student_t_large_df_study/test_vba_large_df_source.py",),
    ("student_t_large_df_study/test_inverse_polish.py",),
    ("student_t_large_df_study/test_above8.py",),
    ("tchi_large_df/test_tchi_merge.py",),
    ("gamma_cf_study/coverage/test_precision_probe.py",),
    ("gamma_cf_study/coverage/test_precision_v2.py",),
    ("gamma_cf_study/coverage/test_research_checks.py",),
    ("gamma_cf_study/coverage/test_precision_arithmetic_audit.py",),
    ("gamma_cf_study/coverage/test_cf_enclosure.py",),
    ("gamma_cf_study/coverage/test_scaled_product.py",),
    ("gamma_cf_study/coverage/test_scaled_dd.py",),
    ("gamma_cf_study/coverage/test_scaled_gamma.py",),
    ("gamma_cf_study/coverage/test_scaled_gamma_vba_port.py",),
    ("gamma_cf_study/coverage/test_extended_gamma.py",),
    ("gamma_cf_study/coverage/test_wide_gamma.py",),
    ("gamma_cf_study/coverage/test_public_integration.py",),
    ("gamma_cf_study/coverage/test_public_fresh.py",),
    ("gamma_cf_study/coverage/test_scaled_dd_vba_port.py",),
    ("gamma_cf_study/coverage/test_scaled_vba_port.py",),
    ("gamma_cf_study/replay_baseline.py", "isolation"),
    ("gamma_cf_study/replay_baseline.py", "candidate"),
    ("gamma_cf_study/test_promote_x2_ladder.py",),
    ("gamma_series_small_shape/test_gamma_series_small_shape.py",),
    ("gamma_series_small_shape/analyze_gamma_series_small_shape.py", "--check"),
    ("gamma_lower_q/test_gamma_lower_q.py",),
    ("gamma_lower_q/test_promote_x05_ladder.py",),
    ("gamma_lower_q/analyze_gamma_lower_q.py", "--check"),
    ("gamma_route_study/test_binary64_map.py",),
    ("gamma_route_study/analyze_binary64_map.py", "--check"),
    ("gamma_route_study/check_disjoint.py",),
    ("gamma_route_study/test_route_model.py",),
    ("gamma_route_study/test_b64_fit.py",),
    ("gamma_route_study/test_pair_candidate.py",),
    ("gamma_route_study/test_scalar_candidate.py",),
    ("gamma_route_study/test_holdout_route.py",),
    ("gamma_route_study/holdout_route_build.py", "--check"),
    ("gamma_route_study/replay_fast_route.py", "--check"),
    ("gamma_route_study/holdout_route_analyze.py",),
    ("gamma_inverse_scale_study/test_holdout.py",),
    ("gamma_inverse_scale_study/holdout_build.py", "--check"),
    ("gamma_inverse_scale_study/holdout_build.py", "--check", "--holdout", "2"),
    ("gamma_inverse_scale_study/holdout_analyze.py", "--holdout", "2"),
    ("f_inverse_large_df_study/production/test_production.py",),
    ("f_inverse_large_df_study/production/build_grid.py", "--check"),
    ("f_inverse_large_df_study/production/analyze.py",),
    ("gamma_cache_study/test_cache_vba.py",),
    ("gamma_recurrence_study/test_prepare.py",),
    ("gamma_route_study/analyze_b64_fit.py", "--check"),
    ("normal_tail_study/test_kernel_model.py",),
    ("normal_tail_study/test_vba_kernel_source.py",),
    ("normal_tail_study/test_holdout_record.py",),
    ("normal_tail_study/test_conditioning.py",),
    ("normal_tail_study/test_fit_convergence.py",),
    ("normal_tail_study/test_result_binding.py",),
    ("normal_tail_study/test_ulps_apart.py",),
    ("beta_balanced_survival/test_beta_route_model.py",),
    ("beta_balanced_survival/test_promote_worst_points.py",),
    ("test_excel_certification.py",),
    ("test_apply_excel_export.py",),
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
    ("test_df_cap_mirrors.py",),
    ("check_df_cap_mirrors.py",),
    ("test_survival_references.py",),
    ("cdf_large_shape/test_generate_cdf_large_shape.py",),
    ("check_survival_references.py",),
    ("holdout/test_analyze_holdout.py",),
    ("holdout_older/test_analyze_holdout_older.py",),
    ("holdout_older/check_holdout_older_summary.py",),
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
    if args.checks_only and command != ('gamma_cf_study/replay_baseline.py', 'candidate') and (not os.path.basename(command[0]).startswith('check_')
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
