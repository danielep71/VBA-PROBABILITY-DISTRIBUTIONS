# Verification-depth contract

`benchmark/verification_depth.json` is the v1.0.0 inventory of controls whose behavior can block or certify release evidence. It exists to prevent a false-green assurance layer: a checker is not considered adequately protected merely because it exists or because its current input happens to pass.

## Required proof

Every inventoried control records:

- its stable control id;
- the live checker command;
- an executable proof command;
- the material negative behavior that proof exercises; and
- the expected current repository state.

`benchmark/check_verification_depth.py` fails when an entry is incomplete, duplicated, references a missing script, uses an unknown state, disappears from the mandatory v1.0.0 set, or names a proof that is not executed by `benchmark/test_evidence_tools.py`.

When full regression proofs are selected, the proof shim runs before `compute_errors.py`. This ordering is contractual. The former expected-red pre-export waiver is retired: the numerical gate is now green. Proofs still run first so a numerical or provenance failure cannot skip tests that the evaluators and guards reject controlled violations. At `bec5708`, the inventory contains 17 controls, including grid-regeneration safety, exact-SHA Excel certification, and three static repository checks: committed whitespace, VBA procedure jumps and release semantics. The root README assurance renderer (#28) is the 18th: its live command `render_readme_assurance.py --check` runs last in the Accuracy Gate, and `test_render_readme_assurance.py` runs in the proof shim. The module-list binding (#103) is the 19th: `check_module_lists.py` binds the Excel evidence policy's `regression_sources` to the tracked `src/` and `tests/` modules, and `test_module_lists.py` proves that a missing, extra, untracked or duplicated module is rejected. The CI test-coverage guard (#104) is the 20th: `tools/check_ci_test_coverage.py` fails when a tracked `test_*.py` file is executed by no pull-request workflow step and no runner list, and `tests/test_ci_test_coverage.py` proves it against fixture repositories. The df-cap mirror check (#86) is the 21st: `check_df_cap_mirrors.py` parses `PROB_F_MAX_DF`, `PROB_T_MAX_DF` and `PROB_CHI_MAX_DF` from `src/M_STATS_PROBDIST_TFAMILY.bas`, converts each literal to its binary64 Double and fails unless `F_MAX_DF`, `T_MAX_DF` and `CHI_MAX_DF` in `_contract_eval.py` match bit for bit, and `test_df_cap_mirrors.py` proves that a changed value on either side, a missing, renamed or duplicate constant, and an unparsable or inexact literal are rejected. The older-families holdout summary check (#164) is the 22nd: `holdout_older/check_holdout_older_summary.py` re-renders `holdout_older_summary.md` from the committed grid and contract registry and fails when the committed file differs, and `holdout_older/test_analyze_holdout_older.py` proves that a contract tightened or retired after the summary was written, a hand-edited verdict and a missing summary are rejected. The check is about currency: a current summary that records FAIL verdicts passes. Since #171 it also fails when a grid row is neither scored by a contract with a threshold nor covered by a `characterization_only` contract, and the same proof rejects such a row. The survival-reference audit (#159) is the 23rd: `check_survival_references.py` recomputes every committed survival reference by a direct upper-tail route at the exact binary64 argument and fails when one misses the 25-digit floor without a registered debt entry in `survival_reference_debt.json`, and `test_survival_references.py` proves that a reference taken at the decimal token, one formed as 1 - CDF, stale or orphaned debt and an unlisted grid are rejected. Its state is `green_with_registered_transition_debt`. It is in the inventory and both commands run in the proof shim, but it is not in the checker's mandatory `required` set. The accuracy tiers (#327) are the 24th: `accuracy_tiers.py` fails unless `accuracy_guarantees.csv` is exactly the output of the guarantee rule over the contract registry, and `test_accuracy_tiers.py` proves that a hand-edited, stale or too-strict guarantee is refused, that a guarantee breach fails both the strict gate and the holdout analyzer while a miss of the measured target alone passes, and that silent wrong values, monotonicity reversals and `P + Q` mismatches fail whatever the tier. `python benchmark/check_verification_depth.py` prints the current count.

## Guarantee and measured target

Since #327 the strict gate and both holdout analyzers score each contract against its **guarantee**, derived by `benchmark/accuracy_tiers.py`, and report its **measured target**, the registry threshold, beside it. A target miss inside the guarantee is visible in every summary as "PASS · target missed" and never blocks. The always-gating checks in `benchmark/_invariants.py` (silent wrong values, monotonicity, `P + Q`) run on every row of each grid, including rows that no gating contract scores.

## Positive path versus negative path

A live green check and a negative-control test answer different questions:

1. **Positive path:** does the current repository satisfy the rule?
2. **Negative path:** does the rule actually fail when its protected invariant is deliberately violated?

A release-blocking control needs both. The checker accepts only green `current_state` values; a red control is not recorded as PASS.

Examples of controlled violations include public-API signature drift, manifest-only evidence rebinding, deleted coverage rows, missing README sections, unavailable reference helpers, stale generated tables, restated measured thresholds, incomplete-gamma dispatch drift and corrupted Student-t coefficient authority.

## Adding a release-blocking checker

A new checker is not complete until the same change:

1. adds it to `verification_depth.json` **and** the mandatory control set in `check_verification_depth.py`;
2. provides at least one meaningful mutation or malformed-input case that must fail;
3. makes that proof command exit non-zero when the checker fails to detect the mutation;
4. wires the proof into `test_evidence_tools.py`, before the strict numerical verdict; and
5. runs `python benchmark/check_verification_depth.py` successfully.

Utilities, exploratory studies and report generators that cannot affect a release verdict do not belong in this inventory merely because they exist.

## Excel exact-SHA certification

`excel-exact-sha-certification` is active and mandatory. Its portable negative proof is `benchmark/test_excel_certification.py`. Since #47 its live command also validates the retained `excel_regression_record.json` against each commit in the Accuracy Gate, so a mismatched retained record cannot coexist with a green gate. The runtime record must bind the candidate SHA, imported-source hashes, the exact regression assertion count in `.github/excel-evidence-policy.json` (`expected_assertions`), Excel/build/bitness, and cleanup status. Fresh-grid export claims separately require matching committed grid digests and row counts.

**Current evidence state.** The Excel run, grid bindings and export certification, coverage, contract and holdout verdicts, and registered release blockers are generated from committed records in the [README assurance block](../README.md#assurance-at-a-glance) and checked by the Accuracy Gate on every change. They are not restated here.

See [Exact-SHA Excel certification](EXCEL_CERTIFICATION.md) for validation and current limitations.

## Conditional CI execution

I keep the `Strict accuracy gate (pure Python, no Excel)` status on every
pull request and main push. Only the `Protect main` ruleset requires it; release
branches require only `Repository integrity (static, no Excel)`. It aggregates
the selection job, live evidence job, and applicable positive-ratio study job; a
failed or cancelled prerequisite cannot produce a successful gate. The study
runs in parallel with the core job.

For changes confined to the explicit issue/label metadata allowlist in
`tools/select_accuracy_checks.py`, I run the proof shim with `--checks-only`.
This executes the live `check_*` commands from the same authoritative command
list, excluding `--self-test` fixtures. All later live numerical, summary,
provenance and certification steps remain enabled. The separate study is skipped.
Local `test_evidence_tools.py` and `refresh_evidence.py` still run the full suite
by default. Every executed shim command reports elapsed time.

The selection checks the complete base-to-candidate diff and intervening history,
including deletions and both paths of renames. README, numerical documentation,
source, benchmark data, dependencies, workflows other than the explicitly listed
label/triage workflows, and unknown paths all trigger full proofs. Missing history,
non-ancestral comparisons and empty change sets also run full proofs. A failure
of the selection job itself blocks the required gate.

Manual runs and the Monday 03:23 UTC schedule always run the full suite. This
conservative policy deliberately reruns studies for some unrelated code changes;
I only expand the skip list after verifying that the files are not proof inputs.
