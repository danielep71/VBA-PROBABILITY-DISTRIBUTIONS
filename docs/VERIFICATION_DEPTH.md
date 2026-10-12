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

The proof shim runs before `compute_errors.py`. This ordering is contractual. The former expected-red pre-export waiver is retired: the numerical gate is now green. Proofs still run first so a numerical or provenance failure cannot skip tests that the evaluators and guards reject controlled violations. At `bec5708`, the inventory contains 17 controls, including grid-regeneration safety, exact-SHA Excel certification, and the three static repository checks added during this audit. The root README assurance renderer (#28) is the 18th: its live command `render_readme_assurance.py --check` runs last in the Accuracy Gate, and `test_render_readme_assurance.py` runs in the proof shim.

## Positive path versus negative path

A live green check and a negative-control test answer different questions:

1. **Positive path:** does the current repository satisfy the rule?
2. **Negative path:** does the rule actually fail when its protected invariant is deliberately violated?

A release-blocking control needs both, except where the live repository is deliberately red pending external Excel evidence. Those cases are identified explicitly in `current_state`; they are not converted into PASS.

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

`excel-exact-sha-certification` is active and mandatory. Its portable negative proof is `benchmark/test_excel_certification.py`. Since #47 its live command also validates the retained `excel_regression_record.json` against each commit in the Accuracy Gate, so a mismatched retained record cannot coexist with a green gate. The runtime record must bind the candidate SHA, imported-source hashes, 909 regression assertions, Excel/build/bitness, and cleanup status. Fresh-grid export claims separately require matching committed grid digests and row counts.

**Current evidence state.** The Excel run, grid bindings and export certification, coverage, contract and holdout verdicts, and registered release blockers are generated from committed records in the [README assurance block](../README.md#assurance-at-a-glance) and checked by the Accuracy Gate on every change. They are not restated here.

See [Exact-SHA Excel certification](EXCEL_CERTIFICATION.md) for validation and current limitations.
