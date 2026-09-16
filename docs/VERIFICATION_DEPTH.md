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

The proof shim runs before `compute_errors.py`. This ordering is contractual. During the current pre-export phase the strict numerical gate is intentionally red because committed Excel observations are stale relative to changed source. That must not prevent the repository from proving that its evaluators, provenance guards, compatibility guard and documentation guards still reject controlled violations.

## Positive path versus negative path

A live green check and a negative-control test answer different questions:

1. **Positive path:** does the current repository satisfy the rule?
2. **Negative path:** does the rule actually fail when its protected invariant is deliberately violated?

A release-blocking control needs both, except where the live repository is deliberately red pending external Excel evidence. Those cases are identified explicitly in `current_state`; they are not converted into PASS.

Examples of controlled violations include public-API signature drift, manifest-only evidence rebinding, deleted coverage rows, missing README sections, unavailable reference helpers, stale generated tables, restated measured thresholds, incomplete-gamma dispatch drift and corrupted Student-t coefficient authority.

## Adding a release-blocking checker

A new checker is not complete until the same change:

1. adds it to `verification_depth.json`;
2. provides at least one meaningful mutation or malformed-input case that must fail;
3. makes that proof command exit non-zero when the checker fails to detect the mutation;
4. wires the proof into `test_evidence_tools.py`, before the strict numerical verdict; and
5. runs `python benchmark/check_verification_depth.py` successfully.

Utilities, exploratory studies and report generators that cannot affect a release verdict do not belong in this inventory merely because they exist.

## Excel certification

P0.1 / issue #38 is being implemented separately. Its exact-SHA Excel certification validator already carries dedicated negative fixtures. When that control is merged into `main`, it becomes release-blocking and must be added to this inventory in the same integration change. The matrix must never list an unmerged checker as though it were active.

## Excel exact-SHA certification

`excel-exact-sha-certification` is a release-blocking assurance control. Its portable negative proof is `benchmark/test_excel_certification.py`; a green portable proof does not substitute for the self-hosted Excel runtime record, which remains pending until the runner executes this candidate. The live runtime record must bind the full candidate SHA, exact imported-source hashes, 909 regression assertions, Excel/build/bitness, and cleanup status. Fresh-grid export claims remain separate and require explicit matching grid digest and row-count evidence.
