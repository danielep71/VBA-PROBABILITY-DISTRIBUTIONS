# Exact-SHA Excel certification

The Excel/VBA regression is release evidence only when it is bound to the exact
source that Excel executed. This repository therefore uses a machine-readable
certification record in addition to the human-readable regression log.

The contract is defined by `.github/excel-evidence-policy.json` and validated by
`benchmark/excel_certification.py` / `benchmark/check_excel_certification.py`.
The self-hosted workflow emits `artifacts/excel-vba-ci/excel-certification.json`.

The policy's `regression_sources` stays an explicit list, because the Excel
runner reads it and imports the modules in list order. `benchmark/check_module_lists.py`
fails when that list differs from the `.bas` files Git tracks under `src/` and
`tests/`, in either direction, or when such a file is present but untracked, so
a new module cannot escape the evidence binding. The comparison ignores order.

**Current evidence state.** The Excel run, grid bindings and export certification, coverage, contract and holdout verdicts, and registered release blockers are generated from committed records in the [README assurance block](../README.md#assurance-at-a-glance) and checked by the Accuracy Gate on every change. They are not restated here.

## When a change needs an Excel wave

The evidence a change needs depends on what it can break (#328). The same
table is in [CONTRIBUTING.md](../CONTRIBUTING.md#evidence-by-kind-of-change);
`tools/check_restated_counts.py` fails if the two differ.

<!-- BEGIN evidence rules -->
| Change | Evidence required |
| --- | --- |
| New numerical claim: a new kernel or route, a new or tighter guarantee | A preregistration with predictions, frozen before evaluation, and an untouched holdout |
| Defect fix inside existing guarantees | A regression test for the defect, and the main-grid export of the batch wave that certifies it |
| VBA change with identical code: comments, layout, a refactor proven equivalent | The existing certification, reused under #301; no wave |
| Python, CI or documentation only | The static checks and the evidence-tool suite |
<!-- END evidence rules -->

Excel waves are batched: one wave, dispatched on `release/*` with the
automated export below, certifies every change merged since the last one. The
evidence commit is the direct child of the wave's candidate, as always.

## Carried certification

**Decision (2026-11-13, #301).** A VBA change that provably alters only
comments and whitespace keeps the Excel certification and the grid bindings of
the commit Excel executed. No wave is run for it.

Reasons for:
- Excel executes the same code, so a new wave measures nothing new;
- one runner serves every wave, and a comment correction cost a full one
  (in #242 the wave on `fa891a2` was cancelled and rerun on `c856ddd` for
  comments only);
- documentation in the source stays current instead of waiting for a wave.

Reasons against, and the answer to each:
- *The proof depends on a lexer.* `benchmark/vba_code_identity.py` uses
  `tools/vba_lexer.py`, the lexer of the VBA source checks. It compares the
  comment-free logical lines of both versions. Outside string literals and
  `[...]` expressions it collapses runs of whitespace; it never inserts or
  removes whitespace between tokens. `test_vba_code_identity.py` covers a
  changed literal, a `Rem` line, a line continuation, an apostrophe in a
  string and a change inside `#If`, among others. Any doubt refuses the carry:
  a non-ASCII byte, an unreadable file, a missing commit, or recorded bytes
  that are not the commit's.
- *The record no longer binds the exact bytes in the tree.* The records are
  not rewritten. They still name the executed commit and its bytes. The
  verifiers prove from Git that the current module holds that commit's code,
  and every report says **carried from** that commit and lists the modules:
  the accuracy summary, the gate output and the README assurance block.

Limits:
- **A release is never certified on a carried binding.** The release workflow
  already requires an Excel run on the exact release commit; in addition,
  `render_readme_assurance.py` refuses `certified: true` while any module is
  carried.
- A module that uses a `[...]` expression outside a string literal is never
  carried, changed or not. Inside one, an apostrophe, a colon, `Rem` and a
  quote do not mean what they mean in a statement. No tracked module uses the
  construct.
- `check_excel_certification.py` accepts a carried source only with
  `--allow-carried`, which the Accuracy Gate and `refresh_evidence.py` pass.
  Release certification, a fresh export and `--candidate-sha` validation do
  not pass it and still require the exact bytes.
- A missing-only export still requires the exact bytes of the bound source.
- Study records that hash a VBA procedure or module keep their own rules.

To see what a change does before committing:
`python benchmark/vba_code_identity.py --from <executed commit>`.
After a comment-only change, run `python benchmark/refresh_evidence.py`: the
summaries and the README then carry the label.

## What a regression record proves

A green record binds:

- the full 40-character candidate Git SHA;
- SHA-256 of the canonical Git bytes for every imported VBA module;
- the policy-pinned CI entry point and exact assertion count;
- Excel version/build and actual Office executable bitness;
- Windows and runner identity plus workflow run/attempt identity;
- import, execution-backed compile, regression, and cleanup outcomes;
- the retained regression-log SHA-256.

The policy's `expected_assertions` is exact: a run with one assertion more or
one fewer is incomplete even if every executed assertion passes. The study
batteries, run apart since #300, are held to their own count in
`.github/excel-study-policy.json`.

`compile = PASS` in an automated record means the complete imported project was
accepted for execution by Excel and the full harness ran through
`Application.Run`. It is execution-backed compile evidence; it does not claim
that the VBE **Debug > Compile VBAProject** command was separately observed.
That explicit manual compile remains part of final release certification when
required by `RELEASING.md`.

## What it does not prove

The ordinary regression workflow does **not** export either numerical grid.
Its record therefore contains both policy-declared grids with:

```json
{
  "exported": false,
  "sha256": null,
  "row_count": null,
  "exported_utc": null
}
```

A green regression PASS cannot be used to rebind an accuracy manifest. This
separation is deliberate: regression execution and numerical observation export
are different evidence events. A dispatched export run produces both events in
one job, as two artifacts; see [Automated fresh export](#automated-fresh-export).

## Validate a downloaded workflow artifact

Extract `excel-certification.json` and `test-result.txt` into the same directory
and check it against the exact candidate checkout:

```bash
python benchmark/check_excel_certification.py \
  --record <artifact-dir>/excel-certification.json \
  --candidate-sha <full-40-character-sha> \
  --evidence-ref <full-40-character-sha> \
  --require-pass \
  --log-directory <artifact-dir>
```

Validation fails on, among other things, a short/wrong candidate SHA, changed
source bytes, wrong assertion count, failed cleanup, a changed/missing retained
log, or inconsistent harness counters.

<a id="automated-fresh-export"></a>

## Automated fresh export

The self-hosted runner exports the grids unattended (#181). The only manual
part is keeping the runner online; the evidence commit is still made by a
person, because the runner never writes to the repository.
[EXCEL_VBA_CI.md](EXCEL_VBA_CI.md#unattended-grid-export) describes what the
runner does. The steps below use the GitHub CLI.

1. **Dispatch** the Excel VBA Regression workflow on the commit to export.
   Select only the grids you mean to claim:

   ```bash
   gh workflow run excel-vba-regression.yml --ref <branch> \
     -f export_main=true -f export_holdout=true \
     -f export_holdout_older=false -f main_mode=full
   ```

   Study grids are named in one input, for example
   `-f export_studies=prs_gamma_probe,prs_chisquare_probe`, and applied with
   `--grid prs_gamma_probe --grid prs_chisquare_probe`.

   The run exports from the head of `<branch>` when it starts. Note that SHA:
   the evidence commit must be its child.

2. **Wait** for the run and **download** its export artifact into an empty
   directory outside the checkout:

   ```bash
   gh run list --workflow excel-vba-regression.yml --branch <branch> --limit 1
   gh run watch <run id>
   gh run download <run id> --name excel-grid-export-<run id> --dir <artifact-dir>
   ```

   A failed run still uploads the artifact; `export-manifest.json` and
   `export-log.txt` give the reason.

3. **Apply** it from a clean checkout of the exported commit:

   ```bash
   git switch <branch>     # its head must still be the exported SHA
   python benchmark/apply_excel_export.py --artifact <artifact-dir> \
     --grid main --grid holdout
   ```

   `--grid` names the grids you expect, and the tool refuses an artifact that
   holds others. It refuses, writing nothing, unless the export passed on the
   runner, the artifact names the full `HEAD`, the regression record is green
   for `HEAD` with its retained log, the environment file names `HEAD`, and
   every requested grid is present with the raw SHA-256 the manifest records.
   It then checks each grid against `HEAD` again: only the `observed_vba` cells
   its exporter owns may change, every one of them must be filled, and the
   bytes must be the CRLF lines Excel writes. A study grid (`export_studies`,
   #225) is written whole by its probe, so it must instead keep the committed
   header, line count and each row's key fields. A missing-only main export keeps
   the observations already in the grid, so it is accepted only while no `.bas`
   module differs from the `source_binding` in `observation_manifest.json`;
   otherwise run a full export. Then it:

   1. copies the grids and `excel_environment.json` into place;
   2. runs `finalize_excel_certification.py --from-fresh-export` with `--main`
      and/or `--holdout`, which records each grid's raw digest from the bytes
      Excel wrote;
   3. rewrites the grids with the LF bytes Git stores;
   4. runs `refresh_evidence.py` with `--bind-exported-main` and/or
      `--bind-exported-holdout`;
   5. validates the finalized record with `check_excel_certification.py`
      against a snapshot commit of the working tree (`git stash create`, which
      moves no ref), so the exact bytes you will commit are checked;
   6. prints the files to commit, the commit message and the validation
      commands for step 5.

   Before the evidence commit exists, the `retained Excel record` step of
   `refresh_evidence.py` compares the new record with `HEAD` and fails when a
   grid's bytes changed. The tool expects that one failure; the snapshot check
   replaces it. Any other failure leaves the files in place and ends with
   `NOT READY`. The usual one is a release gate that the export satisfies while
   `benchmark/release_readiness.json` still lists its blocker open: resolve that
   entry with this evidence, run `python benchmark/refresh_evidence.py` again
   (no bind flags; the manifests are already bound), and continue.

   `holdout_older`, `survival_boundary`, `tchi_large_df` and the study grids
   have no exporter in the evidence policy, so the record makes no claim for
   them and no manifest binds them.
   The tool copies the grid, and the plain refresh regenerates
   `holdout_older_summary.md`.

4. **Commit** the changed files together, with the printed message as the
   body.

5. **Validate** the evidence commit with the printed commands: for each
   claimed grid,

   ```bash
   python benchmark/check_excel_certification.py \
     --record benchmark/excel_regression_record.json \
     --candidate-sha <exported-sha> --evidence-ref HEAD --require-pass \
     --require-grid benchmark/probability_accuracy_grid.csv \
     --log-directory <artifact-dir>
   ```

   then `python benchmark/check_manifest_provenance.py --since <exported-sha>`
   and `python benchmark/refresh_evidence.py --check`. Push and open the pull
   request; the Accuracy Gate repeats the checks.

The dispatched run's regression artifact, `excel-vba-regression-<run id>`, is
the same as any other run's, so a successful export run also certifies its
commit for [release](#exact-tag-certification).

`excel_environment.json` from an automated export carries two fields that
`Export_ExcelEnvironment` does not write: `commit_sha`, the exported commit,
and `commit_sha_recorded_by`, set to `ci/Export-ExcelGrids.ps1`. The finalizer
refuses an environment file whose `commit_sha` is not the candidate.
`write_manifest.py` still records the manifest's `commit_sha` from the
checkout's `HEAD`, which the apply step requires to be the same commit.

## Manual fresh export (fallback)

Use this path when the runner is unavailable. It is intentionally separate and
must be run from the exact candidate checkout immediately after the selected
grid or grids were really exported in Excel. `benchmark/excel_environment.json`
must have been refreshed by the same Excel environment.

Start with the green regression record produced for that exact candidate, then:

```bash
python benchmark/finalize_excel_certification.py \
  --record <artifact-dir>/excel-certification.json \
  --main \
  --holdout \
  --from-fresh-export
```

Select only the grids actually exported. `--main` and `--holdout` may be used
independently.

The finalizer:

1. requires the record's candidate SHA to equal the current full `HEAD`;
2. revalidates the exact regression source inventory;
3. requires Excel version/build/bitness in `excel_environment.json` to match the
   regression record;
4. adds the relevant exporter module to the certified source inventory;
5. confirms Git stores each selected grid as LF, then records its canonical and raw SHA-256
   (see [Grid byte policy](#grid-byte-policy)), row count, and UTC timestamp;
6. revalidates the completed record structure, candidate and source inventory;
7. writes `benchmark/excel_regression_record.json`.

The finalizer cannot compare the new claims with a commit that does not exist yet, so finalizer success alone is insufficient for release certification: the evidence commit must pass the validation below.

It refuses to write without `--from-fresh-export`. The flag is an explicit
operator assertion; the Python tool cannot itself observe an Excel export.

After the record is finalized, bind the corresponding provenance manifest only
for grids that were actually exported:

```bash
python benchmark/refresh_evidence.py \
  --bind-exported-main \
  --bind-exported-holdout
```

The manifest guard evaluates commits separately. A manifest-only update can use
`benchmark/excel_regression_record.json` only when the canonical record is
green, still matches the evidence commit's VBA source bytes, and explicitly
claims that manifest's grid as freshly exported with matching committed SHA-256
and row count. A regression-only record is rejected.

After committing the evidence, validate each exported grid against that evidence commit (use the original runtime candidate SHA, not the evidence commit SHA):

```bash
python benchmark/check_excel_certification.py \
  --record benchmark/excel_regression_record.json \
  --candidate-sha <runtime-candidate-full-sha> \
  --evidence-ref <evidence-commit-full-sha> \
  --require-pass \
  --require-grid benchmark/probability_accuracy_grid.csv \
  --log-directory <retained-log-dir>
```

Repeat with `--require-grid benchmark/holdout/holdout_grid.csv` for the holdout. Retain the regression log beside the downloaded record or supply its directory as above. Any digest mismatch blocks certification. Do not hand-edit a digest to manufacture a passing record; the only permitted edit is a recorded correction (see [Corrections](#corrections)).

<a id="grid-byte-policy"></a>

## Grid byte policy

Excel writes CSV exports with CRLF line endings. `.gitattributes` stores every
`*.csv` as `text eol=lf`, so Git commits the same rows with LF line endings. An
exported grid entry therefore carries two exact digests:

| Field | Digest of |
| --- | --- |
| `sha256` | the canonical grid: the bytes Git stores, with each CRLF converted to LF |
| `raw_sha256` | the bytes Excel wrote, before Git normalized them |

Against a commit, the validator requires `sha256` to equal the committed blob
exactly, `raw_sha256` to equal one exact line-ending form of that blob (the blob
itself, or its CRLF form when the blob contains no CR), and `row_count` to match.
Neither digest is normalized before comparison, so a changed value, row or byte
other than a line ending fails both. Regression-only records written by the
Excel workflow carry no export claim; their grid entries may omit `raw_sha256`.

<a id="corrections"></a>

## Corrections

The retained record predates this policy: its finalizer stored the raw CRLF
digest as `sha256`, so it could never validate against the committed grid
(#47). It was reconciled from its own evidence, without a new run or export:
the recorded digest moved to `raw_sha256`, where it matches the CRLF form of the
committed grid exactly, and `sha256` names the committed LF bytes. The record's
`corrections` array documents the change: issue, date, field, recorded value,
corrected value and reason.

The validator accepts a correction only if it re-expresses a recorded export
digest in this way: the recorded value must be the grid's `raw_sha256` and the
corrected value its `sha256`. A correction cannot add an export, change a count
or alter any other field.

The Accuracy Gate validates the retained record on every change with:

```bash
python benchmark/check_excel_certification.py \
  --record benchmark/excel_regression_record.json --require-pass
```

## Evidence-only follow-up commits

The candidate SHA identifies the source revision actually executed in Excel.
A later commit may retain that record only while every certified VBA source byte
is unchanged. This permits an evidence-only commit without pretending that
Excel executed the evidence file itself. Any later change to a certified VBA
module invalidates the record and requires a new Excel run.

Retaining the record keeps its grid-export evidence valid; it does not certify
the later commit. A release needs an Excel run of the tagged commit itself, as
described next.

<a id="exact-tag-certification"></a>

## Release certification of the exact tagged commit

A commit cannot contain a record of itself, so the retained record always names
an earlier candidate than the commit that holds it. Validated without
`--candidate-sha`, it would also pass at any later commit with identical VBA
bytes, such as a documentation-only commit that Excel never ran (#76). The
checker says which binding it checked: `exact candidate: not required` or
`exact candidate: required and matched`.

The release workflow (`.github/workflows/release.yml`) therefore certifies each
release commit separately:

1. It takes the checked-out `HEAD^{commit}` as the candidate. On a tag push the
   candidate must equal the tag's peeled commit; on a rehearsal it is the commit
   of the dispatched ref.
2. It validates the retained record with `--evidence-ref <candidate>`, which
   proves the certified VBA bytes and grids are unchanged there.
3. `.github/scripts/certify_release_candidate.py` finds the successful `push`
   and `workflow_dispatch` runs of the Excel VBA Regression workflow whose head
   commit is the candidate, downloads each run's artifact, newest first, and
   validates its record against the candidate:

   ```bash
   python benchmark/check_excel_certification.py \
     --record <artifact-dir>/excel-certification.json \
     --candidate-sha <candidate> --evidence-ref <candidate> \
     --require-pass --log-directory <artifact-dir>
   ```

   The record must also name the self-hosted runner and the run it was
   downloaded from. Pull request runs are not used, because they execute a
   synthetic merge commit.
4. The first run that passes makes the candidate **CERTIFIED**. Otherwise it is
   **UNCERTIFIED**, and the job summary lists every reason. A tag push then
   fails before anything is built, attested or drafted. A rehearsal reports a
   warning, builds its practice bundle and never attests or releases it.

[RELEASING.md](../RELEASING.md#exact-tag-binding) describes how to produce
that run before tagging. `tests/test_certify_release_candidate.py` holds the
positive fixture (a green run of the exact commit) and the negative ones,
including a later commit with identical numerical source bytes.

## Failure and cleanup semantics

The record is written after cleanup. Import, compile, regression, and cleanup
are independent stages; `NOT_RUN`, `FAIL`, and `TIMEOUT` are never translated to
PASS. The runner closes only the workbook and Excel instance that it owns and
does not kill unrelated Excel processes.

A cleanup failure makes the overall run non-green even when all regression
assertions passed.

### Certified candidates outside the integration ancestry

The retained record can name an experimental Excel candidate that is not an
ancestor of the integration commit. Even a non-shallow single-branch clone may
lack that object. Run `python benchmark/check_excel_certification.py
--record benchmark/excel_regression_record.json --require-pass --fetch-candidate`
to fetch the exact recorded SHA from `origin` if absent. Accuracy and release
workflows opt in to this retrieval. Offline validation remains the default and
fails if the object is absent. Retrieval neither rewrites the record nor changes
the source/grid checks; an unavailable candidate remains a blocking failure.
The B6 run 447 candidate is also retained by branch
`evidence/26-b6-export-447` at `21941f31886883e6e2e58f863b142dbe6c2f97e6`;
keep this reference while its certification record is retained.

### Mixed exports and isolated tooling fixtures

A fresh export can change the main grid while reproducing the holdout bytes
exactly. The manifest writer validates the complete recorded export against a
snapshot of the pending tracked files before accepting the unchanged grid.
It does not move the branch or the stash reference. A mismatched companion grid
or source digest still fails validation; the export record and both manifests
must be committed atomically with their corresponding grids.

Positive tooling fixtures use a coherent historical evidence snapshot in
throwaway clones: `test_manifest_provenance.py` uses
`fdd57f84060ac0ed171edf8651fc1cc779b19674` (the release base), and
`test_render_readme_assurance.py` and `test_refresh_evidence.py` use
`8fd0af5b8fa3754d112c06f4845e2a896933c7b6`, the first commit whose summaries carry the guarantee and measured-target
columns (#327). Both are reachable from this change's ancestry. They run with the current
writer, renderer and refresh tools. Their success does not certify the live
checkout. CI separately checks live source identities, export records, grids,
counts and summaries, and the negative controls continue to reject stale or
contradictory evidence.
