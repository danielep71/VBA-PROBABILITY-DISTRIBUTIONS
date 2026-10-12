# Exact-SHA Excel certification

The Excel/VBA regression is release evidence only when it is bound to the exact
source that Excel executed. This repository therefore uses a machine-readable
certification record in addition to the human-readable regression log.

The contract is defined by `.github/excel-evidence-policy.json` and validated by
`benchmark/excel_certification.py` / `benchmark/check_excel_certification.py`.
The self-hosted workflow emits `artifacts/excel-vba-ci/excel-certification.json`.

**Current evidence state.** The Excel run, grid bindings and export certification, coverage, contract and holdout verdicts, and registered release blockers are generated from committed records in the [README assurance block](../README.md#assurance-at-a-glance) and checked by the Accuracy Gate on every change. They are not restated here.

## What a regression record proves

A green record binds:

- the full 40-character candidate Git SHA;
- SHA-256 of the canonical Git bytes for every imported VBA module;
- the policy-pinned CI entry point and exact assertion count;
- Excel version/build and actual Office executable bitness;
- Windows and runner identity plus workflow run/attempt identity;
- import, execution-backed compile, regression, and cleanup outcomes;
- the retained regression-log SHA-256.

The policy currently requires 909 assertions. A run with 908 or 910 assertions
is incomplete even if every executed assertion passes.

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

A 909/909 regression PASS cannot be used to rebind an accuracy manifest. This
separation is deliberate: regression execution and numerical observation export
are different evidence events.

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

## Add fresh grid-export claims

This step is intentionally separate and must be run from the exact candidate
checkout immediately after the selected grid or grids were really exported in
Excel. `benchmark/excel_environment.json` must have been refreshed by the same
Excel environment.

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

## Failure and cleanup semantics

The record is written after cleanup. Import, compile, regression, and cleanup
are independent stages; `NOT_RUN`, `FAIL`, and `TIMEOUT` are never translated to
PASS. The runner closes only the workbook and Excel instance that it owns and
does not kill unrelated Excel processes.

A cleanup failure makes the overall run non-green even when all regression
assertions passed.
