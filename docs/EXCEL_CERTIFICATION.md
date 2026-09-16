# Exact-SHA Excel certification

The Excel/VBA regression is release evidence only when it is bound to the exact
source that Excel executed. This repository therefore uses a machine-readable
certification record in addition to the human-readable regression log.

The contract is defined by `.github/excel-evidence-policy.json` and validated by
`benchmark/excel_certification.py` / `benchmark/check_excel_certification.py`.
The self-hosted workflow emits `artifacts/excel-vba-ci/excel-certification.json`.

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
5. records each freshly exported grid's SHA-256, row count, and UTC timestamp;
6. writes `benchmark/excel_regression_record.json`;
7. revalidates the completed record before returning success.

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
