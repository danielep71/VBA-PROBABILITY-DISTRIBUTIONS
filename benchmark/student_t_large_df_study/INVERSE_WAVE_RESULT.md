# Excel wave result — Student-t inverse polish and contracts (#34)

This is the wave of `INVERSE_PREREGISTRATION.md` §4. `check_inverse_wave.py
--baseline f1019e9` checked it against the predictions that
`inverse_predictions.json` recorded before any VBA. Its output is
`inverse_wave_check.json`.

## The export

- **Run.** Excel VBA Regression run 42260290029, candidate `9cbe5d9`.
- **Regression.** 955/955, including the five assertions of
  `Test_TF_StudentTLargeDFQuantile` at 5E-16.
- **Grids.** The main grid, `tchi_large_df`, and the holdout, with the 56
  untouched rows of §6, each exported once.
- **Evidence commit.** `evidence: Excel export of main, holdout, tchi_large_df
  at 9cbe5d9`.

## The predictions

| Prediction | Rows | Result |
| --- | ---: | --- |
| PI, polished inverse: observation in its recorded outcome set, error within the quantile bound | 30 | **holds**: every observation is a member; worst error 2.81E-16 |
| PU, every other Student-t, density and F row bit-identical | 361 + 144 + 111 | **holds**: none changed |
| PH, untouched holdout under the PT1 and PI rules | 56 | **holds**: worst error 2.54E-16 |
| Regression green | 955 | **holds** |

The 7 rows the first wave's PT2 failed were 6.9E-16 to 9.5E-16 from their
quantiles. They now lie within their 1- or 2-Double outcome sets.

## Consequences under §7

Every prediction holds, so the three contracts frozen in
`PRODUCTION_PREREGISTRATION.md` §7 and committed before the export are adopted:

| Contract | Threshold and domain |
| --- | --- |
| `StudentT_Survival.envelope_domain.output_rel` | 5E-15, df 1E5..1E8, \|t\| ≤ 8 |
| `StudentT_Cumulative.envelope_domain.output_rel` | 5E-15, df 1E5..1E8, \|t\| ≤ 8 |
| `StudentT_InverseCumulative.envelope_domain.quantile_rel` | 5E-15, df 1E5..1E8, \|p − 0.5\| ≥ 0.1, \|q\| ≤ 8 |

- **Coverage.** The nine Student-t inverse `envelope_domain` rows leave the
  coverage debt.
- **Holdout.** The holdout summary passes all three contracts. The worst
  holdout errors are 2.5E-16, 3.4E-16 and 3.3E-16.

## Not established

- **`|t| > 8` at df 1E5..1E8** stays on the continued fraction. It is a
  characterized limitation (`numerical_limitations.csv`).
- **Tails above 0.4** (`|p − 0.5| < 0.1`) are not polished and are outside the
  inverse contract.
- **The 44 holdout rows of the first wave** are no longer untouched. They are
  checked here as PI and PU rows, not as holdout evidence.
