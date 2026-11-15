# Retained-row comparison — 2026-11-12

## Decision

Neither F-A nor F-B passes the frozen necessary retained-row conditions.
Stop before fitting-grid, holdout or shared-path invariance evaluation.
No production change is selected and #35 remains open.

The float-operation baseline passes the prescribed scaling bounds on all
8 affected rows and the 2^-48 bound on all 5 controls. References recomputed
at 50 and 80 dps agree to the required 40 relative digits. Excel iteration
counts are unavailable in the retained tables; no iteration-count agreement
or new Excel execution is claimed.

Both candidates pass 2/8 affected rows and 5/5 controls. For p=0.9,
df1=100, df2=1e10, error falls from 9.31e-9 to 4.63e-9, still far above
2^-48. For p=0.5, df1=1e8, df2=10, it falls to 1.78e-16. The orientation
change alone is therefore insufficient under this model. The remaining
mechanism is not isolated by this result.

## Reproduction

Run from the repository root:

```sh
python benchmark/f_inverse_large_df_study/fix_comparison/test_vba_order_model.py
python benchmark/f_inverse_large_df_study/fix_comparison/run_baseline.py
python benchmark/f_inverse_large_df_study/fix_comparison/run_retained_candidates.py
```

The candidate runner rejects a stale/failed baseline. JSON records retain
exact input hex, supplied complement, pair, result and iteration traces.
Source and model hashes are recorded in the baseline. The prescribed
binary64 model is isolated from the existing extended-intermediate model.
LogGammaDelta is ported without substituting high-precision production values.

## Limits and next step

The retained shapes are covered; the reused forward model still rejects
LogGamma arguments below 0.5. Do not use this implementation to claim coverage
of the deferred small-shape invariance set. Fitting and holdout were not
needed to reject these candidates and remain unevaluated. Frozen protocols,
old results, live grids and VBA are unchanged.

A follow-up diagnostic must separate forward residual, density, iteration
and pair effects on the now-revealed retained rows. That diagnostic is
exploratory; any new candidate needs a separately frozen prospective test.
This record is not component-substitution closure evidence.

Five focused tests pass. The local strict suite initially reported 14 passes
and four unavailable-tool failures (Ruff, mypy and two actionlint checks);
this is not a full green CI claim.
