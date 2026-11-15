# Exploratory residual diagnosis — 2026-11-12

This follows rejection of F-A and F-B. All 13 identities were already revealed.
No fitting grid, invariance set or independent holdout was evaluated.
These substitutions are diagnostic oracles, not implementable VBA fixes.

## Observations

| p | df1 | df2 | F-A | HP CF, rounded argument | HP CF, exact complement |
|---:|---:|---:|---:|---:|---:|
| 0.5 | 1000000.0 | 3.0 | 2.915e-15 | 2.915e-15 | 2.915e-15 |
| 0.5 | 100000000.0 | 10.0 | 1.780e-16 | 1.780e-16 | 1.780e-16 |
| 0.9 | 2.5 | 1000000.0 | 2.155e-11 | 1.476e-12 | 1.381e-15 |
| 0.95 | 2.5 | 1000000.0 | 7.795e-12 | 9.989e-14 | 1.231e-15 |
| 0.99 | 2.5 | 1000000.0 | 6.888e-12 | 4.683e-13 | 2.470e-15 |
| 0.9 | 5.0 | 100000.0 | 3.540e-13 | 1.229e-13 | 4.601e-16 |
| 0.9 | 100.0 | 10000000000.0 | 4.627e-09 | 2.421e-11 | 1.823e-16 |
| 0.99 | 100.0 | 10000000000.0 | 3.049e-10 | 1.971e-10 | 3.250e-16 |

The high-precision CF substitution alone leaves substantial error on the six
large-df2 failures. Keeping that same high-precision CF but deriving its
reflected argument as 1-x at working precision removes those failures.
The latter experiment keeps the original binary64 Loader prefactor and
inverse density, bracket and stop rules. Its errors on all 13 rows remain
below 2^-48. Precision-50 and precision-80 experiments return identical
binary64 results and pairs for all seven variants on all 13 identities.

The final logarithm/exponential reconstruction contributes at most about
4.1e-15 relative error in these experiments; it does not explain the 1e-9
residual. Replacing only the prefactor is insufficient. Using a high-precision
prefactor with the supplied rounded complement even degrades clean controls,
so this is not evidence that the production Loader prefactor should be removed.

## Interpretation and limits

The evidence supports a second loss at the forward reflected-CF argument,
in addition to the original inverse pair unwind. CF arithmetic also matters;
the successful substitution changes both arithmetic precision and argument
representation relative to production. It does not prove that preserving the
complement alone in the existing binary64 recurrence will suffice.

Next: design a pair-aware CF evaluation that can retain the small member
without reconstructing it from a rounded large argument, and separately
assess recurrence rounding. Freeze its predictions and seams before evaluating
a new candidate. Preserve the rejected protocol and these revealed results.
No production correction or issue closure is justified by this diagnostic.

## Reproduce

```sh
python benchmark/f_inverse_large_df_study/fix_comparison/diagnose_residual.py
python benchmark/f_inverse_large_df_study/fix_comparison/test_residual_diagnostic.py
```

The JSON includes exact input identities, per-experiment pair/result/last
iteration, precision stability, transform error and source bindings. Two
diagnostic tests and targeted Ruff pass. No Excel validation was run.
