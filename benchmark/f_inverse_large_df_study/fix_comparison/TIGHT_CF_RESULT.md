# Tight-CF screen result — 2026-11-12

Protocol frozen at 7f562835588720283b68b9f4716e9f266b313c1e before evaluating
the eight new points. Result: SCREEN_PASS_NOT_ADOPTION.

P2 retains P1's pair-aware two-Double recurrence and changes only its relative
CF delta-stop from 3e-14 to 2^-100. The production inverse and prefactor models
are unchanged. This stop is a convergence condition, not an output-error
acceptance threshold: the required output bound remains 2^-48.

All 13 revealed fitting identities and all 8 prospective screen identities
pass. The worst new-row relative error is 2.03e-15; the worst revealed-row
error is 3.45e-15. The five controls also meet the four-times-F-0 condition.
All references agree at 50/80 dps to the required 40 relative digits.

| Population | p | df1 | df2 | F-0 error | P1 error | P2 error |
|---|---:|---:|---:|---:|---:|---:|
| revealed | 0.5 | 1000000.0 | 3.0 | 7.771e-11 | 2.915e-15 | 2.915e-15 |
| revealed | 0.5 | 100000000.0 | 10.0 | 9.490e-10 | 3.912e-15 | 1.780e-16 |
| revealed | 0.9 | 2.5 | 1000000.0 | 3.870e-11 | 5.752e-15 | 1.381e-15 |
| revealed | 0.95 | 2.5 | 1000000.0 | 1.581e-11 | 4.057e-15 | 1.231e-15 |
| revealed | 0.99 | 2.5 | 1000000.0 | 8.991e-12 | 7.476e-16 | 2.470e-15 |
| revealed | 0.9 | 5.0 | 100000.0 | 1.001e-12 | 1.343e-15 | 4.601e-16 |
| revealed | 0.9 | 100.0 | 10000000000.0 | 9.312e-09 | 1.823e-16 | 1.823e-16 |
| revealed | 0.99 | 100.0 | 10000000000.0 | 7.002e-09 | 3.250e-16 | 3.250e-16 |
| revealed | 0.9 | 1000000.0 | 3.0 | 1.832e-15 | 1.832e-15 | 1.832e-15 |
| revealed | 0.99 | 1000000.0 | 3.0 | 1.756e-15 | 1.148e-17 | 1.756e-15 |
| revealed | 0.5 | 2.5 | 1000000.0 | 3.445e-15 | 3.445e-15 | 3.445e-15 |
| revealed | 0.9 | 100000000.0 | 10.0 | 2.628e-15 | 1.045e-15 | 2.628e-15 |
| revealed | 0.5 | 100.0 | 10000000000.0 | 2.166e-15 | 2.166e-15 | 2.166e-15 |
| new | 0.2 | 20000.0 | 6.0 | 5.634e-14 | 4.335e-16 | 4.335e-16 |
| new | 0.2 | 6.0 | 20000.0 | 1.112e-16 | 1.112e-16 | 1.112e-16 |
| new | 0.2 | 40.0 | 20000000.0 | 1.339e-16 | 1.919e-15 | 1.339e-16 |
| new | 0.2 | 20000000.0 | 40.0 | 8.638e-12 | 3.520e-16 | 3.520e-16 |
| new | 0.8 | 20000.0 | 6.0 | 3.568e-16 | 1.347e-15 | 1.347e-15 |
| new | 0.8 | 6.0 | 20000.0 | 5.630e-14 | 4.237e-16 | 4.237e-16 |
| new | 0.8 | 40.0 | 20000000.0 | 8.638e-12 | 9.273e-17 | 9.273e-17 |
| new | 0.8 | 20000000.0 | 40.0 | 2.029e-15 | 2.029e-15 | 2.029e-15 |

## Why a separate protocol was needed

`pair_floor_diagnostic.json` is exploratory: it separates signed pair error
from final-transform error, substitutes a high-precision density-prefactor
value, and tightens the CF stop on already revealed rows. For the three P1
failures, the tighter CF stop removes the failures without changing Newton
or the final transform. The prefactor substitution alone does not. The public
log transform can still add or cancel a few ulps; these experiments do not
establish a universal conditioning limit or the best possible implementation.
The new eight-point screen was specified only after that diagnostic and
frozen before its evaluation. No result is retroactively preregistered.

## Reproduction and remaining gates

```sh
python benchmark/f_inverse_large_df_study/fix_comparison/diagnose_pair_floor.py
python benchmark/f_inverse_large_df_study/fix_comparison/run_tight_cf.py
python benchmark/f_inverse_large_df_study/fix_comparison/test_pair_cf_screen.py
```

Six pair-screen tests pass, including scoped stop restoration after failure
and uniqueness/disjointness of the new rows. Targeted Ruff passes.
Numerical JSON records keep exact binary64 inputs, complements, per-variant
traces and source bindings. Historical rejected P1 results remain unchanged.

No fitting-grid expansion, balanced-huge comparator, small-shape caller
invariance, CF/geometry seam wave or independent holdout was evaluated here.
No public VBA was changed. Before adoption, freeze that broader design and
complete those gates, including support for currently unmodelled small-shape
branches, followed by candidate-specific real-Excel evidence. #35 remains
open; this screen is not release certification or a merge recommendation.
