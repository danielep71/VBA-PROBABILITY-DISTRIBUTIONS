# Excel wave result — Student-t large-df route (#34)

The wave of `PRODUCTION_PREREGISTRATION.md` (§5, §8), as amended by
`PRODUCTION_AMENDMENT_1.md`. `check_excel_wave.py --baseline 5273b0c` checked
it against the predictions `production_predictions.json` recorded before any
VBA. Its output is `excel_wave_check.json`.

## The export

- **Run.** Excel VBA Regression run 42245883623, candidate `c8bb0ae`.
- **Regression.** 950/950, including all 19 assertions of
  `Test_TF_StudentTLargeDFTail`.
- **Grids.** The main grid, `tchi_large_df`, and the 44 untouched holdout rows
  of §6, each exported once.
- **Evidence commit.** `evidence: Excel export of main, holdout, tchi_large_df
  at c8bb0ae`.

## The predictions

| Prediction | Rows | Result |
| --- | ---: | --- |
| PT1 forward, within 4 ulps of the predicted Double and ≤ 2^-50 | 58 | **holds**: worst 1.93E-16, at most 1 ulp |
| PT2 inverse, ≤ (2^-50 + 2^-52) · min(p, 1−p) / (\|q\| f(q)) + 2^-51 | 18 | **falsified** at 7 rows |
| PT3 every other Student-t row and every F row bit-identical | 343 + 104 | **holds**: none changed |
| PT4 untouched holdout, PT1 and PT2 rules | 44 | **holds**: worst 7.14E-16 |

**PT2 failures.** These are `p = 0.99` at `df = 1E6, 1E7` and `p = 0.999` at
`df = 1E5, 1E6, 1E7`, in the main grid and in `tchi_large_df`. Their errors
are 6.9E-16 to 9.5E-16, against bounds of 5.5E-16 to 6.2E-16. They are 3 to 6
ulps off the exact quantile, with the route accurate to below 4E-16 at every
Double around it.

**Cause.** The cause is the inverse's existing iteration, not the route.
- The iteration is `PROB_TryStudentTInvTail`, a safeguarded Newton.
- Near the root, the Newton candidate falls outside the bracket, so the
  iteration takes bisection steps.
- It stops when a step is at most `PROB_EPS = 1E-15` relative, about 5 ulps at
  these quantiles. It does not converge to the nearest Double.
- A Python replica of the iteration on the route model stops 3 and 4 ulps short
  at two of the failing rows.
- PT2's last term, 2^-51, assumed the stopping rule ends within about 2 ulps.
  That assumption was never modeled.

## Consequences under §8

- **Contracts.** A falsified prediction stops adoption, so none of the §7
  contracts is adopted. `accuracy_contracts.csv` returns to its pre-#34
  contracts, and the nine Student-t inverse `envelope_domain` rows return to
  the coverage debt. The 5E-15 contracts committed before the export were
  never adopted.
- **The route stays.** Every PT3 row is unchanged. No PT1 row is worse than the
  current route's committed observation. The seven PT1 rows not strictly
  better are bit-identical to it, because there the continued fraction already
  returned the correctly rounded Double.
- **What it needs.** The route is a large accuracy repair: the old continued
  fraction erred by up to 4.2E-9 forward and 1.8E-10 on the quantile at these
  rows. It ships without a tightened contract until a new preregistration
  adopts one.

## Next

A new preregistration, with a fresh untouched holdout:
- fixes the inverse stopping rule to converge to the nearest Double;
- predicts the inverse from a model of that iteration;
- adopts the forward and inverse contracts in one more Excel wave.

The rows of §6 have been observed, so they are not used as untouched evidence
there.
