# Gamma-series prefactor study (#23)

This study isolates the accuracy floor reported in issue #23 before changing
production source. The frozen release contracts remain constraints:

- `Gamma_Cumulative.all.output`: relative error at most `3E-15`;
- `Gamma_Survival.all.output`: relative error at most `5E-15`;
- no threshold is widened from evidence observed during this remediation.

## Current measured evidence

The real-VBA fitting ladder is already committed in
`benchmark/positive_ratio_subnormal/gamma_q_complement.csv`. At `X = 0.5`,
`ScaleParam = 1`, the raw series-P result and the public Gamma CDF breach the
CDF contract for every measured shape below `0.5`:

| Shape | CDF relative error |
| ---: | ---: |
| 1E-06 | 7.97E-15 |
| 3E-06 | 1.05E-14 |
| 1E-05 | 1.14E-14 |
| 3E-05 | 1.28E-14 |
| 1E-04 | 9.43E-15 |
| 3E-04 | 1.03E-14 |
| 1E-03 | 1.05E-14 |
| 3E-03 | 9.38E-15 |
| 1E-02 | 9.32E-15 |
| 3E-02 | 9.13E-15 |
| 1E-01 | 1.07E-14 |
| 5E-01 | 9.58E-17 |

References are evaluated at the actual binary64 inputs. The oracle is
cross-validated against MPFR 4.2.1 / Rmpfr 0.9.5 in
`benchmark/positive_ratio_subnormal/oracle_cross_validation.md`.

## Root-cause decomposition

The current routine forms:

```text
P(A,X) = SeriesSum(A,X) * GammaPrefactor(A,X)
```

Decomposing the committed real-VBA output against 110-digit references shows:

- the ascending series sum contributes approximately `2E-17` to `3E-16`
  relative error across the fitting ladder;
- the inferred prefactor contributes approximately `8E-15` to `1.3E-14`;
- the observed CDF error follows the prefactor almost one-for-one.

The exact source path is:

```text
PROB_TryGammaSeriesP
  -> PROB_TryGammaPrefactor
     -> PROB_TryGammaLogPdf
        -> PROB_StirlingError(Shape)
```

For a small off-grid shape, `PROB_StirlingError` recurs to `Shape + 1` and
uses the log-gamma identity there. Its own absolute `1E-13` contract is met,
but an approximately `1E-14` absolute error in a log-prefactor becomes an
approximately `1E-14` relative error in `P`. The defect is therefore a
composition failure, not a failure of series convergence and not a breach of
the standalone Stirling-error contract.

## Initial candidate and source-scope rejection

The first candidate was a small-shape normalized series:

```text
NormalizedSum = 1 + X/(A+1) + X^2/((A+1)(A+2)) + ...

LogP = A*Log(X) - X - LogGamma1p(A) + Log(NormalizedSum)
P    = Exp(LogP)
```

This avoids both `1 / A` and the small-shape Loader/Stirling prefactor. Its
binary64 mirror predicts relative error below `5E-16` at every existing
`X = 0.5`, `Shape <= 0.1` fitting point.

It was not selected for production because the preregistered source-scope probe
predicts the same prefactor floor at ordinary off-grid shapes above `0.5`.
Dispatching only on small shape would mask the demonstrated ladder while
leaving the root composition error in place.

## Selected source change

The production change removes the inaccurate off-grid source at its origin.
`PROB_StirlingError` keeps its exact half-integer table and its existing
asymptotic branch, but an off-grid `N <= 15` now recurs upward:

```text
delta(N) = delta(N + 1)
           + (N + 0.5) * Log1p(1 / N)
           - 1
```

The recursion reaches `N > 15` in at most sixteen calls, where the existing
asymptotic series supplies the anchor. For `N < 0.5`, the first step retains
the overflow-safe `Log1p(N) - Log(N)` form; subsequent reciprocals are bounded
by two. Exact half-integer behavior and the large-shape Loader regime are
unchanged.

This is smaller than introducing a new Gamma-series dispatch and fixes the
composition that produced the error. The binary64 source mirror predicts the
three preregistered off-grid probes move from roughly `9E-15`–`4E-14` to below
`2E-15`. That prediction is not closure evidence; the post-change values must
be exported from real VBA.

## Scope probe

The same source mirror predicts that ordinary non-half-integer shapes above
`0.5` may also inherit the prefactor floor, while exact half-integers use the
stored Stirling table and remain clean. These points are preregistered in
`HOLDOUT_DESIGN.md`. They do not widen #23 until real Excel confirms them.

## Closure evidence still required

- real-Excel pre-fix results for the preregistered scope probe;
- a source-change movement comparison against `MOVEMENT_MANIFEST.md`;
- real-Excel post-fix fitting and independent holdout results;
- permanent VBA regressions on the public Gamma and equivalent Chi-square
  routes;
- promotion of representative rows into the main grid under the existing
  contracts;
- strict Accuracy Gate and full Excel/VBA regression with no unrelated
  regression.
