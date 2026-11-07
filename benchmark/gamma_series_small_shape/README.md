# Gamma-series prefactor study (#23)

*Status, 2026-11-07.* The selected source change below shipped in `8c1506f`
(2026-08-27). The tables and the root-cause decomposition describe the code
before that commit. Closure evidence is still pending and #23 is open. One
piece of off-grid evidence has since arrived: the `holdout_older` re-export at
`7fb08ae` (2026-11-06) measures its Gamma rows at shapes 2.7 and 4.3 at
1.03E-15 or less, and they pass their contracts
(`../holdout_older/holdout_older_summary.md`).

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

## The Excel wave (2026-11-07)

The study grid holds exactly the frozen inputs of `HOLDOUT_DESIGN.md`:

| Regime | Rows | Content |
| --- | ---: | --- |
| `holdout` | 44 | shapes `2E-6 ... 2E-1` (11) at standardized arguments `0.125` and `0.75`, on the Gamma route `Gamma_Cumulative(z, s, 1)` and the Chi-square route `ChiSquare_Cumulative(2z, 2s)` |
| `scope_probe` | 3 | `Gamma_Cumulative` at `(0.5, 0.6, 1)`, `(1.0, 5.3, 1)`, `(10.8, 10.3, 1)` |
| `counterexample` | 2 | `Gamma_Cumulative(0.5, 1E-4, 1)` and `ChiSquare_Cumulative(1, 2E-4)` |

There are no seam rows: the design asks for them only if the fix uses a shape
boundary, and the selected fix does not.

Each reference is `P(a, x)` at the binary64 inputs, to 60 digits. The design's
own 25-digit scope references are `P` at the decimal inputs; at `(1.0, 5.3)`
they differ from the binary64 value by 3.2E-16, a tenth of the threshold, so
the grid does not use them. `generate_gamma_series_small_shape.py --check`
recomputes both readings. The exporter writes back the binary64 inputs it
parsed (`echo_arg1`, `echo_arg2`), and the analyzer scores a row only when they
equal the reference's inputs.

The wave exports this grid and the fitting ladder
(`../positive_ratio_subnormal/gamma_q_complement.csv`) on the self-hosted
runner, `export_studies=gss_small_shape,prs_gamma_q_complement` (#225).
`analyze_gamma_series_small_shape.py` scores the grid at `3E-15` relative
(Part A) and checks the ladder against its pre-fix export at `5c884bc`, rule by
rule of `MOVEMENT_MANIFEST.md` (Part B). Its results are in
`wave_results.json`.

## Files

| File | Purpose |
| --- | --- |
| `HOLDOUT_DESIGN.md` | frozen holdout and scope-probe design |
| `MOVEMENT_MANIFEST.md` | frozen movement predictions, and the post-selection explanation |
| `generate_gamma_series_small_shape.py` | builds the grid from the design; `--check` verifies it |
| `gamma_series_small_shape_grid.csv` | the study grid (13 columns; the last three are written by Excel) |
| `M_STATS_PROBDIST_GSSS.bas` | `Export_GammaSeriesSmallShape`, the exporter |
| `analyze_gamma_series_small_shape.py` | scores the wave; `--write` records `wave_results.json`, `--check` reproduces it |
| `test_gamma_series_small_shape.py` | fixture proofs of the generator and the analyzer |

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
