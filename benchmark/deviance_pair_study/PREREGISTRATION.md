# #498: the Loader deviance as a pair of Doubles — preregistration

Registered 2026-11-26 (Europe/Rome), before the candidate exists in VBA. The
180 holdout points are committed in
`../excel_function_map/deviance_pair_holdout_grid.csv` with references only.
No file under `src/` changes in the commit that carries this file.

## The defect, as measured

The second holdout of #175 (Excel run 45789721558,
`../beta_prefactor_study/RESULT_2.md`) leaves five points where the library
is less accurate than Excel by more than 1E-13 relative:

| Call | Library | Excel |
| --- | ---: | ---: |
| `Poisson_PMF(5600, 4000)` | 4.4E-13 | 5.1E-14 |
| `Gamma_Density(7500, 5000, 1)` | 3.0E-13 | 8.6E-14 |
| `Gamma_Density(3500, 5000, 1)` | 1.9E-13 | 5.9E-14 |
| `Binomial_PMF(4076, 8000, 0.4)` | 1.9E-13 | 3.3E-14 |
| `Binomial_PMF(4466, 300000, 0.02)` | 1.1E-13 | 5.4E-14 |

## The mechanism

Each of these functions returns `Exp` of a logarithm whose large term is the
Loader deviance `bd0(X, Mean) = X Log(X / Mean) + Mean - X`. The error of the
deviance, as an absolute number, is the relative error of the result. Three
things put it there.

1. **Cancellation between the mean and about three times the mean.** At
   `Poisson_PMF(5600, 4000)` the deviance is 284, formed as 1884 - 1600. The
   logarithm is rounded at 1884.
2. **The size of one Double.** A deviance of size D, even correctly rounded,
   is off by up to `D * 2^-53`: 6E-14 at D = 540. The logarithm of the result
   then adds its own terms to a sum of that size.
3. **In the Binomial, the two means.** `N p` and `N (1 - p)` are rounded
   before the deviance sees them. An error `e` in a mean moves the deviance
   by `(1 - X / Mean) e`: 5E-14 at `Binomial_PMF(4076, 8000, 0.4)`.

Extending the series that `PROB_TryDeviancePart` already uses near the mean
removes the first only. On 20000 random pairs with a deviance below 700 it
leaves a worst of 3.2E-13 against 1.1E-12 today, and 121 pairs above 1E-13.
It is not the rule.

## The rule

The deviance is formed as a pair of Doubles (high, low), and the pair is
carried to the final exponential.

**R1. The deviance as a pair.** A new private procedure of
`M_STATS_PROBDIST_SPECIALFUNCS.bas` returns `bd0(X, Mean)` as a pair, for a
Double `X` and a mean that is itself a pair:

    k        = the integer nearest Log2(X / Mean)
    s        = X / 2^k                                  (exact)
    w        = (s - Mean) / (s + Mean)                  (in pairs)
    Log(X / Mean) = k Log 2 + 2 (w + w^3/3 + w^5/5 + ...)   (in pairs)
    bd0      = X * Log(X / Mean) + (Mean - X)           (in pairs)

`Log 2` is a pair of constants. The series stops when a term is at most
1E-34 of the sum, which also covers `X = Mean`, where `w` is zero. The pair
arithmetic is that of #35 (`PROB_BP_Sum`, `PROB_BP_Add`, `PROB_BP_Mul`,
`PROB_BP_Div`): every operation is stored in a Double, because VBA keeps
64-bit intermediates inside an expression and the error-free transforms do
not survive them.

One form serves every distance from the mean: there is no threshold between
a series and a logarithm.

**R2. The mass functions and densities.** Where they are computed for the
caller as a value (not as a logarithm), these four take the pair:

| Function | Means, as pairs | Other terms |
| --- | --- | --- |
| `K_STATS_Poisson_PMF` | `Mean` | as today |
| `K_STATS_Binomial_PMF` | `N * p` exactly, and `N * (1 - p)` with `1 - p` exact | as today |
| `K_STATS_Gamma_Density` | `X / ScaleParam` | as today |
| `K_STATS_ChiSquare_Density` | `X / 2` | as today |

The other terms of the logarithm (the Stirling errors, the half-logarithms,
`Log(ScaleParam)`) are the Doubles of today. They are added to the pair, and
the result is `Exp(high) * (1 + low)`.

**R3. Where the rule does not apply.** With `X` or the mean outside 1E-100
to 1E+100, on the positive-ratio subnormal branch of the Gamma and chi-square
densities, and on every edge the functions answer before they reach the
deviance (`K = 0`, `K = N`, a zero mean, `X = 0`), the route is that of
today.

Nothing else changes. `PROB_TryDeviancePart` keeps its code, so every other
caller keeps its results: the incomplete beta and its inverses, the
cumulative and survival functions of the Binomial, Poisson and Negative
Binomial, the Beta, F and Student t densities, the Negative Binomial,
Geometric and Hypergeometric mass functions, the Gamma kernel, and the four
`LogPMF` and log-density functions, which return a logarithm in one Double.

## Development evidence

`pair_model.py` evaluates the rule in plain binary64. Without the rule each
function is the port of the source in
`../beta_balanced_survival/beta_route_model.py`; at the three Poisson and
Binomial points above it returns the errors Excel measured (4.4E-13,
1.9E-13, 1.1E-13). At the two Gamma points it gives 2.4E-13 and 2.5E-13
where Excel measured 3.0E-13 and 1.9E-13: it is not faithful there. Its
figures bound the rule; they do not predict observations.

The deviance pair alone, on 400 pairs at every distance from the mean: the
sum of the pair is within 1E-24 relative of the 60-digit value.

The functions, on random arguments with a result above 1E-290:

| | Points | Today, worst | Today, above 1E-13 | Rule, worst | Rule, above 1E-14 |
| --- | ---: | ---: | ---: | ---: | ---: |
| `Poisson_PMF` | 1283 | 7.1E-13 | 61 | 6.4E-16 | 0 |
| `Binomial_PMF` | 1463 | 1.2E-12 | 283 | 2.2E-15 | 0 |
| Gamma density, four scales | 1237 | 1.3E-12 | 91 | 2.2E-15 | 0 |

The five points of the issue under the rule: at most 7.4E-16.

The model was run on the holdout points before this registration
(`pair_study.py --model`): under the rule the worst is 1.8E-15, at
`Binomial_PMF(10222, 20000, 0.45)`; today 29 of the 180 are above 1E-13, the
worst 1.5E-12. Writing the holdout found one defect of the model, a series
that did not stop at `X = Mean`; it is fixed and tested. The points are not
untouched for the model. They are untouched for Excel: no source has
evaluated them there.

## The rows

`pair_study.py --build` writes 180 points, 180 rows, in the format of the
Excel floor grid: the library call and Excel's function for the same
quantity, so that Excel evaluates both sides.

For each parameter set the points are chosen by how far the logarithm of the
function has fallen from its centre: by 0.5, 20, 150, 400 and 600, on each
side where the function reaches it. The deviance the rule carries is about
that large. Points whose value is below 1E-290 are left out.

| Points | Class | What they hold |
| ---: | --- | --- |
| 51 | `pm` | `Poisson_PMF` at six means from 7 to 250000, two of them not integers |
| 46 | `bn` | `Binomial_PMF` at five `(N, p)`, N from 600 to 900000, p from 0.05 to 0.93 |
| 56 | `gd` | `Gamma_Density` at six `(Shape, Scale)`, shapes from 9 to 12000, scales 1, 3, 0.7 and 0.01 |
| 27 | `cs` | `ChiSquare_Density` at df 14, 900 and 30000 |

No point is in either holdout of #175. References are 50-digit values from
the definitions, at the binary64 arguments.

Before the candidate is written, the grid is exported once from the source of
today. `pair_study.py --freeze-before` keeps those verdicts, and those of the
floor grid of #493 and of the two holdouts of #175 as committed, in
`pair_before.json`.

## Acceptance rules

`pair_study.py --check` implements them on the exported grids.

| Id | Rule |
| --- | --- |
| J1 | Every holdout point is within **1E-14** of the reference |
| J2 | No holdout point is inferior to Excel; every row holds an observation and no library value is an error |
| J3 | Each of the five points of the issue, rows of the second holdout of #175, is within **1E-14** |
| J4 | In the floor grid of #493 and in each holdout of #175, no point is newly inferior to Excel with a library error above **3E-14**, and no more points are inferior than before |

1E-14 is the floor of `CONTRIBUTING.md`: at or below it a point is not
inferior to Excel whatever Excel returns, so J1 implies the first clause of
J2. 3E-14 is the loss bound of the second registration of #175.

Also required in the same wave: the full Excel regression, and the strict
accuracy gate on the main grid and the consolidated holdout with no contract
changed. Observations of the four functions are expected to change in the
main grid and the holdouts; each is read against its contract by the gate.

Reported, not gating: each class before and after; the cost per call, if it
is measured. Under the floor rule a route that is inferior to Excel is
corrected whatever it costs, so no cost bound is registered.

## Predictions

1. J1 to J4 hold.
2. The worst holdout error is below 5E-15; before the rule it is above
   3E-13.
3. Before the rule at least 20 of the 180 holdout points are inferior to
   Excel; after it, none.
4. Each of the five points of the issue is within 3E-15.
5. In the floor grid and in the two holdouts of #175, fewer points are
   inferior than before, and none is newly inferior.

## After acceptance

Dispositions of floor-grid points that stop being inferior are removed from
`floor_dispositions.json`. `numerical_limitations.csv` and the changelog
record what was measured. Whether the same pair should serve the Beta, F and
Student t densities, the Hypergeometric mass function (#496) and the
incomplete-beta prefactor is decided on the measurements, as separate work.

## Rejection policy

A failure of J1 to J4, of the regression or of the gate rejects the rule. The
bounds are not retuned and the result is recorded.

## What this does not address

- The Beta, F and Student t densities. The second holdout of #175 has none of
  them above 1E-13; they stay recorded against #494.
- The `LogPMF` and log-density functions: a logarithm of size D in one Double
  carries `D * 2^-53` by construction.
- Results below 1E-290, and the subnormal branch of the Gamma density.
- The cumulative functions, which sum or integrate these masses by other
  routes.
