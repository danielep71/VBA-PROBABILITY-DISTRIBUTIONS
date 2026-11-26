# #175: the incomplete-beta prefactor and the far deviance — preregistration

Registered 2026-11-26 (Europe/Rome), before the candidate exists in VBA. The
317 holdout points are committed in
`../excel_function_map/prefactor_holdout_grid.csv` with references only. No
production source changes in the commit that carries this file.

This is the light registration of the evidence rules: a defect fix that
changes a route states its predictions and its rules before the Excel wave.
It reuses evidence that is already revealed, and says so.

## The defect, as measured

The Excel floor comparison of #493 (run 45703479787, `floor_results.json`)
records twelve points against this issue: the library is less accurate than
Excel there, by more than 1E-13 relative.

| Call | Library | Excel |
| --- | ---: | ---: |
| `NegativeBinomial_Survival(118, 500, 0.85)` | 4.9E-13 | 3.2E-14 |
| `NegativeBinomial_Survival(354, 200, 0.4)` | 2.2E-13 | 9.4E-16 |
| `NegativeBinomial_Survival` and `Cumulative(88, 500, 0.85)` | 2.0E-13, 1.8E-13 | 1.5E-15 |
| `NegativeBinomial_Survival` and `Cumulative(200, 50, 0.2)` | 1.5E-13, 1.4E-13 | 1.7E-15 |
| `Beta_Survival(0.410058, 50, 100)` | 1.4E-13 | 3.8E-15 |
| `Beta_Cumulative` and `Survival(0.333333, 50, 100)` | 1.3E-13 | 3.7E-16 |
| `Beta_Cumulative` and `Survival(0.666667, 100, 50)` | 1.3E-13 | 3.7E-16 |
| `StudentT_Survival(1, 1000)` | 1.6E-13 | 6.5E-17 |

The comparison of #365 with `T.DIST.RT` adds Student t with df below 1999: 53
of 72 points inferior, up to 4.9E-13.

## The mechanism

All of the error is in the logarithm of the prefactor
`Bt = x^a y^b / B(a, b)`; the continued fraction is not involved: on the six worst points
of the Loader form a stop of `2^-60` in place of 3E-14 changes no figure, and
the error of `Log(Bt)` alone accounts for the whole of it. There are two
causes.

1. **Below a + b = 1000 the prefactor is the literal form**,
   `Exp(a Log x + b Log y - PROB_LogBeta(a, b))`. `PROB_LogBeta` comes from a
   Lanczos sum with 15-digit coefficients, and the three terms are of the size
   of `LogGamma(a + b)`. This is the cause the issue names. The model gives
   1.30E-13 at `I_x(50, 100)`, `x = 0.333333`, where Excel measures 1.3E-13.
2. **In the Loader form, away from the mean, the deviance subtracts two
   logarithms.** `PROB_TryDeviancePart` computes
   `X * (Log(X) - Log(Mean)) + Mean - X`. Each logarithm is rounded at its own
   size, so the difference loses a factor of about `Log(X) / Log(X / Mean)`.
   This is why the issue's own proposal, the Loader form at every shape, is
   not enough: in the model it repairs the twelve points and is worse than
   today in the far tails, up to 3.7E-13.

## The rule

Two changes in `M_STATS_PROBDIST_SPECIALFUNCS.bas`.

**R1. The Loader form at every shape.** `PROB_TryBetaRegularized` takes the
Loader prefactor (`PROB_TryBetaLogPdf`) whenever both shapes are at least
1/2, and from a + b = 1000 as today. With a shape below 1/2 and a + b below
1000 the literal form stays: the model does not port that branch, so the rule
does not touch it.

**R2. One logarithm in the far deviance.** In `PROB_TryDeviancePart`, where
`|X - Mean| >= 0.1 (X + Mean)`:

    Difference = X - Mean
    If Mean / 2 <= X <= 2 Mean:   Result = X * Log1p(Difference / Mean) - Difference
    Otherwise:                    Result = X * Log(X / Mean) - Difference

Within a factor of two the difference is exact, so `Log1p` receives a ratio
with one rounding. Outside it `1 + Difference / Mean` would cancel for
`X << Mean`, so the ratio `X / Mean` is formed directly. The series branch
near equality does not change.

`PROB_TryDeviancePart` is shared. R2 therefore reaches every caller: the
incomplete beta in the Loader form at any shape, the Beta, F and Student t
densities, the Gamma density, and the Binomial and Poisson PMF. That is
intended; the rules below judge all of them.

Nothing else changes: not the continued fractions, not `PROB_LogBeta`, not
`PROB_TryBetaLowerNearOne`, not the inverses, which inherit the forward
functions.

## Development evidence, on revealed data

`prefactor_model.py` evaluates the rule in two ports of the source: `x87`
(every intermediate in extended precision; it reproduces Excel on the
balanced route and does not port the unbalanced one) and `double`. Neither is
bit-faithful to Excel everywhere. Their figures bound the rule; they do not
predict observations.

Relative error of `I_x(a, b)` on 2235 points with a + b below 1000 and both
shapes from 1/2 to 950, at nine distances from the mean (the `double` port;
the `x87` port agrees on the 1109 points it models):

| | Worst | Above 1E-13 | Above 3E-14 |
| --- | ---: | ---: | ---: |
| Today | 1.3E-12 | 193 | 539 |
| R1 alone (the issue's proposal) | 3.7E-13 | 24 | 104 |
| R1 and R2 | 1.2E-13 | 1 | 15 |

The one point above 1E-13 under the rule is `I_x(950, 0.5)` at `x = 0.9972`,
1.2E-13 with or without the rule: an argument near one, the mechanism of
#365, not the prefactor.

The twelve points of the issue, under the rule: at most 6.1E-15.

The far deviance alone, as an absolute error (the relative error of the
density), on 3000 random pairs with a deviance below 700: today worst
1.9E-11 and 290 pairs above 1E-13; under R2 worst 8.0E-13 and 61 above
1E-13. What remains grows with `X`: a product of size `X Log(X / Mean)`
rounded once. R2 does not claim more than that.

The model was run on the holdout points before this registration
(`prefactor_study.py --model`): the rule's worst is 8.4E-14 (`x87`) and
5.6E-14 (`double`), at `StudentT_Survival(1.9, 1900)`, against 5.0E-13 and
4.3E-13 today. They are not untouched for the model. They are untouched for
Excel: no source has evaluated them there.

## The rows

`prefactor_study.py --build` writes 317 points, 422 rows, in the format of the
Excel floor grid: the library call and every Excel formulation of the same
quantity, so that Excel evaluates both sides.

| Points | Class | What they hold |
| ---: | --- | --- |
| 100 | `ib` | `Beta_Cumulative` and `Beta_Survival` at ten shape pairs with a + b below 1000, from -8 to +5 standard deviations |
| 40 | `ib` | `NegativeBinomial_Cumulative` and `Survival` at four `(s, p)` |
| 40 | `ib` | `Binomial_Cumulative` and `Survival` at four `(n, p)` with n below 1000 |
| 48 | `ib` | `F_Cumulative` and `F_Survival` at four pairs of degrees of freedom |
| 35 | `ib` | `StudentT_Survival` at df from 7 to 1900 and t from 0.4 to 15 |
| 54 | `dv` | `Binomial_PMF`, `Poisson_PMF`, `Gamma_Density`, `Beta_Density`, `F_Density`, `StudentT_Density`, far from the mean |

Every `ib` point reduces to an incomplete beta inside the region of R1.
References are 50-digit values from the definitions; an upper tail is taken
as the lower tail of the swapped shapes. Points whose value is below 1E-290
are left out.

Before the candidate is written, the grid is exported once from the source of
today. `prefactor_study.py --freeze-before` keeps those verdicts, and those of
the whole floor grid, in `prefactor_before.json`. That export also gives
Excel's side at these points for the first time.

## Acceptance rules

`prefactor_study.py --check` implements them on the exported grids.

| Id | Rule |
| --- | --- |
| H1 | Every `ib` point is within **2E-13** of the reference (the balanced Beta guarantee) |
| H2 | No holdout point is inferior to Excel that was not inferior before, and no more holdout points are inferior than before |
| H3 | Every row holds an observation, and no library value is an error |
| H4 | Each of the twelve points of the issue is within **3E-14** of the reference, and the floor grid of #493 has no inferior point that was not inferior before |

"Inferior" is the rule of `CONTRIBUTING.md`: Excel's relative error is
smaller and the library's is above 1E-14.

Also required in the same wave: the full Excel regression, and the strict
accuracy gate on the main grid and the consolidated holdout with no contract
changed. Observations of the main grid and of the holdouts are expected to
change, since R2 reaches many functions; each change is read against its
contract by the gate.

Reported, not gating: the `dv` points before and after; how many of the 86
points recorded against #494 stop being inferior.

## Predictions

1. H1 to H4 hold.
2. The worst `ib` error is below 1E-13; before the rule it is above 2E-13.
3. Each of the twelve points of the issue is within 1E-14.
4. Fewer holdout points are inferior to Excel than before, and fewer points
   of the floor grid.
5. No `dv` point gets worse by more than 1E-14, and their worst error does
   not rise.
6. With df below 1999 and `t` up to 16, the Student t survival is no longer
   inferior to `T.DIST.RT` at any compared point of #365; above 16 it still
   is, by the mechanism #365 keeps.

## After acceptance

The dispositions of the points that stop being inferior are removed from
`floor_dispositions.json` (the floor check fails on a disposition that is no
longer needed). The balanced Beta contracts are tightened to what Excel
measures in a separate commit, as the issue asks, and the header of
`PROB_TryBetaRegularized` is corrected.

## Rejection policy

A failure of H1 to H4, of the regression or of the gate rejects the rule. The
bounds are not retuned and the result is recorded.

## What this does not address

- Shapes below 1/2 with a + b below 1000: the literal form, unchanged.
- Arguments near one at large `a` and small `b` (#365), and the far Student t
  tail above `|t| = 16`.
- The incomplete beta at very large shapes near the mean (#495).
- The residual of the far deviance at large `X`, described above.

## Amendment 1 — 2026-11-26, before the candidate is run in Excel

Two points found while writing the rule in VBA. Neither changes a bound, a
rule or a prediction, and no export of the candidate existed when they were
made.

- **R1 also governs `PROB_TryBetaTailsPair`.** That procedure forms the same
  prefactor for the beta inverse, and its header says so. Leaving it on the
  old dispatch would make the inverse solve a function that differs from the
  forward call by the very error this rule removes. One private predicate,
  `PROB_IBetaUsesLoader`, now decides for both. The registration said the
  inverses only inherit the forward functions; with this they also change
  where their own tails are formed, for shapes in the region of R1. They are
  judged by their contracts in the gate and by the floor grid under H4.
- **R2 keeps the two logarithms at extreme magnitudes.** `X / MeanPart`
  overflows in VBA, as a run-time error, when the arguments are far enough
  apart. With either argument outside 1E-100 to 1E+100 the far branch is the
  expression of before. The model's points and the holdout are far inside
  that range.

`beta_balanced_survival/beta_route_model.py` ports the source as amended;
`prefactor_model.py` restores the dispatch and the far deviance of before for
the comparison, and reproduces the figures above unchanged.
