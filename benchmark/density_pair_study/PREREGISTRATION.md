# #501: the Beta, F and Student t densities as pairs of Doubles — preregistration

Registered 2026-11-26 (Europe/Rome), before the candidate exists in VBA. The
114 holdout points are committed in
`../excel_function_map/density_pair_holdout_grid.csv` with references only.
No file under `src/` changes in the commit that carries this file.

## The defect, as measured

After the rule of #498 (Excel run 45822509492) ten points of the two holdouts
of #175 are less accurate than Excel, one of them by more than 1E-13
relative:

| Call | Library | Excel |
| --- | ---: | ---: |
| `F_Density(6, 300, 900)` | 1.3E-13 | 1.6E-14 |
| `F_Density(5, 500, 700)` | 7.6E-14 | 4.2E-15 |
| `StudentT_Density(20, 900)` | 4.6E-14 | 3.9E-14 |
| `F_Density(0.2, 300, 900)` | 4.3E-14 | 2.0E-14 |
| `F_Density(3, 300, 900)` | 2.1E-14 | 1.7E-14 |
| `StudentT_Density(16, 700)` | 1.7E-14 | 8.7E-16 |
| `StudentT_Density(9, 80)` | 1.7E-14 | 2.1E-15 |
| `Beta_Density(0.812111, 150, 90)` | 1.3E-14 | 2.5E-16 |
| `StudentT_Density(20, 80)` | 1.2E-14 | 1.1E-14 |
| `Beta_Density(0.362112, 480, 500)` | 1.1E-14 | 8.0E-15 |

## The mechanism

Three densities, three causes.

1. **Beta.** `PROB_TryBetaLogPdf` takes two Loader deviances,
   `bd0(Alpha, N X)` and `bd0(Beta, N Y)`, in one Double each: what #498
   found in the Poisson and Binomial mass functions. Here `N = Alpha + Beta`,
   `N X` and `N Y` are also rounded before the deviance sees them.
2. **F.** The same, and before it the arguments. `K_STATS_F_Density` forms
   `U = r / (1 + r)` and `V = 1 / (1 + r)`, `r = X d1 / d2`, through an
   exponential of a sum of three logarithms. A relative error `e` in `U`
   moves the deviance by `(N U - Alpha) e`: at `F_Density(6, 300, 900)`,
   250 times `e`.
3. **Student t.** No deviance. `PROB_TryStudentTPDF` forms the product
   `(df + 1) / 2 * Log(1 + t^2 / df)`, rounded at its own size, and a
   constant from a difference of log-gammas (`PROB_LogGammaHalfDiff`), off
   by up to 1E-14 at small degrees of freedom.

## The rule

Every large term is formed as a pair of Doubles and carried to
`Exp(high) * (1 + low)`, on the pair arithmetic and the pair deviance of
#498.

**R1. The logarithm of a ratio as a pair.** The part of
`PROB_TryDeviancePair` that forms `Log(X / Mean)` becomes a private
procedure of its own, which the deviance calls. The arithmetic of the
deviance does not change.

**R2. Beta density.** `1 - X` exactly, `N = Alpha + Beta` exactly, `N X` and
`N (1 - X)` as pairs; both deviances as pairs. The other terms of
`PROB_TryBetaLogPdf` (the half-logarithms, the three Stirling errors,
`Log(X)`, `Log(1 - X)`) are the Doubles of today, added to the pair.

**R3. F density.** `r = X d1 / d2`, `U = r / (1 + r)` and `V = 1 / (1 + r)`
as pairs, then R2 at shapes `d1 / 2` and `d2 / 2` without its two
logarithms, which cancel against those the F density adds, and `-Log(X)`.

**R4. Student t density.** Loader's form, which has no log-gamma:

    f(t) = Exp(-bd0(df/2, (df+1)/2) + d((df+1)/2) - d(df/2)
               + (df+1)/2 * Log(df / (df + t^2))) / Sqr(2 Pi)

with `d` the Stirling error, the deviance and the logarithm as pairs, and
`df + t^2` as a pair.

**R5. Where the rule does not apply.** With an argument outside 1E-100 to
1E+100, where a pair procedure returns FALSE, and on every edge the
functions answer before they reach these terms (`X = 0`, `X = 1`, the
validation, the shape envelope), the route is that of today.

**Not changed: `PROB_TryBetaLogPdf`.** It also forms the prefactor of the
incomplete beta (`PROB_TryBetaRegularized`, `PROB_TryBetaTailsPair`,
`PROB_TryBetaLowerNearOne`). The issue asked to decide, on measurements,
whether those callers take the pair. They do not, here: on the 275
incomplete-beta points of the second holdout of #175 the worst is 4.4E-14,
none above 1E-13, and the prefactor multiplies a continued fraction that
stops at 3E-14. The densities take new procedures; the log-density keeps its
code and every cumulative function keeps its results. The 34 points recorded
against #494 stay there.

The `LogDensity` functions, where they exist, return a logarithm in one
Double and are not changed.

## Development evidence

`density_model.py` evaluates the rule in plain binary64. Without the rule the
Beta density is the port of the source in
`../beta_balanced_survival/beta_route_model.py`, and returns the 1.26E-14
Excel measured at `Beta_Density(0.812111, 150, 90)`. The F and Student t
densities of today are ports written for this study; at
`F_Density(6, 300, 900)` the F port gives 2.1E-13 where Excel measured
1.3E-13. None is bit-faithful. The figures bound the rule; they do not
predict observations.

On random arguments with a density above 1E-290:

| | Points | Today, worst | Today, above 1E-13 | Rule, worst | Rule, above 1E-14 |
| --- | ---: | ---: | ---: | ---: | ---: |
| `Beta_Density` | 880 | 2.2E-13 | 21 | 1.7E-15 | 0 |
| `F_Density` | 1404 | 1.1E-12 | 130 | 1.2E-15 | 0 |
| `StudentT_Density` | 1327 | 1.4E-13 | 8 | 4.0E-16 | 0 |

For the Student t a smaller change was tried first, the logarithm as a pair
with the constant of today: 1.1E-14 worst, four points above 1E-14, all at
few degrees of freedom. Loader's form removes the constant.

The model was run on the holdout points before this registration
(`density_study.py --model`): under the rule the worst is 9.1E-16; today 15
of the 114 are above 1E-13, the worst 1.7E-12, at an F point. The points are
not untouched for the model. They are untouched for Excel.

## The rows

`density_study.py --build` writes 114 points, 114 rows, in the format of the
Excel floor grid. For each parameter set the points are where the logarithm
of the density has fallen from its mode by 0.5, 20, 150, 400 and 600, on each
side where the density reaches it (one side for the Student t, which is
symmetric). Points whose value is below 1E-290 are left out.

| Points | Class | What they hold |
| ---: | --- | --- |
| 51 | `bd` | `Beta_Density` at six shape pairs, from (2.5, 640) to (3500, 8000) |
| 38 | `fd` | `F_Density` at five pairs of degrees of freedom, from (8, 30) to (3000, 5000) |
| 25 | `td` | `StudentT_Density` at df 3, 11, 64, 450, 5000 and 40000 |

No point is in a holdout of #175 or #498. References are 50-digit values
from the definitions, at the binary64 arguments.

Before the candidate is written, the grid is exported once from the source of
today. `density_study.py --freeze-before` keeps those verdicts, and those of
the floor grid of #493, of the two holdouts of #175 and of the holdout of
#498 as committed, in `density_before.json`.

## Acceptance rules

`density_study.py --check` implements them on the exported grids.

| Id | Rule |
| --- | --- |
| L1 | Every holdout point is within **1E-14** of the reference |
| L2 | No holdout point is inferior to Excel; every row holds an observation and no library value is an error |
| L3 | Each of the ten points of the issue, rows of the two holdouts of #175, is within **1E-14** |
| L4 | In the floor grid of #493, in each holdout of #175 and in the holdout of #498, no point is newly inferior to Excel with a library error above **3E-14**, and no more points are inferior than before |

Also required in the same wave: the full Excel regression, and the strict
accuracy gate on the main grid and the consolidated holdout with no contract
changed.

The wave exports the main grid, the consolidated holdout, `holdout_older`,
the floor grid, the two holdouts of #175, the holdout of #498 and this one.
It exports no frozen study grid: the wave of #498 did, in error, and the
files had to be restored.

Reported, not gating: each class before and after; the cost per call, if it
is measured.

## Predictions

1. L1 to L4 hold.
2. The worst holdout error is below 5E-15; before the rule it is above
   5E-13.
3. Before the rule at least 15 of the 114 holdout points are inferior to
   Excel; after it, none.
4. Each of the ten points of the issue is within 3E-15.
5. In the floor grid, the two holdouts of #175 and the holdout of #498, no
   point is newly inferior.

## After acceptance

Dispositions of floor-grid points that stop being inferior are removed.
The changelog records what was measured. No measured figure is written in a
source header.

## Rejection policy

A failure of L1 to L4, of the regression or of the gate rejects the rule. The
bounds are not retuned and the result is recorded. A failure of one density
does not reject the others if they can be separated without changing a rule;
that would be a new registration, not an amendment after the wave.

## What this does not address

- The prefactor of the incomplete beta, as said above.
- The `LogDensity` functions.
- Densities below 1E-290, and arguments outside 1E-100 to 1E+100.
- The Negative Binomial and Hypergeometric mass functions (#496), which go
  through the Binomial logarithm in one Double.
