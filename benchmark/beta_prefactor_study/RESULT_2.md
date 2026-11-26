# #175: the incomplete-beta prefactor — result of the second wave, 2026-11-26

Registration `PREREGISTRATION_2.md` (`2eeb392`). The verdicts of before were
frozen from Excel run 45785388153 (`76350d8`, `17bd24e`). The candidate is
`8721b6c`, whose source is that of `dec7696` byte for byte. Excel run
45789721558 (Excel 16.0 build 20719, 64-bit); the regression passed 3,067 of
3,067 assertions. The export was judged on the run's artifact before it was
applied, then applied as `751bcbb`.

**Verdict: ACCEPTED.** K1 to K4 hold, and the strict accuracy gate passes on
the main grid and the consolidated holdout with no contract changed.

    python prefactor_study2.py --check

## The rules

| Id | Rule | Outcome |
| --- | --- | --- |
| K1 | Every `ib` point within 2E-13 | held: worst 4.4E-14 (1.03E-12 before) |
| K2 | No newly inferior holdout point above 3E-14; at most 2 percent newly inferior; the inferior points at least halve | held: three new points (0.9 percent), the largest 2.24E-14; 26 inferior against 165 |
| K3 | Every row holds an observation, no library error | held |
| K4 | The twelve points of the issue within 3E-14; fewer inferior points in the floor grid, none new above 3E-14 | held: worst 5.1E-15; 61 against 123, none new |

The three newly inferior points, all below the loss bound:

| Point | Before | After | Excel |
| --- | ---: | ---: | ---: |
| `Poisson_PMF(168, 120)` | 3.8E-15 | 1.26E-14 | 3.7E-15 |
| `Beta_Cumulative(0.993091, 880, 5)` | 6.8E-15 | 1.80E-14 | 3.6E-16 |
| `Gamma_Density(375, 150, 1)` | 6.0E-15 | 2.24E-14 | 3.6E-15 |

## The predictions

| | Prediction | Outcome |
| --- | --- | --- |
| 1 | K1 to K4 hold | held |
| 2 | Worst `ib` below 1.5E-13; above 5E-13 before | held: 4.4E-14; 1.03E-12 |
| 3 | At most a third as many inferior holdout points | held: 26 of 165 |
| 4 | The twelve points within 1E-14 | held: 5.1E-15 |
| 5 | The worst `dv` error falls; no `dv` point rises by more than 5E-14 | held: 5.1E-12 to 4.4E-13; largest rise 1.65E-14 |

## What the wave measured

- **The first holdout**, exported again in this wave, reads as in the first
  wave: 38 inferior points against 173, worst `ib` 9.93E-14.
  `prefactor_study.py --check` still reports the two points of H2: that
  registration's verdict stands as recorded.
- **The floor grid of #493**: 62 dispositions are no longer needed and are
  removed from `floor_dispositions.json`, the twelve of this issue and 50 of
  those recorded against #494. 61 inferior points remain, all disposed: 36
  against #494, 17 against #495, one against #496, and the seven of the fast
  normal inverse, excluded by decision.
- **Other grids**: 191 observations of the main grid change, 65 of the
  consolidated holdout and 26 of `holdout_older`; the accuracy gates pass
  with no contract changed.
- **#163**: `Beta_Survival(0.82, 2.7, 3.4)`, 5.45E-15 until now against a
  5E-15 contract, measures 1.8E-16.
- **The model**: on the rows exported after the rule,
  `beta_route_model.py` returns Excel's Double at all 246 it evaluates.

## What is still inferior to Excel

26 points of the second holdout. Twenty-one are below 1E-13: the incomplete
beta at most 4.4E-14 (`StudentT_Survival(2.4, 1700)`), three `ib` points
above 3E-14 in 275.

Five are above 1E-13, all densities and mass functions far from the mean at
large arguments, and all inferior before the rule as well:

| Point | Library | Excel |
| --- | ---: | ---: |
| `Poisson_PMF(5600, 4000)` | 4.4E-13 | 5.1E-14 |
| `Gamma_Density(7500, 5000, 1)` | 3.0E-13 | 8.6E-14 |
| `Gamma_Density(3500, 5000, 1)` | 1.9E-13 | 5.9E-14 |
| `Binomial_PMF(4076, 8000, 0.4)` | 1.9E-13 | 3.3E-14 |
| `Binomial_PMF(4466, 300000, 0.02)` | 1.1E-13 | 5.4E-14 |

This is the residual the first registration named: one product of size
`X Log(X / Mean)`, rounded once. The rule lowered it from 5.1E-12 and does
not remove it. Under the floor rule of `CONTRIBUTING.md` these points are
open, as #498; they are not on the floor grid of #493, so no disposition
covers them.

## Departures from the registration

- "Newly inferior points below 3E-14 are recorded against #494 in
  `floor_dispositions.json`": that file disposes of points of the floor grid
  only, and the three points are holdout points. They are recorded in the
  table above instead.

## Evidence that was not exported again

Some committed observations are still those of the source of before: the
study grids without an unattended export (`beta_f_unbalanced`,
`cdf_large_shape`, `density_large_shape`, `excel_comparison`) and the 356
rows of the main grid that the main exporter does not fill. Of these,
`measure_balanced_survival.py` finds 141 Beta rows the rule can change; the
model, which reproduces every fresh row, differs from 23 of them, by up to
517 ULPs. They are counted apart under `fidelity.predates_source` in
`../beta_balanced_survival/results.json`. #499 asks for their export.

## Not done here

- The balanced Beta contracts are not tightened and the header of
  `PROB_TryBetaRegularized` is not corrected yet: separate commits, as
  registered.
- `|t|` above 16 in the Student t tail (#365), the incomplete beta at very
  large shapes (#495), and shapes below 1/2 are outside this rule.
