# #213: the Student-t tail above |t| = 8 at large df — preregistration

Registered 2026-11-14 (Europe/Rome), before the candidate is run in Excel.
The 32 rows with regime `above8` are committed in
`tchi_large_df/tchi_large_df_grid.csv` with references only. No production
source changes in the commit that carries this file.

This is the light registration of the evidence rules: a defect fix that
changes a route states its predictions and its rules before the Excel wave.
It reuses evidence that is already revealed, and says so.

## The defect, as measured

`SEAM_STUDY.md` (#210) exported the Doubles on either side of `|t| = 8`,
where the large-df expansion hands over to the incomplete-beta continued
fraction. Against the 80-digit references, on the committed export:

| df | Survival at 8 + 1 ulp, relative error |
| ---: | ---: |
| 1E5 | 4.33E-13 |
| 1E6 | 4.0E-12 |
| 1E7 | 8.1E-11 |
| 1E8 | 7.78E-10 |

All four are outside the 1E-13 forward guarantee. At df = 1E7 the survival
rises across the seam, and the inverse at the lower-tail probabilities of
`t = -8 - 1 ulp, -8, -8 + 1 ulp` is off by 1.29E-12, outside the 1E-12
inverse guarantee.

## The mechanism

Not the one #213 first named. The prefactor is sound: the Loader form is
insensitive to the argument. The error is the conditioning of the continued
fraction in its argument. With `a = df / 2`, `b = 1/2` and
`x = df / (df + t^2)`, the fraction's leading denominator is

    1 - (a + b) x / (a + 1)  =  (1/2 + (a + 1/2) y) / (a + 1),   y = 1 - x

a difference of nearly equal numbers. The rounding of `x`, half a unit in
the last place, is multiplied by about `a / (1/2 + a y)`: 1.5E5 at df = 1E7
and `t = 8`. The logistic form that builds `x` and `y` from logarithms adds
a few units more.

## The rule

In `PROB_TryStudentTTail`, for `df >= 1E5` and `|t| > 8` (the expansion
keeps `|t| <= 8`):

    y    = t^2 / (df + t^2)                  formed directly
    tail = PROB_TryBetaLowerNearOne(y, df / 2, 1/2) / 2

`PROB_TryBetaLowerNearOne` (new, in SPECIALFUNCS) returns `I_(1 - y)(a, b)`
with the Loader prefactor of `PROB_TryBetaRegularized`, unchanged, and the
pair continued fraction of #35 at the argument `1 - y` carried as an exact
pair of Doubles. It serves `0 < y < 1/2` with the argument below the mean
and a positive prefactor; anywhere else it declines and the route of today
answers.

The inverse is not changed. Its loop and its polish below 8 stay as #34
left them; they inherit the corrected tail.

Nothing changes for `df < 1E5` or `|t| <= 8`.

## Development evidence, on revealed data

`above8_model.py` models the rule in binary64 with the models of the #35
study. On the Double above 8 and on the six probe values of `t`, at four df:

| Population | Model of the source today | Model of the rule |
| --- | ---: | ---: |
| `t <= 14`, worst relative error | 4.3E-10 | 1.2E-14 |
| `t = 20, 30`, worst | — | 9.7E-14 |
| Survival not rising across 8 | fails at df = 1E7 in Excel | holds at all four df |

A simulation of the inverse loop with the modelled tail returns the seam
quantiles within 1.2E-15.

The model's points are the rows of this registration, so those rows are not
untouched evidence for the model. They are untouched for Excel: the rule has
not been run there. The seam rows of #210 were exported from the source of
today and are re-exported by the wave.

## The rows

`above8_study.py --append` adds 32 rows, regime `above8`:

| Rows | What they hold |
| ---: | --- |
| 24 | `StudentT_Survival(t, df)` at t in {8.25, 9, 10, 14, 20, 30} and df in {1E5, 1E6, 1E7, 1E8} |
| 8 | `StudentT_InverseCumulative(p, df)` at the correctly rounded lower-tail probability of t in {8.25, 10}, same df |

References are 80-digit values, as for the seam rows.

## Acceptance rules

`above8_study.py --check` implements them on the exported grid.

| Id | Rule |
| --- | --- |
| F1 | The survival at 8 + 1 ulp and the cumulative at -(8 + 1 ulp), at each of the four df, and every probe with `t <= 14`, are within **1E-13** of the reference |
| F2 | Across `\|t\| = 8`, at each df and both signs, the survival does not rise and the cumulative does not fall |
| F3 | Every seam round trip with `df >= 1E5` and every `above8` inverse is within **1E-12** of the reference |
| F4 | Every `above8` row holds a number |

Also required in the same wave: the full Excel regression, and the strict
accuracy gate on the main grid and the consolidated holdout with no
contract changed.

Reported, not gating: the probes at `t = 20` and `t = 30`, whose values are
below 1E-88 and whose error the model puts near 1E-13 from the final
exponential; and how many inverses miss the 5E-15 measured target.

## Predictions

1. F1 to F4 hold.
2. The worst gated forward error is below 2E-14.
3. Every seam round trip at `df >= 1E5` is within 5E-15.
4. No observation with `|t| <= 8` or `df < 1E5` changes in any grid.
5. The seam's remaining monotonicity breaks are those across `df = 1E5`,
   unchanged; the limitation `StudentT.LargeDFSeamMonotonicity` keeps that
   part.

## Rejection policy

A failure of F1 to F4, of the regression or of the gate rejects the rule.
The bounds are not retuned and the result is recorded.

## What this does not address

The same conditioning exists below `df = 1E5`, where the continued fraction
serves every `t`: the seam study records up to 3.2E-12 just below 1E5 at
`t = 2`. That is outside #213 and outside this rule.
