# #365: the Student-t tail below df = 1E5 — preregistration

Registered 2026-11-25 (Europe/Rome), before the candidate is run in Excel.
The 207 rows with regimes `below1e5_holdout` and `below1e5_seam` are
committed in `tchi_large_df/tchi_large_df_grid.csv` with references only. No
production source changes in the commit that carries this file.

This is the light registration of the evidence rules: a defect fix that
changes a route states its predictions and its rules before the Excel wave.
It reuses evidence that is already revealed, and says so.

## The defect, as measured

`below1e5_study.py` mapped `StudentT_Survival` on 25 df from 100 to the Double
below 1E5 and 13 t from 0.5 to 40: 314 rows, exported by Excel from the source
of today (run 45602906527, `below1e5_results.json`).

| | Rows |
| --- | ---: |
| Mapped | 314 |
| Above 1E-13 | 60 |
| Above 1.3E-12, the forward guarantee of the Student-t contracts | 7 |

The seven are at `t = 2` and `t = 3` for df from 42170 up to the seam, worst
**3.16E-12** at `t = 2` on the Double below 1E5. The first row above 1E-13
is at df = 562.

The 60 rows fall in three groups with three causes:

| Group | Rows above 1E-13 | Worst | Cause |
| --- | ---: | ---: | --- |
| `df >= 1999`, `1.74 <= t <= 16` | 36 | 3.16E-12 | the continued fraction at a rounded argument near one |
| `df < 1999` | 12 | 2.5E-13 | the beta factor through log-gamma, used while `a + b < 1000` |
| `t > 16` | 12 | 8.2E-13 | the tail is the exponential of a logarithm of size up to 700 |

Only the first group breaks the guarantee, and it is the mechanism of #213:
with `a = df / 2`, `b = 1/2` and `x = df / (df + t^2)`, the fraction's leading
denominator `1 - (a + b) x / (a + 1)` is a difference of nearly equal
numbers, and the rounding of `x` is multiplied by about `a / (1/2 + a y)`,
`y = 1 - x`.

The second group ends exactly where the Loader form of the factor begins
(`a + b >= 1000`, df = 1999): no row with `t <= 1.5` is above 1E-13 from
df = 2371 upward. It belongs to #175. The third is the limitation already
recorded for `df >= 1E5` (`IncompleteBeta.StudentTLargeDFFarTail`).

## The rule

In `PROB_TryStudentTTail`, for `df < 1E5` and `|t| <= 16`:

    y    = t^2 / (df + t^2)                  formed directly
    tail = PROB_TryBetaLowerNearOne(y, df / 2, 1/2) / 2

`PROB_TryBetaLowerNearOne` is the function of #213, unchanged. It serves
`a + b >= 1000` and an argument below the mean, which here is `df >= 1999`
and `t^2 > 3 df / (df + 2)`, so `|t|` above about 1.732. Anywhere else it
declines and the route of today answers.

The cap at 16 keeps the far tail on the route of today. Beyond it the model
puts the rule and that route at the same error, set by the final exponential,
so the rule would change observations there and gain nothing.

Nothing changes for `df >= 1E5`, for `df < 1999`, for `|t|` below the mean
bound or for `|t| > 16`. The inverse is not changed; it inherits the tail.

## Development evidence, on revealed data

`below1e5_model.py` models the rule in binary64 with the models of #213. The
model is not bit-faithful to Excel: 14 of the 314 map rows agree bit for bit,
and on the worst row it gives 6.97E-12 where Excel measures 3.16E-12. Its
figures bound the rule; they do not predict observations.

| Population | Route of today | Model of the rule |
| --- | ---: | ---: |
| The 98 map rows the rule serves, worst | 3.16E-12 (Excel) | 2.7E-14 |
| 600 further model points with `1.74 <= t <= 16`, df from 1999 to 1E5, worst | 7.0E-12 (model) | 3.4E-14 |
| The 144 holdout survival points, worst | 7.8E-12 (model) | 3.0E-14 |
| Survival across df = 1E5 at `t = 2, 4, 8`, worst rise | 3.16E-12 (Excel) | 2.8E-15 |

The model was run on the holdout points before this registration, so they
are not untouched for the model. They are untouched for Excel: no source has
evaluated them there.

## The rows

`below1e5_study.py --append-holdout` adds 207 rows:

| Rows | Regime | What they hold |
| ---: | --- | --- |
| 144 | `below1e5_holdout` | `StudentT_Survival(t, df)` at 12 df from 2048 to 99999 and 12 t from 1.74 to 16, all served by the rule (one t replaced by Amendment 1) |
| 36 | `below1e5_holdout` | `StudentT_Cumulative(-t, df)` at t in {2.25, 6.2, 13.6}, same df |
| 12 | `below1e5_holdout` | `StudentT_InverseCumulative(p, df)` at the correctly rounded lower-tail probability of t in {2.25, 9.1}, six df |
| 9 | `below1e5_seam` | the survival at the Doubles on either side of `t = 16`, at df in {2048, 9999, 50000} |
| 6 | `below1e5_seam` | the survival at the Doubles on either side of df = 1999, at t in {4, 8} |

References are 80-digit values from the oracle of the seam study.
`below1e5_before.json` keeps the observations of the 314 map rows as Excel
exported them from the source of today.

## Acceptance rules

`below1e5_study.py --check` implements them on the exported grid.

| Id | Rule |
| --- | --- |
| G1 | Every map row and every holdout forward row that the rule serves is within **1E-13** of the reference |
| G2 | Every map row that the rule does not serve holds the observation of `below1e5_before.json`, bit for bit |
| G3 | Across df = 1E5, at `t = 2, 4, 8` and both signs (the seam rows of #210), the survival and the cumulative move against the required direction by at most **1E-13** relative |
| G4 | Every holdout inverse is within **1E-12** of the reference |
| G5 | Every row of the three regimes holds a number |

Also required in the same wave: the full Excel regression, and the strict
accuracy gate on the main grid and the consolidated holdout with no contract
changed.

G3 is a bound, not strict monotonicity. One unit in the last place of df
moves the true tail by about 1E-16 relative, less than the rounding of either
route, so no rule at this accuracy can promise the sign of that step.

Reported, not gating: the steps across `|t| = 16` and across df = 1999, where
the rule meets the route of today and inherits that route's error; and the
cost per call, measured in Excel with `benchmark/kernel_profile` before and
after.

## Predictions

1. G1 to G5 hold.
2. The worst forward error on the rows the rule serves is below 5E-14.
3. The worst step across df = 1E5 at `t = 2, 4, 8` is below 2E-14; at
   `t = 0.5` and `t = 1` the seam rows do not change.
4. Every holdout inverse is within 5E-14.
5. No observation with `df >= 1E5`, `df < 1999`, `|t| > 16` or `|t| <= 1.73`
   changes in any grid.
6. Rows above 1E-13 remain for `df < 1999` and for `t > 16`, as mapped, all
   within the 1.3E-12 guarantee.

## Rejection policy

A failure of G1 to G5, of the regression or of the gate rejects the rule.
The bounds are not retuned and the result is recorded.

## What this does not address

- `df < 1999`: the factor through log-gamma, up to 2.5E-13. #175.
- `t > 16`: the final exponential, up to 8.2E-13 in the map. The limitation
  `IncompleteBeta.StudentTLargeDFFarTail` is extended to df below 1E5.
- Both stay within the 1.3E-12 guarantee and above the 1E-13 that holds
  where the rule or the expansion answers.

## Amendment 1 — 2026-11-25, after the first wave

Disclosed, and not a retuning: no bound, rule or prediction changes.

One of the twelve holdout values of t, the tenth, equalled a shape that the
routing holdout of #307 reserves: `gamma_route_study/check_disjoint.py`, a
standing check of the evidence suite, requires that number to occur nowhere
else in the repository. The registration commit did not run that suite, so
the clash was found only when the first wave (run 45618391642, at `23095a0`)
was applied.

- The 24 rows at that value (12 survival, 12 cumulative) are replaced by the
  same rows at `t = 13.6`. This file and `below1e5_study.py` no longer write
  the old value.
- The first wave had already exported the 24 replaced rows: every one was
  within G1, worst 1.87E-14. They are not part of the evidence any more, and
  that figure is recorded here so that their removal hides nothing.
- The other 497 rows of the three regimes keep the observations of the first
  wave, which accepted G1 to G5 on the registered rows. The 24 new rows are
  exported by a second wave of the same source and judged by the same rules.
- The model was not run on the new points before they were written.
