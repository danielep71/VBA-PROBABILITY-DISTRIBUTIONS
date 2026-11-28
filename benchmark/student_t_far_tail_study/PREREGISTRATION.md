# #365: the Student t tail beyond |t| = 16 as pairs of Doubles — registration, 2026-11-28

Written before any VBA of the rule. No file under `src/` changes in the
commit that holds this text.

## The defect

On the 140 comparison points of the issue (`below1e5_study.py --compare`, on
the observations committed today) the library is less accurate than
`T.DIST.RT` at 40, and at 10 by more than 1E-13. All 10 have `|t|` above 16
(t = 17, 24, 40) and df from 1000 to 99999; the worst is 2.8E-13 at
`df = 1999, t = 40`. The other 30 are below 1E-13 and stay recorded against
#494.

## The mechanism

For df below 1E5 and `|t|` above 16 the tail is half of `I_x(df/2, 1/2)` at
`x = df / (df + t^2)`, through `PROB_TryBetaRegularized`. Two things are
rounded there at the size of `t^2 / 2`:

- the logarithm of the prefactor `x^a y^b / B(a, b)` is one Double. It is
  near `-t^2 / 2`, between -130 and -690 where the tail is representable, and
  one rounding of it is up to 8E-14 of the result;
- `x` and `y` are one Double each. A relative rounding of `y` moves the
  logarithm of the tail by about `t^2 / 2` times as much.

The rule of this issue for `|t|` up to 16 (pull request 492) gave the
continued fraction an exact argument and kept the prefactor; its registration
says that beyond 16 the final exponential sets the error.

## The rule

**R1.** A new procedure forms `r = t^2 / df`, `y = r / (1 + r)` and
`x = 1 / (1 + r)` as pairs of Doubles; takes the logarithm of the prefactor
as a pair from `PROB_TryBetaCorePair` (#501, unchanged) at `(x, y, df/2,
1/2)`; takes the continued fraction from `PROB_TryBetaContinuedFractionPair`
(#35, unchanged) at the pair `x`; and returns
`Exp(high) * (1 + low) * fraction / (df / 2) / 2`.

**R2.** `PROB_TryStudentTTail` takes it where `df < 1E5` and `|t| > 16`, after
the cases that answer before the incomplete beta today (zero, the central
expansion, df = 1 and df = 2), which are not changed.

**R3. Where the rule does not apply.** For `|t| <= 16`, for `df >= 1E5`, and
where the new procedure returns FALSE (an argument or `t^2 / df` outside
1E-100 to 1E+100, `x` not below the mean, a deviance without a pair, a
fraction that does not converge), the route is that of today.

The region is a decision, not a finding: in the model the same procedure is
within 8E-16 for every df from 3 to 1E8 and `|t|` from 1.8 to 1000, where
today's routes are within 4.8E-14 up to 16. It is taken beyond 16 below 1E5
only, which is what the issue asks; the rest is not claimed here.

## What calls the changed procedure

Found by scanning every procedure of every module for the name:

| Procedure | Called by |
| --- | --- |
| `PROB_TryStudentTTail` | `K_STATS_StudentT_Cumulative`, `K_STATS_StudentT_Survival`, `PROB_TryStudentTInvTail` |
| `PROB_TryStudentTInvTail` | `K_STATS_StudentT_InverseCumulative` only |

No procedure of `src/` calls the three public functions. So the survival
changes for `|t| > 16`, the cumulative for `t < -16` (for `t > 16` it is one
minus a tail below 1E-9, which the change cannot reach), and the inverse
where its residual is evaluated beyond 16, at probabilities below about
1E-9. `PROB_TryBetaRegularized`, `PROB_TryBetaLowerNearOne`,
`PROB_TryBetaCorePair` and `PROB_TryBetaContinuedFractionPair` keep their
code, so every other function of the library keeps its results.

## Development evidence

`tt_model.py` evaluates the rule in plain binary64, on the Beta core of
`../density_pair_study` and the pair fraction of
`../f_inverse_large_df_study/fix_comparison`. Without the rule it is the
routed model of `../student_t_large_df_study`. It is not bit-faithful: at
`df = 1999, t = 40` it returns 2.8E-13 as Excel does, at `df = 99999, t = 17`
7.2E-14 where Excel measured 1.5E-13. Its figures bound the rule without
predicting observations.

The 32 comparison rows of the issue with `t` above 16 whose value is
representable are development evidence: they were seen in Excel before this
rule was written. Under the rule the model puts all of them within 3E-15.

On 400 random arguments (df from 3 to 1E5, `|t|` from 16 to 1000, tail above
about 1E-280), in the model: today worst 5.6E-13 with 49 above 1E-13; under
the rule worst 1.1E-15, and the pair procedure declines at none.

The model was run on the holdout points before this registration
(`tt_study.py --model`): survival, 102 points, today worst 4.6E-13 with 16
above 1E-13, under the rule worst 8.6E-16; cumulative, 30 points, today worst
3.8E-13 with 6 above 1E-13, under the rule worst 8.6E-16. The points are not
untouched for the model. They are untouched for Excel. The inverse has no
model here.

## The rows

`tt_study.py --build` writes 154 points, 293 rows, in the format of the Excel
floor grid. Every df is an integer, because Excel truncates it.

| Points | Class | What they hold |
| ---: | --- | --- |
| 102 | `sv` | `StudentT_Survival` at 13 df from 5 to 99998 and 14 `t` from 16.03125 to 1E50, where the value is above 1E-290 |
| 30 | `cu` | `StudentT_Cumulative` at `-t` for `t` = 17.5, 27 and 62 |
| 15 | `in` | `StudentT_InverseCumulative` at p = 1E-70, 1E-150 and 1E-270 and five df, where the quantile is beyond -16 |
| 7 | `ed` | `StudentT_Survival` where the rule does not answer: `t = 16` at four df, and `df = 100000` at three `t` above 16 |

No point is a comparison row of the issue and none is in an earlier holdout.
References are 50-digit values from the definitions; a quantile is the root
of the reference tail.

Before the candidate is written, the grid is exported once from the source of
today. `tt_study.py --freeze-before` keeps those verdicts, the errors of the
32 rows of the issue, and the verdicts of the floor grid of #493 and of the
holdouts of #175, #498, #501, #496 and #504 as committed, in
`tt_before.json`.

## Acceptance rules

`tt_study.py --check` implements them on the exported grids.

| Id | Rule |
| --- | --- |
| T1 | Every `sv` and `cu` point is within **1E-14** of the reference |
| T2 | No `sv` or `cu` point is inferior to Excel; every row holds an observation and no library value is an error |
| T3 | Each of the 32 comparison rows of the issue with `t` above 16 is within **1E-14** |
| T4 | Every `in` point is within **1E-12** of the reference and is not newly inferior to Excel |
| T5 | No `ed` observation changes |
| T6 | In the floor grid of #493 and in the holdouts of #175, #498, #501, #496 and #504, no point is newly inferior to Excel with a library error above **3E-14**, and no more points are inferior than before |

Also required in the same wave: the full Excel regression, with the
assertions this pull request adds for the rule in
`tests/M_STATS_PROBDIST_TEST.bas`; the strict accuracy gate on the main grid
and the consolidated holdout with no contract changed; and, after the wave,
`below1e5_study.py --compare` reporting no point above 1E-13 inferior to
Excel.

The wave exports no frozen study grid.

## Predictions

1. T1 to T6 hold.
2. The worst `sv` or `cu` error is below 5E-15; before the rule it is above
   2E-13.
3. Before the rule at least 20 of the 132 `sv` and `cu` points are inferior
   to Excel; after it, none.
4. The 32 rows of the issue are within 5E-15, and `--compare` reports at most
   30 inferior points, all at `|t| <= 16`.
5. Every `in` point is within 1E-13 before and after. This is the least
   certain: the inverse has no model here, and `T.INV` may return no number
   at these probabilities.
6. Outside the Student t cumulative, survival and inverse no observation of
   any exported grid changes; among those three, only rows with `|t|` above
   16 and df below 1E5, or inverses whose quantile is there.
7. Across `|t| = 16` and across `df = 1E5` the two sides now differ in
   accuracy by up to 5E-14: the rule leaves both seams, and they are reported,
   not judged.

## Rejection policy

A failure of T1 to T6, of the regression or of the gate rejects the rule. The
bounds are not retuned and the result is recorded.

## What this does not address

- `|t| <= 16` and `df >= 1E5`: the routes of #213 and of pull request 492
  stay, with their points between 1E-14 and 1E-13 recorded against #494.
- The prefactor of the incomplete beta for every other caller of
  `PROB_TryBetaRegularized` (#495 holds the large shapes).
- Tails below 1E-290.
- The cost per call of the pair route, which is measured before the release.
