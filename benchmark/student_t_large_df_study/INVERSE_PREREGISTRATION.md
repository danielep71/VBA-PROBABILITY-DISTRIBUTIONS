# Inverse preregistration — Student-t large-df quantile and contracts (#34)

**Status: FROZEN before any VBA is written.** This file fixes:
- a change to the Student-t inverse;
- the predictions its Excel wave must meet;
- a fresh untouched holdout;
- the contracts that follow if the predictions hold.

The implementation follows it and does not amend it. If the design proves
wrong, the failure is recorded against #34 and no contract is adopted.

It governs a change to `PROB_TryStudentTInvTail`
(`src/M_STATS_PROBDIST_TFAMILY.bas`). It also governs the adoption of the three
contracts that `PRODUCTION_PREREGISTRATION.md` §7 froze and `EXCEL_WAVE_RESULT.md`
left unadopted.

## 0. Disclosure: what was seen before this file

- **The first wave.** `EXCEL_WAVE_RESULT.md` records the large-df route's
  Excel wave at `c8bb0ae`:
  - PT1, PT3 and PT4 held;
  - PT2 failed at 7 of 18 inverse rows, 3 to 6 Doubles off the quantile;
  - the cause was the inverse's stopping rule, a 1E-15 relative step.
- **Its observations are known.** These include every Student-t row of the
  main grid and `tchi_large_df`, and the 44 holdout rows of
  `PRODUCTION_PREREGISTRATION.md` §6. None of them is untouched evidence here.
- **The model measurement** of §3 was run while this file was written. The
  polish of §1 was chosen first, from a Python replica of the iteration on the
  failing rows. Its constants were not tuned on the measurement.
- **The new holdout** of §6 does not exist yet.

## 1. The change

The existing loop of `PROB_TryStudentTInvTail` ends on a step of at most
`PROB_EPS = 1E-15` relative. Near the root that step is often a bisection, so it
leaves `X` up to about 6 Doubles from the quantile. The large-df route
(`PROB_TF_StudentTLargeDFTail`) makes the tail accurate enough for those
Doubles to matter.

After the convergence check, and before `Result = X`, the inverse polishes `X`
with plain Newton steps on the same tail:

    If DegreesFreedom >= PROB_T_LDF_MIN_DF And Tail <= PROB_T_LDF_POLISH_MAX_TAIL Then
        For PolishIdx = 1 To PROB_T_LDF_POLISH_STEPS
            If X > PROB_T_LDF_MAX_ABS_T Then Exit For
            If Not PROB_TryStudentTTail(X, DegreesFreedom, TailAtX, FailMsg) Then Exit Function
            If Not PROB_TryStudentTPDF(X, DegreesFreedom, Density, FailMsg) Then Exit Function
            If Density <= 0# Then Exit For
            XNew = X + (TailAtX - Tail) / Density
            If XNew = X Then Exit For
            X = XNew
        Next PolishIdx
    End If

The constants are `PROB_T_LDF_POLISH_STEPS = 3` and
`PROB_T_LDF_POLISH_MAX_TAIL = 0.4`. `PROB_T_LDF_MIN_DF` and
`PROB_T_LDF_MAX_ABS_T` are the route's constants (1E5 and 8).
- **Tail ≤ 0.4.** This is `|p − 0.5| ≥ 0.1` (`p = 0.6` gives exactly the Double
  0.4). It keeps the polish off the central-mass branch, which serves
  `|t| ≲ 1E-4`.
- **Unchanged paths.** Below `df = 1E5`, for tails above 0.4, and for quantiles
  above 8, nothing changes.
- **Forward functions.** The forward route and the public functions do not
  change.

`XNew = X + (TailAtX - Tail) / Density` is evaluated as written:
1. the difference, stored to a Double;
2. the quotient, stored to a Double;
3. the sum.

## 2. Source fidelity

The implementation extends `test_vba_large_df_source.py`. It checks:
- the two constants;
- the polish statements and their order, as in §1;
- their position after the convergence check and before `Result = X`.

Deliberate mutations must each fail it:
- a different step count;
- a different tail bound;
- a missing `XNew = X` exit;
- the polish moved before the convergence check.

## 3. Model measurement (recorded)

`predict_inverse.py` models §1 in binary64, with
`production_route_model.route_tail` as the tail. Its outcome set for a row is
every result the polish returns when:
- it starts within ±32 Doubles of the correctly rounded quantile;
- the tail is shifted by up to ±2 Doubles at every step;
- the density is scaled by 1 ± 1E-9.

The existing loop leaves `X` within about 6 Doubles. The tail shift covers the
model's 1-ulp disagreement with Excel, measured in the first wave's PT1.

`inverse_predictions.json` records:
- **B1, accuracy.** On 72 validation points (12 values of `p` from 0.6 to
  1 − 1E-12; `df` = 1E5, 2E5, 5E5, 1E6, 1E7 and 1E8), every outcome is within
  the bound of §4. The worst is **0.51** of its bound, at `p = 0.75`,
  `df = 2E5`. The bound is largest at `p = 0.6`, about 5E-15.
- **B2, the failing rows.** Each of the 7 rows PT2 failed now has an outcome set
  of 1 or 2 Doubles. The worst is 1.7E-16, against bounds of 5.5E-16 to
  6.2E-16.

## 4. Predictions for the Excel wave

Error is relative error against the committed reference at the row's binary64
arguments. The baseline is the observations committed on `release/1.0.0` at
the base of the implementation, which hold the first wave's export.

- **PI, the polished inverse.** 30 committed `StudentT_InverseCumulative` rows
  reach the polish: 6 in the main grid, 12 in `tchi_large_df` and 12 in the
  holdout. For each of them:
  - the Excel observation is a member of its recorded outcome set;
  - its error is at most **(2^-50 + 2^-52) · min(p, 1−p) / (|q| f(q)) + 2^-51**.
- **PU, unchanged.** Every other row is **bit-identical** to its baseline:
  - every other Student-t row of these three grids, including every forward
    row;
  - every `StudentT_Density` row;
  - every `F_*` row.

  That is 361 rows in the main grid, 144 in `tchi_large_df` and 111 in the
  holdout.
- **PH, the untouched holdout.** The rows of §6 are checked once:
  - a forward row by PT1's rule: within 4 Doubles of
    `production_route_model`, and an error of at most 2^-50;
  - an inverse row by PI's rule, with the outcome set and the bound computed by
    `predict_inverse.outcomes` after the export.
- **The regression is green.** That includes the assertions of §7.

A row outside its prediction falsifies it. It is recorded against #34 before
any contract changes. No prediction is re-derived after the export.

## 5. The check

The implementation adds a check script, as the first wave's
`check_excel_wave.py`. It reads `inverse_predictions.json` and the re-exported
grids, applies §4, and writes its result. A grid none of whose PI rows changed
from the baseline is reported as not re-exported, never as passing.

## 6. Untouched holdout

- **Rows.** The wave appends rows to `benchmark/holdout/holdout_grid.csv`,
  generated by `holdout/generate_holdout.py` with `random.Random(3434)`, drawn
  in this order:
  - **16 forward points.** `t = s·U(0.25, 8)`, with `s` negative for the second
    8; `df = 10^U(5, 8)`. Each gives a `StudentT_Survival` and a
    `StudentT_Cumulative` row, 32 rows in all.
  - **24 inverse points.** First, `τ = 10^U(−14, log10 0.4)`; then
    `p = 1 − τ` for the first 12 and `p = τ` for the last 12; then
    `df = 10^U(5, 8)`. Each gives one `StudentT_InverseCumulative` row.
- **Arguments and references.** The arguments are written as the Doubles they
  are. The references are the direct upper tail at those Doubles, at 50 digits,
  with the quantile from the same oracle.
- **Order of events.**
  1. The contracts of §7 are committed before the rows are exported.
  2. The rows are exported once, in the same Excel job as the regression.
  3. They are checked with §4.

  They are not re-generated or re-drawn after any observation.

## 7. Contracts, if the predictions hold

These are the contracts frozen by `PRODUCTION_PREREGISTRATION.md` §7,
unchanged. They are adopted after the export, and only if PI, PU and PH all
hold and the regression is green:

| Contract | Today | After |
| --- | --- | --- |
| `StudentT_Survival.envelope_domain.output_rel` | 5E-9, df 1E6..1E8, t 0.5..2 | **5E-15**, df 1E5..1E8, \|t\| ≤ 8 |
| `StudentT_Cumulative.envelope_domain.output_rel` | 2E-10, df 1E6..1E8, t 0.5..2 | **5E-15**, df 1E5..1E8, \|t\| ≤ 8 |
| `StudentT_InverseCumulative.envelope_domain.quantile_rel` | none | **5E-15**, df 1E5..1E8, \|p − 0.5\| ≥ 0.1, \|q\| ≤ 8; `p = 0.5` returns exactly 0 |

No threshold is widened. At `p = 0.6` the inverse bound of §4 is 4.98E-15.

**VBA regressions.**
- A new `Test_TF_StudentTLargeDFQuantile` asserts the quantile at the rows PT2
  failed, at **5E-16** relative: `p = 0.99` at `df = 1E6, 1E7` and `p = 0.999` at
  `df = 1E5, 1E6, 1E7`.
- Each reference is the correctly rounded quantile.
- Each recorded outcome set lies within 5E-16 of its reference.
- The first wave observed errors of 6.9E-16 to 9.5E-16 there, so these
  assertions would fail without the polish.
- The suite grows from 950 to 955 assertions.

## 8. Failure conditions

Any one of these stops adoption of the contracts:
- the source-fidelity fixture of §2 fails;
- the Excel regression is not green;
- a PI or PH row is outside its prediction;
- a PU row changes.

A failure is recorded against #34. The polish stays only if:
- every PU row is unchanged;
- every PI observation is no further from its reference than the baseline
  observation of that row.

Otherwise it is reverted.

## 9. Out of scope

- **The forward route** does not change.
- **`|t| > 8` at large df** stays a characterized limitation
  (`numerical_limitations.csv`).
- **Tails above 0.4** (`|p − 0.5| < 0.1`) are not polished and are not in the
  inverse contract.
- **`df > 1E8`** remains rejected.
- **F.** No F source, grid, contract or holdout changes (#35).
