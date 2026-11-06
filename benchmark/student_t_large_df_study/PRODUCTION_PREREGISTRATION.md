# Production preregistration — Student-t large-df tail route (#34)

**Status: FROZEN before any VBA is written.** This file fixes the production
route, the arithmetic VBA will use, the predictions the Excel wave must meet,
the untouched holdout and the contracts that follow if the predictions hold.
The implementation commit that follows implements it and does not amend it; if
the design proves wrong, the failure is recorded against #34 and a new
preregistration is written.

Governs the production change to `PROB_TryStudentTTail`
(`src/M_STATS_PROBDIST_TFAMILY.bas`) and its Excel wave. Nothing here changes a
`.bas` file, a contract, a grid, an observation or a manifest.

## 0. Disclosure: what was seen before this file

- **The study** (`PREREGISTRATION.md`, `results.json`) adopted the order-5
  expansion for `df ≥ 1E5`, `|t| ≤ 8`, from its fitting points. Its forward and
  inverse holdouts were evaluated twice (Amendment 1). They are not used here as
  untouched evidence.
- **Finding F1** (`PRODUCTION_FEASIBILITY.md`) measured the route on the
  replaced normal kernel. It failed §8's improvement requirement there, so this
  design waited for the kernel repair. #72 landed that repair through #81, with
  Excel evidence.
- **The committed observations.** The Student-t observations of the main grid
  (`envelope_domain`) and of `tchi_large_df` are the defect's evidence and have
  been read. The predictions in §5 are about their re-export, not a fit to them.
- **The model measurement** in `production_predictions.json` (§4) was run while
  this file was written, with the route already fixed by the study. No route
  parameter was chosen from it.
- **The independent holdout** (`benchmark/holdout/holdout_grid.csv`) holds no
  Student-t row, and was not read. The new holdout rows of §6 do not exist yet.

## 1. The route

Inside `PROB_TryStudentTTail`, after the `AbsX = 0`, central-mass, `df = 1` and
`df = 2` branches and before the incomplete-beta route:

    If DegreesFreedom >= 1E5 And AbsX <= 8 Then
        Result = S(AbsX; DegreesFreedom)          ' the expansion below
        PROB_TryStudentTTail = True
        Exit Function
    End If

with the order-5 expansion of the study,

    S(t; df) = Q(t) + phi(t) * sum_{k=1..5} g_k(t) / df^k.

Outside the predicate the incomplete-beta route is unchanged. The route returns
the upper tail at `|t|`, so the public functions keep their sign handling:
- the survival is `S` for `t >= 0` and `1 - S` otherwise;
- the cumulative is `S` for `t < 0` and `1 - S` otherwise;
- the inverse inverts the same tail, by its existing iteration.

`df` above `PROB_T_MAX_DF = 1E8` is still rejected after the tail is computed.

## 2. Arithmetic

VBA evaluates the route statement by statement, each result stored to a Double:

1. `U = AbsX * AbsX` and `W = 1# / DegreesFreedom`.
2. Each `g_k` is odd: `g_k(t) = t * P_k(U)`. Each `P_k` is evaluated by Horner in
   `U`, from its highest power down, one statement per step.
3. `Series = P_5`, then `Series = Series * W + P_k` for `k = 4, 3, 2, 1`, then
   `Series = Series * W`.
4. `Correction = AbsX * Series`, then
   `Correction = PROB_T_INV_SQRT_TWO_PI * PROB_ExpNegHalfSquare(AbsX) * Correction`
   in that order.
5. `Result = PROB_NormalUpperTail(AbsX) + Correction`.

**Coefficients.**
- Every coefficient of every `P_k` is the Double nearest its exact rational in
  `coefficients.json`.
- It is declared as a `hi + lo` pair of 15-significant-digit literals whose
  Double sum is exactly that Double. The VBE keeps only 15 digits. These are the
  pairs `production_route_model.py` prints and `production_predictions.json`
  records.
- `PROB_T_INV_SQRT_TWO_PI` is the same `hi + lo` pair as the normal kernel's
  `1/sqrt(2 pi)`.

`PROB_NormalUpperTail` and `PROB_ExpNegHalfSquare` are the repaired kernel of
#72, called as they are.

## 3. Source fidelity

The implementation adds `test_vba_large_df_source.py` to the evidence suite. It
checks that:
- every coefficient literal pair in the VBA evaluates exactly to the Double
  recorded here;
- the predicate, the statement order and the operands are the ones of §2;
- the regression references of §7 are reproduced by the model.

As for the normal kernel, deliberate mutations must each fail it:
- a coefficient's last digit;
- a reassociated product;
- a dropped order;
- the predicate bound.

## 4. Model measurement (recorded)

`production_route_model.py` models §1–§2. It takes `Q` and `phi` from
`normal_tail_study/kernel_model.py`, whose repaired kernel reproduces the
committed Excel observations within 4 ulps. `predict_production.py` records:

- **A1, accuracy.** On the validation set V (10 values of `df` from 1E5 to 1E8;
  `t` from 1E-4 to 8, including both sides of the kernel's seams at 3 and
  7.07106781186547), 1670 points reach the route. The worst relative error
  against the study's primary oracle `I_z(df/2, 1/2) / 2` at 60 digits is
  **4.32E-16**, at `t = 4.55`, `df = 1E6`. It is the same at 100 digits, and
  below `B = 2^-50`.
- **A2, improvement.** At the 36 study fitting points in the region, the route
  is more accurate than the current continued fraction wherever that misses `B`.
  The route's worst there is 3.73E-16.
- **A3, seams.** These record the accuracy on each side, not a requirement:
  - *df seam.* Just below `df = 1E5`, the retained continued fraction reaches
    1.25E-12 at `t = 2`; the route at `df = 1E5` is below 4E-16.
  - *t seam.* Just above `|t| = 8`, the continued fraction reaches 8.7E-13 at
    `df = 1E5` and 4.7E-10 at `df = 1E8`. That residual is out of scope (§9).

## 5. Predictions for the Excel wave

Error is relative error against the committed reference at the row's binary64
arguments. `production_predictions.json` records every row's prediction.

- **PT1, forward.** For every committed `StudentT_Survival` and
  `StudentT_Cumulative` row the route reaches (58 rows: 18 in the main grid, 40
  in `tchi_large_df`):
  - the Excel observation is within **4 ulps** of the predicted Double;
  - its error is **≤ 2^-50**.
- **PT2, inverse.** For every committed `StudentT_InverseCumulative` row whose
  quantile the route reaches (18 rows), the quantile's relative error is
  **≤ (2^-50 + 2^-52) · min(p, 1−p) / (|q| f(q)) + 2^-51**:
  - `f` is the Student-t density;
  - the first term carries the tail bound through the inverse's conditioning;
  - the last allows the iteration's stopping rule and the final rounding.
  - The `p = 0.5` rows stay exactly 0.
- **PT3, unchanged.** Every other Student-t row of these grids, every
  `StudentT_Density` row and every `F_*` row is **bit-identical** to its
  baseline observation. That covers the balanced-huge F comparator at
  `df1 = df2 = 1E8`: the change is Student-t specific and must leave F
  unchanged.
  - **Main grid baseline:** its committed observations, exported from current
    source by #81's last export (`0c8c997`).
  - **`tchi_large_df` baseline:** its committed observations are stale. 89
    non-density Student-t rows above `df = 1E8` hold numbers, although the
    current public functions reject those arguments through
    `PROB_T_ValidateEnvelope`; they predate that envelope. Before the
    implementation is exported, the wave therefore exports this grid once from
    the base commit, by the same unattended path. That export is the PT3
    baseline for this grid.
- **PT4, untouched holdout.** The rows of §6 are checked once, with the PT1 and
  PT2 rules.

`tchi_large_df` is exported by its own exporter, `Export_TChiLargeDf`. Like
`survival_boundary`, it finds its grid beside the workbook. The wave adds it to
the unattended export before it takes the baseline.

A row outside its prediction falsifies it. It is recorded against #34 before
any contract changes. No prediction is re-derived after the export.

## 6. Untouched holdout

- **Rows.** The wave appends new rows to `benchmark/holdout/holdout_grid.csv`.
  They are generated by `holdout/generate_holdout.py` with `random.Random(34)`,
  drawn in this order:
  - **16 forward points.** `t = s·U(0.25, 8)`, with sign `s` negative for the
    second 8; `df = 10^U(5, 8)`. Each point gives a `StudentT_Survival` and a
    `StudentT_Cumulative` row, 32 rows in all.
  - **12 inverse points.** `p = U(0.6, 0.999)` for the first 6 and `1 − U(0.6,
    0.999)` for the last 6; `df = 10^U(5, 8)`. Each gives one
    `StudentT_InverseCumulative` row.
- **Arguments and references.** The arguments are written as the Doubles they
  are. The references are the direct upper tail at those Doubles, at 50 digits,
  with the quantile from the same oracle.
- **Order of events.**
  1. The contracts of §7 are committed before the rows are exported.
  2. The rows are exported once, in the same Excel job as the regression.
  3. They are then checked with PT1 and PT2.

  They are not re-generated or re-drawn after any observation.

## 7. Contracts, if the predictions hold

These are frozen now. They are adopted after the export, and only where PT1,
PT2 and PT4 hold:

| Contract | Today | After |
| --- | --- | --- |
| `StudentT_Survival.envelope_domain.output_rel` | 5E-9, df 1E6..1E8, t 0.5..2 | **5E-15**, df 1E5..1E8, \|t\| ≤ 8 |
| `StudentT_Cumulative.envelope_domain.output_rel` | 2E-10, df 1E6..1E8, t 0.5..2 | **5E-15**, df 1E5..1E8, \|t\| ≤ 8 |
| `StudentT_InverseCumulative.envelope_domain.quantile_rel` | none | **5E-15**, df 1E5..1E8, \|p − 0.5\| ≥ 0.1, \|q\| ≤ 8 |

No threshold is widened. The inverse domain excludes `|p − 0.5| < 0.1`, where
the quantile's conditioning `min(p, 1−p) / (|q| f(q))` exceeds 4 and the PT2
bound passes 5E-15.

**VBA regressions** that come with the route:
- **Inside the region, at 2E-15 relative:**
  - the `t = 2` survival and `p = 0.99` inverse ladders at `df = 1E6, 1E7, 1E8`;
  - the clean controls `t = 1` and `p = 0.9`;
  - `p = 0.5`;
  - both signs;
  - the seam points themselves: `df = 1E5` and `|t| = 8`.
- **Just outside the region:** the continued fraction still serves these
  points, and the controls only show that it does. They use its characterized
  accuracy, not the route's:
  - 2E-12 relative at the Double below `df = 1E5`, for `|t| ≤ 8` (1.25E-12
    measured at `t = 2`, §4);
  - 1E-9 relative at the Double above `|t| = 8`, for `df ≤ 1E8` (4.7E-10
    measured at `df = 1E8`, §4 and §9).

The forward references are correctly rounded and reproduced by the model. The
inverse checks are round trips through the survival.

## 8. Failure conditions

Any one of these stops adoption of the contracts:
- the source-fidelity fixture of §3 fails;
- the Excel regression is not green;
- a PT1, PT2 or PT4 row is outside its prediction;
- a PT3 row changes.

A failure is recorded against #34. The route stays only if every PT3 row is
unchanged and every PT1 and PT4 error is below the current route's error at the
same row; otherwise it is reverted.

## 9. Out of scope

- **`|t| > 8` at large df** stays on the continued fraction. Its error reaches
  4.7E-10 at `t = 8 + 1 ulp`, `df = 1E8`, where the tail is about 6E-16. The wave
  records it in `numerical_limitations.csv` as a characterized limitation, and
  the contracts above exclude it.
- **`df > 1E8`** remains rejected (`PROB_T_MAX_DF`).
- **F.** No F source, grid, contract or holdout changes (#35 owns the unbalanced
  F inverse).
