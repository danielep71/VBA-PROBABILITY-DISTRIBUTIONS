# #35: the tail-aware beta inverse in production — preregistration

Registered 2026-11-14 (Europe/Rome), **before either file of the holdout is
exported**. `f_inverse_pair_baseline.csv` and `f_inverse_pair_grid.csv` are
committed with their design and references only; the observed column of
both is blank. No production source changes in the commit that carries this
file.

## Why a new registration

`fix_comparison/TAIL_AWARE_RESULT.md` records that the rule called P4 did not
pass the gates I had frozen for its Python screen: six rows stay above 2^-48
where the current rule is above it by the same amount, and two rows at the
orientation boundary return about three units in the last place where the
current rule returns one. Those gates are not retuned and that result
stands.

This is a different question with its own evidence: whether P4 may replace
the current rule in `PROB_TryBetaInvRegularized`, judged in Excel by the
accuracy tiers of #327. The bounds below were decided on 2026-11-14, after
the screen and before any row of this holdout was evaluated, by the model
or by Excel. The sixteen points reserved by
`FIX_COMPARISON_PREREGISTRATION.md` §4 have never been evaluated and are
used here, once.

## The rule

`PROB_TryBetaInvRegularized(Probability, ComplementProbability, A, B)`
serves `K_STATS_Beta_InverseCumulative`, `K_STATS_F_InverseCumulative` and
the small-df branch of `K_STATS_StudentT_InverseCumulative`. The candidate
is the rule of `fix_comparison/tail_aware_model.py`:

1. **Orientation.** Solve for the member of the pair `(x, 1 - x)` that is
   at most 1/2: the lower member when `Probability <= I_1/2(A, B)`. The
   other member is then at least 1/2 and `1 - U` is exact enough.
2. **Residual on the smaller tail.** The Newton residual is formed on the
   smaller of the two probabilities, `lower - Probability` or
   `ComplementProbability - upper`, so a far upper tail is not lost in
   `1 - upper`.
3. **Pair continued fraction.** The forward value inside the loop carries
   the continued fraction and its argument as pairs of Doubles, with a stop
   at 2^-100. Where the pair evaluation declines (a non-finite or very
   large term, or the iteration cap), the binary64 continued fraction
   answers, as today.
4. **Stagnation.** When Newton proposes the same Double again, and the two
   neighbouring Doubles bracket the root with a larger residual each, the
   loop ends there.

Nothing else changes: the seed, the bracket, the iteration cap, the three
wrappers and their domain checks.

## The holdout

`build_grid.py` writes 126 rows, the same in both files:

| Kind | Rows | What it holds |
| --- | ---: | --- |
| `holdout_f` | 16 | The reserved points: p in {0.1, 0.4, 0.6, 0.98} and (d1, d2) in {(3E5, 7), (7, 3E5), (50, 1E8), (1E8, 50)} |
| `far_f` | 10 | F at p = 1E-11 and 1 - 1E-11 and five new pairs of degrees of freedom |
| `seam_f` | 10 | F at the orientation boundary of (8, 14) and (14, 8), and 1 and 4 units in the last place either side |
| `t_invariance` | 66 | Student-t at df in {1, 2, 3, 5, 10, 30} and the reserved probabilities: 0.5 +- 2^-53, 2^-43, 2^-33, 2^-23, and 0.625, 0.75, 0.875 |
| `beta` | 24 | Beta at six pairs of shapes, one with both below one, and p in {1E-9, 0.2, 0.8, 1 - 1E-9} |

`0.5 + 2^-53` is not a Double and rounds to 0.5, so each df has eleven
probabilities, one of them exactly 0.5 with reference zero.

The seam probabilities come from the model's forward value at 1/2; the
model's inverse is not run on any row. References are the quantile to 40
digits from 120-digit arithmetic (`benchmark/_ibeta.py`).

`M_STATS_PROBDIST_FIP.bas` calls the public function each row names and
records the result as the bit pattern of its Double. The inputs travel as
bit patterns. `Export_FInversePairBaseline` fills the baseline file from the
source before the change; `Export_FInversePair` fills the other from the
candidate.

## Acceptance rules

`analyze.py` implements them. The error is relative to the reference, and
absolute where the reference is zero.

| Id | Rule |
| --- | --- |
| R0 | Every cell of both files is a Double or a worksheet error; none is blank or a runtime fault |
| R1 | Where both files hold a number, the candidate's error is at most the larger of the baseline's and **1E-14** |
| R2 | Every number the candidate returns is within **1E-12** of the reference: the inverse guarantee of #327 |
| R3 | A row the baseline answers with a number, the candidate answers with a number |
| R4 | On the Student-t rows the candidate returns the baseline's Double, one of its two neighbours, or a value closer to the reference |

R4 is the invariance condition of `FIX_COMPARISON_PREREGISTRATION.md` §4,
unchanged. R1's floor of 1E-14 is the declared limit of what I accept as
rounding difference between two rules; the screen's largest was 1.0E-15.

Also required, in the same wave: the strict accuracy gate on the main grid
and on the consolidated holdout, with no contract relaxed, and the full
Excel regression.

Reported, not gating: the worst error of each file by kind; how many rows
changed, moved closer and moved farther; how many candidate rows are above
2^-48; whether the F and Beta rows are the Doubles of the Python model.

## Cost

The pair continued fraction costs more per iteration than the binary64 one.
There is no time limit in the acceptance rules. The Excel wave of the
candidate measures the cost per call of the three public inverses before
and after, and the figures are published in the README and in the result
record.

## Predictions, fixed before evaluation

1. R0 to R4 hold on all 126 rows.
2. The baseline's worst `holdout_f` error is above 1E-10 and the
   candidate's is below 1E-14.
3. No Student-t row moves by more than one unit in the last place.
4. Some observations of the main grid and of the consolidated holdout
   change for the Beta, F and Student-t inverses; every contract holds
   with its present target and guarantee.
5. A public inverse call costs between two and ten times what it costs
   today.

## Rejection policy

Any failure of R0 to R4, of the accuracy gate or of the regression rejects
the candidate. The bounds and the rows are not retuned against the revealed
files, the current rule stays in production, and the result is recorded. A
later candidate needs its own registration and its own untouched evidence.

## Order of work

1. This file, the grid in two blank copies, the exporter, the analyzer and
   the tests are committed and pushed. No file under `src/` changes.
2. One Excel wave on that commit fills the baseline file, with the main
   grid and the consolidated holdout.
3. The rule changes in `PROB_TryBetaInvRegularized`, with its regressions
   and the cost benchmark.
4. One Excel wave on the candidate fills `f_inverse_pair_grid.csv`.
   `analyze.py` gives the verdict.
