# P4 tail-aware residual — prospective protocol, 2026-11-14

Frozen with `tail_aware_model.py` and `run_tail_aware.py` before the stage-2
rows are evaluated. P3 and every earlier frozen study stay as they are.

## Why P3 is not adoptable

P3 passed its staged screen (53 of 53), whose probabilities run from 0.05 to
0.999. An exploratory probe of farther tails, run on 2026-11-14 before any
production change, shows that P3 loses the upper tail where it solves for a
small member at a probability near one:

| p | (d1, d2) | F-0 error | P3 error |
| --- | --- | ---: | ---: |
| 1 - 1E-10 | (5, 2E6) | 1.9E-12 | 6.0E-9 |
| 1 - 1E-13 | (5, 2E6) | 1.5E-12 | 1.3E-5 |
| 1E-10 | (2E6, 5) | 2.7E-12 | 2.8E-9 |
| 1 - 1E-12 | (8, 5E4) | 4.6E-14 | 6.9E-7 |

P3 orients the solve by geometry and then matches the lower tail of the
solved member. When that tail is a number near one, its residual no longer
holds the upper tail. The current rule (F-0) avoids this by solving on the
smaller probability tail, and pays for it with the lost pair member that #35
is about. Neither orientation alone is right.

## P4

P3 unchanged in orientation (F-A), pair continued fraction, 2^-100 stop,
Newton step, bracket and guarded stagnation. One change: **the residual is
formed on the smaller of the two probability tails of the solved member, and
that tail is computed directly.** The forward evaluation returns both tails:
below the mean the lower one is the continued fraction and the upper its
complement; above it the upper one is the continued fraction and the lower
its complement. The sign convention of the residual is unchanged.

## Disclosure

The probe that found the defect also evaluated P4, on the ten probe rows and
on nine revealed rows, before this file was written. Those nineteen rows are
fitting data. The stage-2 rows below were not evaluated by any candidate.

## Stage 1 — revealed rows

The 53 rows of `stagnation_result.json`, with the gates P3 met there:
retained, revealed-screen and new rows below 2^-48; controls within four
times F-0; fitting rows within twice F-0 and below 2^-48 wherever F-0 is.
And the ten probe rows (`PROBE_ROWS` in the runner): below 2^-48 and no
worse than F-0.

Median Newton count over the 24 fitting rows at most F-0's plus two.

## Stage 2 — only if stage 1 passes in full

47 rows, none evaluated before:

- far tails: p in {1E-9, 1 - 1E-9, 1E-14, 1 - 1E-14} crossed with
  (d1, d2) in {(4, 3E5), (3E5, 4), (60, 5E8), (5E8, 60), (300, 500)};
- middle: p in {0.3, 0.7} crossed with {(9E4, 12), (12, 9E4), (4E6, 4E6)};
- orientation seam: for (d1, d2) in {(6, 10), (10, 6), (80, 120)}, the Double
  `h` that the model's forward evaluation gives for the lower tail at 1/2,
  and `h` moved by 1 and by 4 units in the last place either way;
- tail seam: p in {the predecessor of 1/2, 1/2, its successor} crossed with
  {(7, 2E4), (2E4, 7)}.

Gates per row: P4 relative F error below 2^-48, and at most
`max(2 * F-0 error, 2^-52)`. The floor of 2^-52 is stated here, before
evaluation: below one unit in the last place a factor of two between two
correct answers is rounding, not regression. Median Newton count over the
stage-2 rows at most F-0's plus two.

## Oracle and records

As in the P3 screen: the checked continued-fraction oracle of
`run_extended_fit.py`, references at 50 and 80 digits agreeing to 40
relative digits, exact binary64 identities, errors, iteration counts and
source hashes in `tail_aware_result.json`. A failure stops the later stage
and nothing is tuned. Exceptions and unstable references are blockers.

## What a pass means

SCREEN_PASS_NOT_ADOPTION, as before. A pass selects P4 as the candidate to
implement. The frozen holdout grid and the Student-t invariance set of
`FIX_COMPARISON_PREREGISTRATION.md` are not evaluated here; they are kept
for the production implementation in Excel. No production VBA edit, holdout
inspection, merge or closure happens under this protocol.
