# #359: a cheaper beta inverse — result, 2026-11-16

Registration `cefcafd`, with Addendum 1 in `15d8e28`. Baseline: run
43916635364 on `cefcafd`, the source of v1.0.0, evidence `69c5ea6`.
Candidate: run 43920764675 on `15d8e28`, evidence `3485b69`. Both on Excel
16.0 build 20710, 64-bit, on the same runner.

**Verdict by the registered rules: REJECTED.** `analyze.py` finds R0 to R4,
C0 and C1 on every row and line, and C2 on three of its four lines. On the
fourth, Beta (0.7, 25) at p = 0.9, the candidate takes 6.10 ms where the
limit, half of v1.0.0, is 5.80 ms: 1.9 times faster where the rule asks for
2. The bound is not changed after the export.

**Decision: adopted all the same**, by the maintainer on 2026-11-16, knowing
the verdict. C2 was a cost target I set for the part of the routine that
still uses pair arithmetic. Every rule that protects a result holds on a
holdout nothing had seen; on the twelve other lines the cost is back to what
it was before #35; and on the failing line the candidate is faster than
v1.0.0, not slower. Rejecting it would keep the slower routine everywhere.
The four pair lines are the open part of the problem and go to an issue of
their own, #411.

The Excel regression passed in both runs, and the strict accuracy gate
passes on the main grid and on the consolidated holdout with no contract
changed.

## Accuracy

Worst error against the 120-digit reference, relative, on the holdout:

| Kind | Rows | v1.0.0 | Candidate | Rows whose Double changed | Closer | Farther |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `direct_beta` | 42 | 1.19E-14 | 1.04E-14 | 24 | 7 | 17 |
| `small_shape` | 20 | 1.61E-13 | 1.61E-13 | 10 | 3 | 7 |
| `balanced_f` | 36 | 1.93E-14 | 1.92E-14 | 20 | 9 | 11 |
| `large_f` | 36 | 2.87E-15 | 2.87E-15 | 1 | 0 | 1 |
| `branch_seam` | 10 | 1.5E-15 | 1.5E-15 | 5 | 3 | 2 |
| `half_seam` | 10 | 3.44E-15 | 3.69E-15 | 8 | 2 | 6 |
| `student_t` | 24 | 6.92E-15 | 6.92E-15 | 0 | 0 | 0 |

The largest move away from the reference on any row is 1.86E-15, on a
`balanced_f` row; R1 allows 3E-14. The largest move toward it is 2.71E-15.
No row returns a worksheet error in either file.

`small_shape` is above 1E-13 in v1.0.0 already, on one row; the candidate
returns the same Double there. R4 is conditional on the baseline for this
case.

Of the 36 `large_f` rows one changes. Where one degree of freedom is far
above the other the settled iterate is above the mean, the pair loop
answers, and it returns the Double v1.0.0 returns. These are the rows #35
repaired, and they keep its accuracy and most of its cost.

Main grid: 26 observations changed. Consolidated holdout: 28. Every contract
passes with its target and guarantee unchanged.

## Cost

Microseconds per call of the public function, each timing repeated until
0.5 s had elapsed. "Before #35" is the recorded time before, carried to the
baseline run as the registration says.

| Function | Shapes or df | p | v1.0.0 | Candidate | Faster by | Times the cost before #35 | Rule |
| --- | --- | ---: | ---: | ---: | ---: | ---: | --- |
| Beta | (0.5, 0.5) | 0.1 | 1,937.9 | 33.9 | 57x | 1.1 | C1 holds |
| Beta | (0.5, 0.5) | 0.9 | 1,709.0 | 30.5 | 56x | 1.1 | C1 holds |
| Beta | (0.7, 25) | 0.1 | 1,571.7 | 28.1 | 56x | 1.2 | C1 holds |
| Beta | (0.7, 25) | 0.9 | 11,596.7 | 6,103.5 | 1.9x | 113 | **C2 fails** |
| Beta | (2, 5) | 0.1 | 839.2 | 119.2 | 7x | 0.9 | C1 holds |
| Beta | (2, 5) | 0.9 | 656.1 | 28.4 | 23x | 1.1 | C1 holds |
| Beta | (50, 80) | 0.1 | 31,738.3 | 60.1 | 528x | 1.1 | C1 holds |
| Beta | (50, 80) | 0.9 | 7,324.2 | 2,594.0 | 2.8x | 47 | C2 holds |
| F | (5, 10) | 0.1 | 2,563.5 | 44.8 | 57x | 0.8 | C1 holds |
| F | (5, 10) | 0.9 | 3,418.0 | 48.2 | 71x | 0.9 | C1 holds |
| F | (50, 1E8) | 0.1 | 5,493.2 | 82.0 | 67x | 1.2 | C1 holds |
| F | (50, 1E8) | 0.9 | 6,103.5 | 2,960.2 | 2.1x | 11 | C2 holds |
| F | (100, 100) | 0.1 | 27,587.9 | 69.6 | 396x | 1.0 | C1 holds |
| F | (100, 100) | 0.9 | 8,422.9 | 68.7 | 123x | 1.0 | C1 holds |
| F | (1E5, 2) | 0.1 | 892.6 | 373.8 | 2.4x | 6.5 | C2 holds |
| F | (1E5, 2) | 0.9 | 846.9 | 34.3 | 25x | 0.7 | C1 holds |
| Student-t | 3 | 0.9 | 28.6 | 28.1 | | | reported |
| Student-t | 10 | 0.9 | 20.0 | 19.6 | | | reported |
| Student-t | 30 | 0.9 | 179.3 | 181.2 | | | reported |

The split between the twelve lines and the four was fixed from the model
and the source follows it: the four lines are the four that stay in
milliseconds.

A worksheet column of 10,000 F inverses at (100, 100) took one to four
minutes in v1.0.0 and takes about 0.7 s.

## Predictions

| # | Prediction | Outcome |
| --- | --- | --- |
| 1 | R0 to R4 hold on every row | Holds |
| 2 | Most rows change; about as many closer as farther | **False**: 68 of 178 change, 24 closer and 44 farther, none by more than 1.86E-15 |
| 3 | Few `student_t` rows change, or none | Holds: none |
| 4 | C1 holds, most lines within 3 times the cost before #35 | Holds: all within 1.2 times |
| 5 | C2 holds, with a gain between 2 and 10 times | **False**: 1.9 to 2.8 times; one line fails |
| 6 | `small_shape` is the kind most likely to break R1 | It does not: its largest move is 1.68E-15 |
| 7 | Main grid and consolidated holdout change; every contract holds | Holds: 26 and 28 |

I had written before the run that C2 had little margin. The pair loop
still makes at least three pair evaluations from the iterate the first loop
hands it: one for the residual and two for the stagnation guard.

## What this does not show

- The four pair lines: an inverse whose solved member settles at or above
  its mean still costs 0.4 to 6 ms. That is where the F inverse with degrees
  of freedom far apart lives.
- The legacy regime, a lower-tail target that rounds to one: untouched and
  unmeasured here.
- Evaluations per call and terms per evaluation in Excel. They are measured
  in the model only (`FIT_RESULT.md`).
- The cost on other hosts. One run each, on one runner.

## After the verdict

The header of `PROB_TryBetaContinuedFraction` gained an optional argument,
and `../../beta_balanced_survival` binds that procedure's text. The study was
measured again with `measure_balanced_survival.py`: every figure is the
same and only the recorded hash changed, since the default stop is the one
it had.
