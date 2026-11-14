# #35: the tail-aware beta inverse in production — result, 2026-11-14

Registration `7d517cf`, Addendum 1 `93bdf3a`, candidate `f0c0d5f`.
Baseline export: run 43606344252, evidence `c78bc47`. Candidate export:
run 43609869324, evidence `4fd5072`. Both on Excel 16.0 build 20430, 64-bit.

**Verdict: ACCEPTED.** `analyze.py` finds R0 to R4 on all 126 rows. The
Excel regression passed 1,940 of 1,940 assertions and the strict accuracy
gate passed on the main grid and on the consolidated holdout with no
contract changed.

**The cost is far above what I predicted.** A Beta or F inverse takes 6 to
580 times longer than before: 0.6 to 28 ms per call where it took 0.02 to
0.25 ms. The registration set no time limit, so this does not change the
verdict; it is the main thing to know about this change. Figures below.

## Accuracy

Worst error against the 120-digit reference, relative:

| Kind | Rows | Before | After | Rows whose Double changed |
| --- | ---: | ---: | ---: | ---: |
| `holdout_f`, the reserved points | 16 | 1.56E-10 | 2.55E-15 | 11 |
| `far_f` | 10 | 3.54E-9 | 1.77E-15 | 5 |
| `seam_f` | 10 | 4.48E-15 | 4.57E-15 | 8 |
| `beta` | 24 | 1.1E-13 | 3.73E-14 | 11 |
| `t_invariance` | 66 | 0.25 | 0.25 | 0 |

Of the 35 rows that changed, 24 moved closer to the reference and 11
farther. The largest move away is on `ho012`, from 4.9E-16 to 1.1E-15; the
largest error among the eleven is 4.57E-15, on a seam row that was at
4.06E-15. R1's floor is 1E-14.

No row returns a worksheet error in either file.

### The Student-t rows

All 66 return the same Double before and after. R4 holds trivially, and
the timing below shows why: at these degrees of freedom and probabilities
the Student-t inverse does not reach the routine that changed. The set
reserved in `FIX_COMPARISON_PREREGISTRATION.md` §4 therefore says nothing
about the shared path; it does record, in Excel, the defect of #22 near the
median, which this change leaves as it was.

### Rows above 2^-48

Reported, not gating: 53 candidate rows are above 2^-48 = 3.55E-15. Forty
are the Student-t rows of #22. The other thirteen are between 3.55E-15 and
3.73E-14, the largest on a Beta row whose baseline error was 1.1E-13.

### The Python model

Reported, not gating: 21 of the 60 F and Beta rows are the Doubles of
`tail_aware_model.py`. The model was faithful enough to choose the rule;
it is not a bit-exact image of the VBA, whose logarithm, exponential and
log-beta round differently. The VBA also keeps the representability check
of the source for shapes below one, which the model replaces.

### Main grid and consolidated holdout

39 observations of the main grid and 41 of the consolidated holdout
changed. Every contract passes with its target and guarantee unchanged.
`F_InverseCumulative.validated`: quantile error 7.77E-11 before, 1.51E-12
after; tail residual 6.91E-11 before, 6.44E-12 after.

## Predictions

| # | Prediction | Outcome |
| --- | --- | --- |
| 1 | R0 to R4 hold on all 126 rows | Holds, with R2 as Addendum 1 leaves it |
| 2 | Baseline `holdout_f` worst above 1E-10, candidate below 1E-14 | Holds: 1.56E-10 and 2.55E-15 |
| 3 | No Student-t row moves by more than one unit | Holds: none moves |
| 4 | Some main-grid and holdout observations change; every contract holds | Holds: 39 and 41 |
| 5 | A public inverse call costs two to ten times what it costs today | **False**: 6 to 580 times for Beta and F |

## Cost

Milliseconds per call of the public function, each timing repeated until
0.5 s had elapsed. Before: run 43608731176 on `b1462b6`, the source before
the change. After: run 43609869324 on `f0c0d5f`. Same runner.

| Function | Shapes or df | p | Before, ms | After, ms | Ratio |
| --- | --- | ---: | ---: | ---: | ---: |
| Beta | (0.5, 0.5) | 0.1 | 0.027 | 1.686 | 61x |
| Beta | (0.5, 0.5) | 0.9 | 0.025 | 1.503 | 60x |
| Beta | (0.7, 25) | 0.1 | 0.021 | 1.396 | 67x |
| Beta | (0.7, 25) | 0.9 | 0.049 | 10.559 | 216x |
| Beta | (2, 5) | 0.1 | 0.121 | 0.751 | 6x |
| Beta | (2, 5) | 0.9 | 0.024 | 0.591 | 25x |
| Beta | (50, 80) | 0.1 | 0.049 | 28.320 | 582x |
| Beta | (50, 80) | 0.9 | 0.051 | 6.744 | 133x |
| F | (5, 10) | 0.1 | 0.055 | 2.365 | 43x |
| F | (5, 10) | 0.9 | 0.046 | 2.991 | 64x |
| F | (50, 1e+08) | 0.1 | 0.062 | 4.944 | 80x |
| F | (50, 1e+08) | 0.9 | 0.246 | 5.463 | 22x |
| F | (100, 100) | 0.1 | 0.061 | 24.292 | 398x |
| F | (100, 100) | 0.9 | 0.062 | 7.446 | 119x |
| F | (100000, 2) | 0.1 | 0.051 | 0.793 | 16x |
| F | (100000, 2) | 0.9 | 0.043 | 0.759 | 18x |
| Student-t | 3 | 0.9 | 0.044 | 0.025 | 0.6x |
| Student-t | 10 | 0.9 | 0.034 | 0.022 | 0.6x |
| Student-t | 30 | 0.9 | 0.214 | 0.153 | 0.7x |

The Student-t ratios below one are differences between two runs on the
same code path, not a gain.

Where the time goes, by reading the code and not yet by measurement: every
Newton step evaluates the continued fraction in pair arithmetic to 2^-100,
about a hundred floating-point operations per term where the binary64
fraction used ten and stopped at 3E-14; the orientation costs one more
evaluation at 1/2; and a settled iterate costs two more at its neighbours.
The balanced cases, where the fraction converges slowest, are the dearest.

A worksheet column of 10,000 F inverses at (100, 100) took about 0.6 s and
now takes one to four minutes.

## What this does not show

- Nothing about the Student-t path: see above.
- No row with a shape above 3E9 or a probability beyond 1E-11 from 0 or 1.
- The cost on other hosts. It was measured on one runner.

## After the verdict: probabilities within 2^-53 of zero

Review of pull request 358 found two failures of the candidate, both in one
regime: a probability, or a complement, below 2^-53 whose quantile lies on
the far side of 1/2. The member the rule solves for then has a lower-tail
target of one in binary64.

| Call | Candidate `f0c0d5f` | Source before #35 |
| --- | --- | --- |
| `K_STATS_Beta_InverseCumulative(1E-20, 100, 2)` | `#VALUE!`: the normal seed has no quantile at one | 0.6080769921891604 |
| `K_STATS_Beta_InverseCumulative(1E-200, 1000, 1)` | no convergence: the series seed starts at 0.001, the root is at 0.369 | 0.630957344480193 |

`tail_aware_model.py` fails at the same points; it is left as the record of
the screen.

**What I tried and withdrew.** `f6b26cd` took the normal seed from the upper
tail, and passed Excel with the holdout unchanged (run 43618326390,
evidence `40413b7`). For the second case I modelled a mirrored series seed
and a step to the neighbouring Double where Newton stalls. A sweep of
1,451 points of the regime against the model of the source before #35 gave
53 points that fail only with the new rule so repaired, and 93 that fail
only with the old one; both models fail on 566, so the models themselves
are not reliable evidence there. I stopped repairing a regime that no
registered evidence covers.

**What is in the source.** `b939c76`: where the lower-tail target of the
member below 1/2 rounds to one, and where the evaluation at 1/2 fails,
`PROB_TryBetaInvRegularized` runs the routine it had before #35,
unchanged. Everywhere else it is the accepted candidate; the seed change
of `f6b26cd` is withdrawn. Two assertions guard the two calls above.

Run 43624391221 on `b939c76` exported the holdout, the main grid and the
consolidated holdout again: all three came back with the same bytes as
accepted, and the regression passed 1,942 of 1,942 assertions (evidence
`2e2597a`). The verdict and every figure above stand for `b939c76`.

**What this leaves open.**

- In that regime the pair reconstruction is not corrected: the small member
  is still `1 - U`, as before #35. No row of any grid measures it.
- The holdout had no row whose complement rounds to one. The rule was wrong
  there and nothing I had registered would have found it; review did.
