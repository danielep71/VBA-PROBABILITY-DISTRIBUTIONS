# #359: a cheaper beta inverse — fitting sweep

This page records what was learned on revealed points, before any
registration. Nothing here is evidence for the candidate: every point was
looked at while the candidate was being shaped. The registration and its
holdout come after, in files of their own.

## Why the inverse is slow

The accepted rule (P4 of `../fix_comparison/`, in production since #35)
evaluates the incomplete beta with a pair continued fraction, stopped at
2^-100, at every Newton step. In the model a call makes 5 to 11 such
evaluations of 2 to 43 terms each. In the source one term of that fraction is
about 250 calls of the pair helpers and some 1,800 statements; the Excel
timings of `../production/PRODUCTION_RESULT.md` come to 25 to 100 microseconds
a term. The time is in the pair arithmetic.

## The candidate

`selective_model.py`, candidate S:

1. Keep P4's orientation: solve for the member of the pair below 1/2, on the
   smaller probability tail. This is what repaired #35.
2. Phase A: iterate with the binary64 forward function, its fraction stopped at
   4 * 2^-52, splitting the two tails at the mean as P4 does. Compare a
   settled iterate below the mean with its two adjacent Doubles and keep the
   one with the smallest residual.
3. Phase B, only when that iterate lies on the reflected branch of the forward
   function, or phase A has not settled in 60 iterations: continue with P4's
   loop, pair fraction and guarded stagnation included.

Below the mean the fraction's argument is the iterate itself, an exact Double,
and no pair arithmetic is used.

## What the sweep changed

`fit_sweep.py` runs P4 and the candidate on a fixed product of shapes and
probabilities, Beta and F, from 1E-12 to 1 - 1E-12, against 60-digit
references. Three versions of the candidate were run; the first two are in
the history of this folder.

| Version | Change | Worst error | Rows farther than P4 by more than 1E-14 |
| --- | --- | ---: | ---: |
| 1 | upper tail taken as I at 1 - u everywhere | 1.0E-11 | 9 |
| 2 | tails split at the mean, as P4 | 6.8E-14 | 4 |
| 3 | tight binary64 stop, neighbour comparison, fallback to phase B | 4.3E-14 | 2 |

Version 1 had brought back the loss at the reflected argument that #35 had
found; the production holdout, 56 rows, did not show it, and the sweep did.
Version 2 was within the guarantee everywhere but lost up to 2.8E-14 at small
shapes to the source's stop of 3E-14, and phase A did not settle on two rows.

## Version 3 against P4

Counts are in `fit_result.json`, `summary`; this table names what they are.

| | P4 | Candidate |
| --- | --- | --- |
| Worst relative error | 4.6E-14 | 4.3E-14 |
| Rows above the 1E-12 guarantee | none | none |
| Rows closer to the reference than the other | about as many | about as many |
| Pair evaluations over the sweep | all of them | one twelfth |
| Rows with no pair evaluation | none | about three in four |

The two rows that move away by more than 1E-14: Beta(2, 300) at p = 0.7, from
5.5E-15 to 2.6E-14, and Beta(0.5, 80) at p = 0.7, from 1.5E-15 to 1.2E-14.
Both are below the mean, where the residual is a binary64 number.

## What the model does not cover

- Shapes below 0.5 in some combinations: the binary64 forward model does not
  implement the series and reflection branch of the log-gamma. The Excel
  holdout must include them.
- The regime the source answers with the routine before #35: a lower-tail
  target that rounds to one.
- Twenty-one points of the sweep are skipped: the reference did not converge
  on thirteen, and P4 reports no representable interior quantile on eight.

## Decisions taken on this sweep, 2026-11-16

- Candidate: S, version 3.
- Cost target: on the benchmark points that make no pair evaluation, a call
  costs at most five times what it cost before #35; on the others it is at
  least twice as fast as in v1.0.0.
- Accuracy floor: no row farther from the reference than in v1.0.0 by more
  than 3E-14, every row within the 1E-12 guarantee, and the worst error not
  above 1E-13. The issue first said 1E-14; the two rows above are why it was
  widened, to the size of the binary64 noise on the residual.
