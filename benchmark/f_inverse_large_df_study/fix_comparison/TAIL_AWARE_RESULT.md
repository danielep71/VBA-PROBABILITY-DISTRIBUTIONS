# P4 tail-aware residual: the staged screen does not pass, 2026-11-14

Protocol, model and runner frozen at `e78fddd`; Addendum 1 (oracle at 80 and
120 digits) at `edb5e68`, before any stage-2 row was evaluated. Machine
evidence: `tail_aware_result.json`; the first, blocked run is
`tail_aware_result_blocked.json`.

**Status: REJECT_OR_BLOCKED.** Stage 1 passes 63 of 63. Stage 2 passes 39 of
47, and the protocol counts any failure. The gates are not retuned here.

| Population | Passed | P4 worst | F-0 worst | P3 worst |
| --- | ---: | ---: | ---: | ---: |
| Retained affected | 8/8 | 2.92E-15 | 9.31E-9 | 2.92E-15 |
| Retained controls | 5/5 | 3.44E-15 | 3.44E-15 | 3.44E-15 |
| Revealed P2 screen | 8/8 | 2.03E-15 | 8.64E-12 | 2.03E-15 |
| Original fitting | 24/24 | 2.98E-15 | 7.61E-9 | 6.57E-15 |
| P3's new rows | 8/8 | 6.14E-16 | 2.07E-13 | 6.14E-16 |
| Far-tail probe | 10/10 | 3.35E-15 | 2.74E-9 | 1.30E-5 |
| Stage 2, far tails | 19/20 | 8.64E-15 | 5.74E-10 | 7.55E-5 |
| Stage 2, middle | 6/6 | 9.63E-16 | 1.24E-13 | 9.63E-16 |
| Stage 2, orientation seam | 8/15 | 6.66E-15 | 6.51E-15 | 6.51E-15 |
| Stage 2, tail seam | 6/6 | 3.07E-16 | 2.91E-13 | 3.07E-16 |

Median Newton count: fitting 11.5 for F-0 and 9.5 for P4; stage 2, 22 and 15.
Both median gates pass. Every reference agrees at 80 and 120 digits.

## The eight rows that fail

| Row | F-0 | P3 | P4 | Gate failed |
| --- | ---: | ---: | ---: | --- |
| p = 1E-14, (4, 3E5) | 8.64E-15 | 8.64E-15 | 8.64E-15 | below 2^-48 |
| five probabilities around 0.97806, (80, 120) | 5.5E-15 to 6.5E-15 | 5.5E-15 to 6.5E-15 | 5.5E-15 to 6.7E-15 | below 2^-48 |
| p = 0.7734374999999998, (6, 10) | 3.19E-16 | 3.19E-16 | 7.19E-16 | at most max(2 F-0, 2^-52) |
| p = 0.2265625000000002, (10, 6) | 1.41E-16 | 1.41E-16 | 6.96E-16 | at most max(2 F-0, 2^-52) |

Two different things are in that table.

**Six rows miss an absolute bound that the current rule misses too.** At
those rows F-0, P3 and P4 return the same error to within eight percent,
between 5.5E-15 and 8.6E-15, above 2^-48 = 3.55E-15. The earlier protocols
let a fitting row stay above 2^-48 where F-0 already was; I wrote the
stage-2 gate without that clause. P4 does not make these rows worse, and
nothing in this study makes them better: the error there is not the pair
reconstruction of #35.

**Two rows are a real, small difference.** One unit off the orientation
boundary, P4 returns about three units in the last place where F-0 and P3
return one or less. Both members of the pair are near 1/2 there, so nothing
is lost; the two rules round differently. It is above the floor I froze.

## What the screen does show

- On every row where #35's defect is present, P4 brings the error from as
  much as 9.3E-9 to below 3.5E-15.
- P3's loss of the far upper tail, up to 7.55E-5 on the new rows, does not
  occur in P4.
- P4 is never worse than F-0 by more than 1.0E-15 in absolute terms on any
  of the 110 rows.

## What it does not show

The frozen gates are not met, so P4 is not selected by this protocol. A
decision is needed on which gates a production candidate must meet: the
study's absolute 2^-48 and factor-of-two rules, which the current rule
itself fails on six of these rows, or the tiers of #327, under which the
inverse guarantee is 1E-12 and the distance to the binary64 floor is
reported. That decision is the maintainer's, and a new registration with
new rows follows it. The frozen holdout grid and the Student-t invariance
set of `FIX_COMPARISON_PREREGISTRATION.md` remain unevaluated.
