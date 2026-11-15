# P4 tail-aware residual — Addendum 1: the oracle's precision, 2026-11-14

Written after the first run of `run_tail_aware.py` and **before any stage-2
row is evaluated**. The candidate, the gates and the rows are unchanged.

## What happened

The first run is kept as `tail_aware_result_blocked.json`. Stage 1 passed 62
of 63 rows; its status is REJECT_OR_BLOCKED and stage 2 did not run, as the
protocol requires.

The one row that did not pass is the probe row p = 1 - 1E-13, (5, 2E6), and
the gate it failed is the oracle's, not the candidate's: the references at
50 and 80 digits agree to about 38 relative digits, short of the 40 the rule
asks. P4's error at that row is 4.9E-17 against the 80-digit reference.

| Digits | Reference F quantile |
| ---: | --- |
| 50 | 14.0103260472254216543088122060242798521097636 |
| 80 | 14.0103260472254216543088122060242798521125517 |
| 120 | 14.0103260472254216543088122060242798521125517 |

The inversion solves on a tail of 1E-13, so a 50-digit solve keeps about 37
digits of it. The 80-digit and 120-digit values agree to every digit shown.

## The change

References are computed at **80 and 120 digits** for every row, with the
same rule: agreement to 40 relative digits, else the row is a blocker.
Errors are measured against the 120-digit value. Stage 2 has rows at
1 - 1E-14, which the 50-digit oracle would have blocked in the same way.

This changes the instrument, not what is measured. Stage 1 is run again in
full under it, and stage 2 runs only if all of it passes. The stage-2 rows
have still not been evaluated by any candidate.
