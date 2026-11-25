# #365: the Student-t tail below df = 1E5 — result, 2026-11-25

Registration `0e8072d`, candidate `23095a0`, evidence `834c60e` and `6a7d58b`. Excel runs 45618391642 (the
rule, with the main grid and the holdout) and 45625806466 (the 24 rows of
Amendment 1), Excel 16.0 build 20719, 64-bit.

**Verdict: ACCEPTED.** `below1e5_study.py --check` finds G1 to G5. The Excel
regression passed 3,067 of 3,067 assertions. The main grid and the
consolidated holdout came back with no observation changed, and the strict
accuracy gate passed with no contract changed.

## Forward

Relative error against the 80-digit reference, on the rows the rule serves
(`1999 <= df < 1E5`, `|t|` from about 1.732 to 16):

| Rows | Before (run 45602906527) | After |
| --- | ---: | ---: |
| The 98 map rows | worst 3.16E-12, 36 above 1E-13, 7 above 1.3E-12 | worst 1.77E-14 |
| 144 holdout survivals | not exported before | worst 2.87E-14 |
| 36 holdout cumulatives at `-t` | not exported before | worst 1.69E-14 |

The worst row of the map, `t = 2` on the Double below 1E5, went from
3.16E-12 to within 7E-15.

The 216 map rows the rule does not serve hold the observations of before,
bit for bit (G2).

## The seam at df = 1E5

Across df = 1E5 at `t = 2, 4, 8`, both signs, the survival and the
cumulative move against the required direction by at most 6.4E-15 (G3 allows
1E-13); before, the step at `t = 2` was 3.16E-12. At `t = 0.5` and `t = 1`
the seam rows did not change: the rule does not serve them, and their steps
stay within 5.4E-16.

This is a bound, as registered. The survival still rises by 6.4E-15 from the
Double below 1E5 to 1E5 at `t = 8`: I do not claim strict monotonicity.

## Inverse

The 12 holdout inverses are within 7.9E-16 (G4 allows 1E-12). The three
round trips of the #210 seam rows just below df = 1E5 changed with the tail
and are within 1.2E-15.

## Predictions

| # | Prediction | Outcome |
| --- | --- | --- |
| 1 | G1 to G5 hold | held |
| 2 | Worst served forward error below 5E-14 | held: 2.87E-14 |
| 3 | Worst step across df = 1E5 below 2E-14; `t = 0.5` and `t = 1` unchanged | held: 6.4E-15; unchanged |
| 4 | Every holdout inverse within 5E-14 | held: 7.9E-16 |
| 5 | No observation changes for `df >= 1E5`, `df < 1999`, `\|t\| > 16` or `\|t\| <= 1.73` | held in the three exported grids: the 109 changed observations are all served rows, or inverses whose quantile is served |
| 6 | Rows above 1E-13 remain for `df < 1999` and `t > 16`, within 1.3E-12 | held: 24 map rows, 12 with `df < 1999` (worst 2.5E-13) and 12 with `t > 16` (worst 8.2E-13) |

## Reported, not gating

- Across `|t| = 16`, at df = 2048, 9999 and 50000, the survival does not
  rise. The Double above 16 is on the route of before and is off by up to
  3.6E-14.
- Across df = 1999, where the rule begins, the survival at `t = 4` rises by
  1.27E-13: the Double below 1999 is on the route of before, off by 1.26E-13
  there (7.6E-14 at `t = 8`). That is the log-gamma factor of #175.

## Amendment 1

One registered holdout value of t clashed with a number reserved by the #307
holdout guard and was replaced by 13.6 after the first wave; the registration
records it. The 24 replaced rows had been exported and were within G1 (worst
1.87E-14). The 24 new rows were exported by run 45625806466 and are within
1.69E-14. Nothing else changed between the two waves: no other observation
of the grid differs.

## Not done: the cost per call

The registration said the cost per call would be measured in Excel with
`benchmark/kernel_profile`. That tool profiles the Gamma kernel only and does
not time a Student-t call, which I had not checked when I wrote it. No cost
was measured. The rule replaces one continued fraction in Doubles by the pair
fraction of #35 on the rows it serves; what that costs per call is not
known, and the criterion of #365 that asks for it is open.

## What remains

- `df < 1999`: up to 2.5E-13, the factor through log-gamma (#175).
- `t > 16`: up to 8.2E-13, the final exponential.
- Both are within the 1.3E-12 guarantee. `numerical_limitations.csv` records
  them with these figures.
