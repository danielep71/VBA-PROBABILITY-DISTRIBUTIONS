# #213: the Student-t tail above |t| = 8 at large df — result, 2026-11-14

Registration `6077914`, candidate `9c4821c`. Excel run 43637035288 (Excel
16.0 build 20430, 64-bit), evidence `36a17ad`.

**Verdict: ACCEPTED.** `above8_study.py --check` finds F1 to F4. The Excel
regression passed 1,952 of 1,952 assertions. The main grid and the
consolidated holdout came back with no observation changed, and the strict
accuracy gate passed with no contract changed.

## Forward

Relative error of the survival at 8 + 1 ulp, against the 80-digit
reference:

| df | Before (run 42281403598) | After |
| ---: | ---: | ---: |
| 1E5 | 4.33E-13 | at most 3.4E-15 |
| 1E6 | 4.0E-12 | at most 3.4E-15 |
| 1E7 | 8.1E-11 | at most 3.4E-15 |
| 1E8 | 7.78E-10 | at most 3.4E-15 |

The 24 probes: at most 1.84E-14 for `t` up to 14, which gates; 9.66E-14 at
`t = 20` and `t = 30`, reported. The tail there is below 1E-88 and the error
is that of exponentiating a logarithm near 700; it is inside the 1E-13
guarantee by a small margin and I do not claim more than was measured.

## Monotonicity

Across `|t| = 8` the survival does not rise and the cumulative does not
fall, at all four df and both signs. Before, they broke by 8.1E-11 at
df = 1E7. `seam_results.json` is regenerated from the new export: the
largest remaining break is 3.16E-12 across `df = 1E5` at `t = 2`, unchanged
and outside this rule.

## Inverse

No line of the inverse changed. With the corrected tail:

| Rows | Worst relative error |
| --- | ---: |
| Seam round trips, `df >= 1E5` (was 1.29E-12 at df = 1E7) | 1.11E-15 |
| `above8` inverses, t = 8.25 and 10 | 8.9E-16 |

None misses the 5E-15 measured target.

## Predictions

| # | Prediction | Outcome |
| --- | --- | --- |
| 1 | F1 to F4 hold | Holds |
| 2 | Worst gated forward error below 2E-14 | Holds: 1.84E-14 |
| 3 | Every seam round trip at `df >= 1E5` within 5E-15 | Holds: 1.11E-15 |
| 4 | No observation with `\|t\| <= 8` or `df < 1E5` changes | Holds: main grid and holdout unchanged; in `tchi_large_df` 45 rows changed, the 32 new ones and 13 seam rows at or beyond 8 + 1 ulp or inverses near it |
| 5 | The remaining monotonicity breaks are those across `df = 1E5` | Holds |

## The mechanism was not the one first named

#213 attributed the error to the continued fraction "overestimating" the
tail, and proposed bracketing the inverse at 8. The prefactor and the
fraction's recurrence are sound. The fraction was given `x = df / (df + t^2)`
as one Double, and its first denominator, `1 - (a + b) x / (a + 1)`,
cancels: the rounding of `x` is multiplied by about `a / (1/2 + a y)`. Giving
it `1 - y` as an exact pair removes the error without touching the inverse.

## What this does not show

- The same conditioning below `df = 1E5`, where the continued fraction
  serves every `t`: up to 3.16E-12 just below 1E5 at `t = 2`. Outside #213.
- The model's points are the registered rows, so the model was not tested
  on untouched points; Excel was.
- `t` beyond 30, and df between the four values measured.
