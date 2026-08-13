# Stirling overflow probe

Establishes whether the `PROB_StirlingError` small-N recurrence is defective,
and if so how. Two hypotheses were tested; one was rejected and one confirmed.

## Background

`PROB_StirlingError` handles `0 < N < 0.5` by recursing upward:

```
delta(N) = delta(N + 1) + (N + 0.5) * Log((N + 1) / N) - 1
```

The `(N + 1)` in that expression is the same `1 + x` shape that
`PROB_TryLogGamma1p` was built to remove from `PROB_LogGamma`, so the
recurrence was proposed as the next remediation in the same chain. The blast
radius looked wide: `PROB_StirlingError` is reached from 39 `K_STATS_*`
entry points across nine families.

## Hypothesis 1, precision - REJECTED

Two arrangements were compared against mpmath at 60 digits, holding
`delta(N + 1)` common so the comparison does not depend on which `LogGamma`
sits underneath:

```
A  (N + 0.5) * Log((N + 1) / N)          the current spelling
B  (N + 0.5) * (Log1p(N) - Log(N))       never forms 1 + N
```

Across 4,920 points spanning `(0, 0.5)`:

| | worst absolute error | at |
| --- | --- | --- |
| A | 4.828E-14 | N = 1E-244 |
| B | 4.828E-14 | N = 1E-244 |

B is strictly better at 442 points, strictly worse at 432, and **bit-identical
at 4,046** - 82 per cent of the sample. That is a coin flip, not a repair.

The `1 + x` analogy fails structurally. In `LogGamma(1 + x)` the entire answer
is `-gamma*x + O(x^2)`, so it lives precisely in the bits that forming `1 + x`
destroys. In `delta(N)` the answer is dominated by `-(N + 0.5) * Log(N)`,
computed from `N` directly; the increment contributes only the `Log1p(N) ~ N`
correction, bounded near one ulp of 1.0 and scaled by `(N + 0.5)`. There is no
precision case for changing the spelling.

## Hypothesis 2, reachability - CONFIRMED

Forming the ratio explicitly is a different defect. Once `1 / N` exceeds the
Double range the quotient is not representable and the expression faults before
`Log()` is called. The threshold is `1 / DoubleMax`, which equals
`2^-1024 * (1 + 2^-53)`; the excess over `2^-1024` is `2^-1077`, an eighth of a
subnormal ULP and below the representable grid, so `1 / DoubleMax` rounds to
exactly `2^-1024`.

The domain admits it. `PROB_CN_ValidateXShapeScale` delegates to
`PROB_IsPositiveWithinSupportedMagnitude`, which tests `X > 0` and an **upper**
bound of 1E100. There is no lower cutoff, normal or subnormal, so
`K_STATS_Gamma_Density(1, 1E-320, 1)` reaches the branch.

## Measured on real VBA

Excel 2016/365 64-bit, Italian locale, at `dfdbbbc`. Fourteen points; extreme
values built by halving and by adding exact multiples of the smallest positive
Double, so no result depends on literal parsing.

| point | N | current | density |
| --- | --- | --- | --- |
| `seam_last_safe` = 2^-1024 + 2^-1074 | 5.5626846462680084E-309 | OK | OK |
| `seam_first_fail` = 2^-1024 | 5.5626846462680035E-309 | **Error 6 Overflow** | **CVErr 2015** |

These are adjacent Doubles. Everything at or below `2^-1024` fails; everything
above works. Predicted boundary and measured boundary are the same Double.

**The failure is visible, not silent.** Error 6 propagates to `Err_Handler` and
surfaces as `CVErr 2015` (`#VALUE!`). At `X = 1, Scale = 1` the density is
`Exp(-1) / Gamma(N) ~ N / e` - small and finite, 3.68E-321 at `N = 1E-320` -
so a legitimate query receives a refusal rather than a wrong number. That
lowers the severity: this is a domain defect, not a silent-wrong-answer defect.

## Candidate behaviour

Both candidates are finite at all fourteen points including `2^-1074`:

```
B  StirlingError(N + 1) + (N + 0.5) * (Log1p(N) - Log(N)) - 1
C  LogGamma1p(N) - (N + 0.5) * Log(N) + N - HALF_LOG_TWO_PI
```

- B and C are **bit-identical at all thirteen points below 0.25**.
- B is **bit-identical to the current implementation at every point where the
  current implementation works**, except `top_of_branch`, where it differs by
  5.55E-17. B is a domain extension, not a numerical change.
- C is better only at `N = 0.25`: 2.14E-18 against B's 1.13E-14, on an error
  already three orders inside the frozen contract.
- Worst for both: **4.828E-14** at `N = 1E-244`, against `StirlingError.all.output`
  at 1E-13 absolute - a margin of 2.07x.

`p2_1023_anchor` reproduces the frozen holdout worst of 3.57E-14 exactly, an
independent check that the probe and the committed evidence agree.

## Conclusion

Adopt B. It removes the reciprocal with a kernel that already exists
(`PROB_Log1p`, which returns `X` unchanged once `1 + X` rounds to one), needs no
signature change, introduces no new branch, and does not disturb the recurrence
structure. C is the superior direct formulation near the top of the branch and
is documented here for that reason, but buying a second seam at 0.25 to improve
an already negligible error has no production justification.

## Margin note for #7

The independent point `N = 1E-244` measures 4.828E-14, worse than the frozen
holdout worst of 3.57E-14. The 1E-13 threshold still holds; the true margin is
2.07x rather than the 2.8x the holdout alone implies. Recorded here for the
margin audit; **no threshold is changed by this probe**.

## Scope

This is a characterization probe. It promotes no grid row, claims no threshold
and touches no registry. `promote_grid_rows.py` remains the only sanctioned
route for grid changes.

## Files

| file | purpose |
| --- | --- |
| `M_STATS_PROBDIST_STIRLINGPROBE.bas` | export macro; run `Probe_StirlingOverflow` |
| `stirling_overflow_probe.csv` | measured output, 14 points, 16 columns |
