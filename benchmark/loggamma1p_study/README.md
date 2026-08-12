# LogGamma1p study (`PROB_TryLogGamma1p`)

Validates the `PROB_TryLogGamma1p` kernel and fixes the series/Lanczos seam
`PROB_LG1P_SERIES_MAX` from measured VBA data. Prerequisite for ICR-P1-01
(issue #12; parent #11, forward #13, inverse #14).

## Why the kernel exists

`PROB_LogGamma(1# + X)` is unusable for small `X`. In binary64 `1# + X` rounds to
exactly 1 for every `X` below `2^-53` (1.11E-16), so the leading term
`-EulerGamma * X` is lost. The result does not merely lose precision — it
returns a positive constant, the rounding residue of the addition, where the
true value is negative and vanishing:

Measured in the workbook, as scaled absolute error:

| X | kernel | `PROB_LogGamma(1 + X)` | ratio |
|---|---|---|---|
| 1E-300 | 4.20E-17 | 3.77E+285 | 9.0E+301 |
| 1E-20 | 2.06E-17 | 3.77E+05 | 1.8E+22 |
| 2^-53 | 1.31E-17 | 3.46E+01 | 2.6E+18 |
| 1E-14 | 2.97E-17 | 4.59E-01 | 1.6E+16 |
| 1E-10 | 1.56E-17 | 4.34E-05 | 2.8E+12 |
| 1E-06 | 1.16E-16 | 3.89E-09 | 3.4E+07 |
| 0.01 | 2.44E-18 | 4.97E-13 | 2.0E+05 |
| 0.2 | 2.72E-17 | 2.85E-14 | 1.1E+03 |

At and below `2^-53` the naive spelling does not merely lose precision: it
returns the rounding residue of the addition, a positive constant of about
3.77E-15, where the true value is negative and vanishing.

This is not a subnormal-shape problem: absorption begins roughly 292 decimal
orders of magnitude above the smallest normal Double, and degradation is gradual
well before it.

## Metric note

The contract metric is the **scaled** absolute error,

```
Abs(observed - reference) / X
```

not ordinary absolute or relative error. The scaled Gamma inverse computes
`[LogProbability + LogGamma1p(Shape)] / Shape`, so an absolute error here is
divided by `Shape` before it reaches the quantile: **the scaled error is the
relative error of the quantile it produces.** Ordinary absolute error would look
flattering at small `X` and would prove nothing about that caller.

## What it measures

Four quantities per point:

- `EchoX` — `X` itself, round-tripped through `Val()`. The grid carries
  subnormal literals down to the smallest positive Double; if the parser
  mis-reads one, every other number in that row measures an argument nobody
  asked for. Making the parsed Double an observation rules that out first.
- `LogGamma1p` — `PROB_TryLogGamma1p(X)`, the kernel.
- `LogGamma1pOverX` — `PROB_TryLogGamma1p(X) / X`, divided **in VBA**, so the
  rounding of the division the inverse actually performs is captured.
- `LogGammaNaive` — `PROB_LogGamma(1# + X)`, the spelling being replaced.

## Regimes

Regimes name *branches*, not intervals of interest, so a contract is never
averaged across two different algorithms.

One evidence regime per contract bucket, so a contract is never averaged across
two algorithms or two error mechanisms.

| regime | condition | contract |
|---|---|---|
| `small` | `3.855E-308 <= X <= 0.25`, series branch | `LogGamma1p.small.scaled_abs`, provisional 5E-16 |
| `series_seam` | one-ulp neighbourhood of `0.25` | `LogGamma1p.series_seam`, characterization |
| `lanczos_handover` | `X > 0.25` | `PROB_LogGamma`'s own contract |
| `subnormal_result` | `X < 3.855E-308` | `LogGamma1p.subnormal_result`, limitation |

The seam neighbourhood is separated because the point one ulp **above** `0.25`
runs the Lanczos route and inherits that kernel's error, roughly two orders of
magnitude above the series contract. Pooling it would either loosen the series
contract or fail the seam for no reason.

The `subnormal_result` boundary is `PROB_MIN_NORMAL / EulerGamma`
(3.854839696505424E-308) — the point where the leading term `EulerGamma * X`
itself becomes subnormal. That boundary is **mathematical, not empirical**:
measured scaled error stays inside the contract somewhat below it, but the
quantization the regime names begins exactly there.

## Grid

50 points x 4 quantities = 200 rows. Decades from the smallest positive
subnormal to 10, the `1 + X` absorption boundary and its approach
(1E-17 … 1E-10, including `2^-53` exactly), the series interior, the seam at
`0.25 ± 1 ulp`, and the hand-over.

References are mpmath at `dps = 60 + |log10(X)| + 10`. The precision **must**
rise with the magnitude of `X`: `mp.loggamma(1 + mpf(x))` silently returns zero
once `1 + x` rounds to 1 at the working precision, reproducing in the reference
the exact defect the kernel exists to remove. At the adaptive precision the
result agrees with the Maclaurin series to every digit carried, at every point
on this grid.

## Files

| File | Role |
|---|---|
| `generate_loggamma1p.py` | Writes `loggamma1p_grid.csv` (200 rows, adaptive-precision refs). |
| `loggamma1p_grid.csv` | The grid; `arg1 = X`. |
| `M_STATS_PROBDIST_LOGGAMMA1P.bas` | Standalone export macro `Export_LogGamma1p` (deps: `PROB_TryLogGamma1p`, `PROB_LogGamma`). |
| `analyze_loggamma1p.py` | Argument integrity, kernel validation by regime, head-to-head against the naive spelling, seam continuity, subnormal-result limitation, threshold recommendation. |

## How to run

1. Import the corrected `src/M_STATS_PROBDIST_SPECIALFUNCS.bas` and
   `M_STATS_PROBDIST_LOGGAMMA1P.bas` into the workbook, then `Debug > Compile`.
2. Run `Export_LogGamma1p`; select `loggamma1p_grid.csv` when prompted.
3. Commit the filled CSV.
4. Analysis: `python3 analyze_loggamma1p.py`.

## Measured result (real VBA)

- Small-series worst scaled error **1.36E-16**, at `X = 1E-13`. This is the
  arithmetic floor of the final `Acc * X`, about two ulps — not a coefficient
  floor, so widening any further coefficient would buy nothing. A 30,000-point
  Python dense sweep of the same interval gives 1.18E-16, consistent.
- The `C1` split expression survived the VBE intact, proven by measurement
  rather than by inspection: scaled error reaches 9.57E-19 at `X = 1E-100` and
  4.68E-18 at `X = 1E-200`, both far below the ~1E-16 floor a one-ulp `C1`
  error would impose.
- The `small` and `subnormal_result` regimes reproduce the Python mirror of the
  emitted code **exactly**, digit for digit. The `series_seam` and
  `lanczos_handover` regimes differ by 1.05x and 1.22x, because those delegate
  to `PROB_LogGamma`, whose VBA `Log` differs from Python's at the ulp level.
  Series arithmetic is bit-reproducible across platforms; the Lanczos route is
  not, at the 1E-14 level.
- Against the naive spelling the improvement runs from 1.1E+03 at `X = 0.2` to
  9.0E+301 at `X = 1E-300`.
- Seam step at `0.25` is **5.68E-14**, which is `PROB_LogGamma`'s own error, not
  a defect in the series. It shrinks only if that kernel improves.
- Subnormal-result degradation sits at or below the half-ulp closed form
  `0.5 * minimum_subnormal / X` = `2^-1075 / X` at every measured point, with a
  worst gap of 11.7x at `X = 1E-315`. The closed form is an upper **bound**, not
  a prediction: where the quantization happens to land favourably the measured
  error is smaller. That the measurement never exceeds it is the evidence that
  this is a representability limit of the output rather than an algorithm
  defect.

Recommended contract: `LogGamma1p.small.scaled_abs <= 5E-16` (4.2x headroom).
Provisional until the independent holdout is populated.

## Open observation, out of scope here

`PROB_LogGamma` publishes "relative error below 6.1E-14 across Z in
[1E-8, 1E+50]". That does not hold near the zeros of `Log(Gamma)` at `Z = 1` and
`Z = 2`: relative error reaches 2.3E-12 at `Z = 1.99` and is unbounded at the
zeros themselves. Absolute error stays at or below 1.5E-14 throughout, which is
the meaningful bound. The claim needs restating as absolute, or domain-excluding
a neighbourhood of the zeros. Tracked separately from ICR-P1-01.
