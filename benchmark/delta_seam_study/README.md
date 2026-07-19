# Delta seam study (PROB_LogGammaDelta + LogBeta crossover)

Validates the `PROB_LogGammaDelta` kernel and characterises the two-regime
`PROB_LogBeta` crossover (`PROB_LOGBETA_STABLE_RATIO`) from measured VBA data.
Implements steps 2 and 4 of the LogBeta correction.

## What it measures

For each `(Large, Small)` point, three quantities are exported and compared to
120-digit mpmath references:

- `LogGammaDelta`  — `PROB_LogGammaDelta(Large, Small)` (validates the kernel);
- `LogBeta_ident`  — `Log(Beta)` via the direct three-log-gamma identity;
- `LogBeta_stable` — `Log(Beta)` via `LogGamma(Small) - PROB_LogGammaDelta`.

The two LogBeta rows share the reference; measuring each *route* at every ratio
is what lets the crossover be chosen from evidence.

## Metric note

`PROB_LogBeta` is assessed by ABSOLUTE error, because downstream code computes
`exp(-LogBeta)` and to first order `|Δexp/exp| ≈ |ΔLogBeta|`. A large relative
error on a tiny `|LogBeta|` is harmless; a large absolute error on a big
`|LogBeta|` is what propagates. The public Beta/F functions are assessed by their
own relative error in `beta_f_unbalanced/`.

## Grid

- **Small**: 0.25, 0.7, 1.3, 2.5, 5.75, 10.25 (non-integer cases included).
- **Seam ratios**: 0.5, 0.2, 0.15, 0.1, 0.08, 0.05, 0.03, 0.02, 0.01, 0.005.
- **Deep ratios**: 1E-3 down to 1E-18.
- **Absolute scales**: fixed Large in {1E2, 1E4, 1E8, 1E12, 1E20, 1E50}.

References are mpmath at 120 digits (needed so `loggamma(Large+Small) -
loggamma(Large)` stays accurate when the two large values nearly cancel).

## Files

| File | Role |
|---|---|
| `generate_delta_seam.py` | Writes `delta_seam_grid.csv` (396 rows, 120-digit refs). |
| `delta_seam_grid.csv` | The grid; `arg1 = Large`, `arg2 = Small`. |
| `delta_seam.bas` | Standalone export macro `Export_Delta_Seam` (deps: `PROB_LogGamma`, `PROB_LogGammaDelta`). |
| `analyze_delta_seam.py` | Delta validation (production regime vs full grid) + crossover envelope. |

## How to run

1. Import `delta_seam.bas` into the workbook and `Debug > Compile`.
2. Run `Export_Delta_Seam`; select `delta_seam_grid.csv` when prompted.
3. Commit the filled CSV.
4. Analysis: `python3 analyze_delta_seam.py`.

## Measured result (real VBA)

- `PROB_LogGammaDelta` in the production regime (ratio < 0.1): worst relative
  error ~4.75E-15. Points at ratio >= 0.1 use the identity, not the delta.
- The stable difference beats the identity everywhere unbalanced; the catastrophic
  cancellation of the direct identity is removed.
- Crossover: identity is better at ratio >= ~0.2, the stable difference is better
  below; the provisional `PROB_LOGBETA_STABLE_RATIO = 0.1` sits safely inside the
  overlap. The exact constant is confirmed here from measured data, not theory.

## Status

The stable-delta correction is shipped. Public Beta/F accuracy at unbalanced
arguments is measured separately in `beta_f_unbalanced/`; the resulting
regime-specific contracts are the authority for what the fix delivers.
