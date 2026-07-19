# Delta seam study (PROB_LogGammaDelta + LogBeta crossover)

Validates the new `PROB_LogGammaDelta` kernel and selects the two-regime
`PROB_LogBeta` crossover (`PROB_LOGBETA_STABLE_RATIO`) from measured VBA data —
implementing steps 2 and 4 of the LogBeta correction plan.

## What it measures

For each `(Large, Small)` point, three quantities are exported and compared to
50+ digit mpmath references:

- `LogGammaDelta`  — `PROB_LogGammaDelta(Large, Small)` (validates the kernel);
- `LogBeta_ident`  — `Log(Beta)` via the direct three-log-gamma identity;
- `LogBeta_stable` — `Log(Beta)` via `LogGamma(Small) - PROB_LogGammaDelta`.

The two LogBeta rows share the reference; measuring each *route* at every ratio
is what lets the crossover be chosen from evidence rather than theory.

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
4. Analysis (done for you): `python3 analyze_delta_seam.py`.

## Reading the result

- **Delta validation** is reported for the *production regime* (ratio < 0.1,
  where the delta is actually used) separately from the full grid. Points at
  ratio >= 0.1 use the identity, so a larger delta error there is expected and
  harmless.
- **Crossover**: the analyzer prints, per ratio, the worst identity error and
  worst stable error and recommends `PROB_LOGBETA_STABLE_RATIO` from the clean
  overlap where both routes are within 5E-15. A Python prototype places this near
  0.1; the VBA measurement is the authority.

The crossover currently coded (`0.1`) is provisional. Adjust the constant in
`M_STATS_PROBDIST_SPECIALFUNCS` only if the measured VBA envelope indicates a
different value.
