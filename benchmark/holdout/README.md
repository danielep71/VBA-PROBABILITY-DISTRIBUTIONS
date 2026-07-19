# Independent holdout

Validates the regime-specific accuracy contracts on data that was **not** used to
set any threshold. If the provisional thresholds hold here, they generalise and
can be frozen (`measured provisional` -> `validated and frozen`).

## Fresh points (none in any fitting set)

- **Beta unbalanced shapes**: (0.55,3000), (1.9,50000), (3.3,40000), (0.42,700),
  (7.7,200), (250,1.15), (0.9,250000), (4.4,90000) — new non-integer shapes.
- **F df (validated range, param < ~1E7)**: (3,5000), (1.5,200000), (7,50000),
  (500000,4).
- **Probabilities**: 0.0001, 0.005, 0.25, 0.75, 0.995, 0.9999 (extra tails).
- **PROB_LogBeta**: fresh Small 0.42/1.9/3.3/7.7 x near-seam and between-decade
  ratios 0.3, 0.15, 0.11, 0.101, 0.099, 0.09, 0.075, 0.03, 0.003, 3E-4, 3E-6, 3E-9
  (stresses the 0.1 crossover from both sides).

134 points; references are mpmath / continued-fraction incomplete beta at 50 digits.

## Files

| File | Role |
|---|---|
| `generate_holdout.py` | Writes `holdout_grid.csv` (references). |
| `_ibeta.py` | Continued-fraction incomplete beta (shared with the inverse study). |
| `holdout_grid.csv` | Fresh grid; main-grid schema. |
| `holdout.bas` | Export macro `Export_Holdout` (7 public functions + `PROB_LogBeta`). |
| `analyze_holdout.py` | Per-contract holdout worst vs frozen threshold + margin. |

## How to run

1. Import `holdout.bas`, `Debug > Compile`.
2. Run `Export_Holdout`; select `holdout_grid.csv`.
3. Commit the filled CSV.
4. `python3 analyze_holdout.py` (reads `../accuracy_contracts.csv`).

## Freeze decision

- **All provisional contracts pass on the holdout** -> flip their provenance to
  `validated and frozen` in `accuracy_contracts.csv`.
- **Any contract exceeds its threshold** -> do not freeze; adjust that single
  threshold to the honest holdout-inclusive worst and record why.
