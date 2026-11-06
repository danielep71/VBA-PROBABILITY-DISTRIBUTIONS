# Independent holdout

Validates the regime-specific accuracy contracts on data that was **not** used to
set any threshold. If the provisional thresholds hold here, they generalise and
can be frozen (`measured provisional` -> `validated and frozen`).

## Fresh points (none in any fitting set)

- **Beta unbalanced shapes**: (0.55,3000), (1.9,50000), (3.3,40000), (0.42,700),
  (7.7,200), (250,1.15), (0.9,250000), (4.4,90000) — new non-integer shapes.
- **F df (validated range, param < ~1E7)**: (3,5000), (1.5,200000), (7,50000),
  (500000,4).
- **PROB_LogBeta**: fresh Small 0.42/1.9/3.3/7.7 x near-seam and between-decade
  ratios (stresses the 0.1 crossover from both sides).
- **Discrete (Binomial / Poisson / Geometric)**: fresh parameters disjoint from
  the main grid — Binomial n in {50, 5000, 500000, 5000000}, Poisson mean in
  {10, 200, 100000}, Geometric p in {0.2, 0.01, 1e-4}.
- **Negative Binomial**: fresh (r, p) in {(2,0.3), (20,0.6), (200,0.4), (2000,0.75)}.
- **Hypergeometric**: fresh (n, K, N) in {(20,30,80), (60,300,500), (200,2000,50000)}.
- **Discrete Uniform**: fresh supports [2,13], [-20,-3] (fully negative), [100,4099],
  [-77777,22223]; inverse probabilities built as (j + 0.37) / n so no point lands
  exactly on a CDF step.
- **Probabilities**: 0.0001, 0.005, 0.25, 0.75, 0.995, 0.9999 for the continuous
  families; non-tie probabilities for the discrete inverses.

388 points; references are mpmath / continued-fraction incomplete beta /
incomplete gamma at 50 digits.

## Files

| File | Role |
|---|---|
| `generate_holdout.py` | Writes `holdout_grid.csv` (references). |
| `_ibeta.py` | Continued-fraction incomplete beta (shared with the inverse study). |
| `holdout_grid.csv` | Fresh grid; 12-column `arg4` schema (Hypergeometric needs four parameters). |
| `M_STATS_PROBDIST_HOLDOUT.bas` | Export macro `Export_Holdout` for all contracted public functions (Beta/F, `PROB_LogBeta`, and the six discrete families). |
| `analyze_holdout.py` | Per-contract holdout worst vs frozen threshold + margin. |

**Survival references (#159, 2026-11-05).** `generate_holdout.py` now computes
every Beta and F survival reference directly as an upper tail, at the exact
Double each written argument token parses to. Before, they were `1 - CDF` at the
unrounded design value, and 15 of the 72 committed survival references miss the
25-digit floor (worst 14.2 digits). `holdout_grid.csv` is bound byte for byte by
`holdout_manifest.json`, so those references can change only together with a
fresh Excel holdout export; until then `../survival_reference_debt.json`
registers them, and the fixed generator supplies the replacement values (every
survival reference it emits for the committed rows meets the floor, the worst at
29.6 digits).

*2026-11-06 (#81).* The #81 wave re-based these references with `../rebase_survival_debt.py`, and the Excel export rebound the grid; `../survival_reference_debt.json` is now empty.

## How to run

1. Import `M_STATS_PROBDIST_HOLDOUT.bas`, `Debug > Compile`.
2. Run `Export_Holdout`; select `holdout_grid.csv`.
3. Commit the filled CSV.
4. `python3 analyze_holdout.py` (reads `../accuracy_contracts.csv`).

## Freeze decision

- **All provisional contracts pass on the holdout** -> flip their provenance to
  `validated and frozen` in `accuracy_contracts.csv`.
- **Any contract exceeds its threshold** -> do not freeze; adjust that single
  threshold to the honest holdout-inclusive worst and record why.

## Known gap

`Export_Holdout` cannot compute F at the extreme validated-range df (200000,
500000); those rows carry `ERROR` and are covered instead by the F study
harnesses. They do not affect the strict gate, which runs on the main grid.
