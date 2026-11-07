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
- **Later additions**: tiny-shape Beta/F (`tiny_unbalanced` and
  `tiny_unbalanced_representable`); Chi-square and F densities;
  `Normal_IntervalProbability`; `Lognormal_ParametersFromMeanStdDev`; and the
  `split_boundary` and `deep_tail` regimes of the standard normal, normal and
  lognormal inverses.
- **Student t, large df (#34)**: 100 `envelope_domain` rows (survival,
  cumulative and quantile) at df between 1E5 and 1E8, appended once by
  `generate_holdout.py --append-student-t-large-df` and
  `--append-student-t-inverse` under the two #34 preregistrations in
  `../student_t_large_df_study/`.

659 points; references are mpmath / continued-fraction incomplete beta /
incomplete gamma at 50 digits.

## Files

| File | Role |
|---|---|
| `generate_holdout.py` | Builds the holdout references. **A bare run overwrites `holdout_grid.csv` and its Excel observations**; use `--out <other file>` to compare. The two `--append-student-t-*` flags append the #34 rows once and refuse a grid that already holds them. |
| `../_ibeta.py` | Continued-fraction incomplete beta, shared by the whole benchmark (single copy). |
| `holdout_grid.csv` | Fresh grid; `arg4` schema (Hypergeometric needs four parameters) plus the `expected_error` envelope marker. |
| `holdout_manifest.json` | Binds the grid's bytes and row count to the production modules and the exporter, with the export commit and Excel environment. |
| `M_STATS_PROBDIST_HOLDOUT.bas` | Export macro `Export_Holdout` for every function in the grid: Beta/F, Student t, `PROB_LogBeta`, the Chi-square and F densities, `Normal_IntervalProbability`, `Lognormal_ParametersFromMeanStdDev`, the normal and lognormal inverses, and the six discrete families. |
| `analyze_holdout.py` | Per-contract holdout worst vs frozen threshold + margin; writes `holdout_summary.md`. |
| `holdout_summary.md` | The generated verdict table. |
| `test_analyze_holdout.py` | Fixtures for the analyzer's metric, envelope and tail-dispatch semantics. |

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
3. Finalize and bind the export as in
   [`../PROVENANCE.md`](../PROVENANCE.md#fresh-export-operating-procedure), then
   commit the filled CSV with its manifest.
4. `python3 analyze_holdout.py` (reads `../accuracy_contracts.csv`).

Steps 1 to 3 can also run unattended: dispatch the Excel VBA Regression
workflow with `export_holdout`, then apply its artifact with
`../apply_excel_export.py --grid holdout`
([Automated fresh export](../../docs/EXCEL_CERTIFICATION.md#automated-fresh-export)).

## Freeze decision

- **All provisional contracts pass on the holdout** -> flip their provenance to
  `validated and frozen` in `accuracy_contracts.csv`.
- **Any contract exceeds its threshold** -> do not freeze; adjust that single
  threshold to the honest holdout-inclusive worst and record why.

## Role in the gate

The Accuracy Gate requires this holdout. `compute_errors.py` fails closed unless
`holdout_manifest.json` binds the committed grid to the current source, and the
gate then reruns `analyze_holdout.py` and fails when `holdout_summary.md`
differs from the rerun (see [`../PROVENANCE.md`](../PROVENANCE.md#what-the-hosted-gate-enforces)).

*Superseded, 2026-11-07.* An earlier version of this section said the F rows at
df 200000 and 500000 carry `ERROR`. That was true while the F envelope was 1E5.
Since `PROB_F_MAX_DF` rose to 1E10 they return values, and they are scored
under `F_*.validated`; the grid has no `ERROR` row.
