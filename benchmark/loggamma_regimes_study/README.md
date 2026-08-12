# PROB_LogGamma regime study

Evidence for replacing the single global `PROB_LogGamma` accuracy claim with
regime-aware contracts, and for validating the Phase 1 small-positive reroute
(ICR-P1-01 prerequisite chain, issue #12).

## Why this study exists

The module has advertised *"relative error below 6.1E-14 across Z in
[1E-8, 1E+50]"*. That single global claim is wrong in two independent ways.

**Subfinding A — a genuine subnormal defect.** The reflection path forms
`Sin(PROB_PI * Z)` after `PROB_PI * Z` has entered the subnormal range and lost
significand bits. `Log(Sin(...))` is then off by that relative error
*absolutely*, on a result of magnitude ~700 — about 4.6E-02 of absolute log
error at the smallest positive Double, a measured 6.2E-05 relative. Corrected by
the `PROB_LogGamma1pSeries(Z) - Log(Z)` branch.

**Subfinding B — the metric is ill-conditioned.** `Log(Gamma(Z))` is zero at
`Z = 1` and `Z = 2`. A global *relative* contract on a function that crosses
zero cannot hold: at `Z = 1.75` a measured 9.3E-14 relative is only about
7.9E-15 absolute on a value of magnitude 0.084. The relative column diverges
because the denominator does, not because the kernel degrades.

## Metric

Absolute error in the logarithm is primary everywhere except `general`, because
that is what downstream callers propagate: the relative error of `Exp(v)` is
approximately the absolute error of `v`. The two exact zeros are absolute-only —
a relative error against zero is not a number.

## Regimes

| regime | condition | metric | contract |
|---|---|---|---|
| `small_positive` | `Z <= 0.25` — the branch Phase 1 replaces | absolute | `LogGamma.small_positive.log_abs` |
| `reflection` | `0.25 < Z < 0.5` — unchanged branch | absolute | must not move |
| `near_zero` | neighbourhoods of `Z = 1` and `Z = 2` | absolute | `LogGamma.near_zero.log_abs` |
| `exact_zero` | `Z = 1`, `Z = 2` | absolute **only** | folded into `near_zero` |
| `general` | elsewhere | relative | `LogGamma.general.output_rel` |

## Grid

71 points × 2 quantities = 142 rows.

- **Subnormal band** where the reflection route actually degrades: 1E-308 down
  to the smallest positive subnormal.
- **Binary landmarks** `2^-1022`, `2^-1030`, `2^-1040`, `2^-1050`, `2^-1060`,
  `2^-1070`, `2^-1074`, with a `bits` column, so the relationship between
  retained significand bits and error is legible directly.
- **Arguments the committed Beta/F rows actually reach**: 9.9E-14, 2.7E-12,
  3E-11, 0.02, 0.04, 0.05, 0.1, 0.125, 0.2, 0.24.
- **Seam and boundary neighbours**: NextDown/NextUp of 0.25 and of 0.5.
- **Zero neighbourhoods**: NextDown/NextUp of 1 and 2, plus 0.9, 0.99, 1.01,
  1.25, 1.5, 1.75, 1.99, 2.01.
- **General regime**: 2.5 through 1E+50.

`EchoZ` round-trips each argument through `Val()`. With eight subnormal literals
in the grid, a parser that mis-reads one would make every other number in that
row measure an argument nobody asked for.

## Run it twice

**No committed observation covers `Z` below 1E-8**, so the subnormal improvement
has nothing to be measured against unless a pre-change baseline exists.

1. With the **current** `SPECIALFUNCS`, import `M_STATS_PROBDIST_LGREGIMES.bas`,
   `Debug > Compile`, run `Export_LogGammaRegimes`.
2. Copy the filled grid to `loggamma_regimes_baseline.csv`.
3. `python3 generate_loggamma_regimes.py` to restore an empty grid.
4. Import the **Phase 1** `SPECIALFUNCS`, `Debug > Compile`, run the export again.
5. `python3 analyze_loggamma_regimes.py --baseline loggamma_regimes_baseline.csv`.

## Files

| File | Role |
|---|---|
| `generate_loggamma_regimes.py` | Writes `loggamma_regimes_grid.csv` (142 rows, adaptive-precision refs). |
| `loggamma_regimes_grid.csv` | The grid; `arg1 = Z`. |
| `loggamma_regimes_baseline.csv` | The same grid exported before the Phase 1 edit. |
| `M_STATS_PROBDIST_LGREGIMES.bas` | Standalone export macro `Export_LogGammaRegimes` (dep: `PROB_LogGamma`). |
| `analyze_loggamma_regimes.py` | Argument integrity, accuracy by regime, classified baseline diff, both subfindings, contract recommendation. |

## Expected result (Python mirror of the emitted VBA; replace with real numbers)

Classified baseline diff: **29 expected improvement, 1 expected neutral
rounding, 0 unexpected, 41 unchanged**, and **no point outside `Z <= 0.25`
moves at all**. That last check is the structural one — Phase 1 edits only the
small-positive branch, so a moved point anywhere else is a defect, not a
consequence.

Largest improvements, absolute log error:

| Z | before | after | factor |
|---|---|---|---|
| 4.94E-324 | 4.61e-02 | 3.49e-14 | 1.3E+12 |
| 1E-322 | 2.67e-03 | 6.97e-14 | 3.8E+10 |
| 1E-320 | 6.55e-05 | 1.98e-14 | 3.3E+09 |
| 1E-318 | 5.72e-07 | 6.18e-14 | 9.3E+06 |
| 1E-315 | 1.49e-10 | 6.60e-14 | 2.3E+03 |

Provisional contracts:

| contract | threshold | worst | headroom |
|---|---|---|---|
| `LogGamma.small_positive.log_abs` | 2E-13 | 8.57e-14 | 2.3× |
| `LogGamma.near_zero.log_abs` | 5E-14 | 1.05e-14 | 4.8× |
| `LogGamma.general.output_rel` | 5E-13 | 1.02e-13 | 4.9× |

The small-positive floor is `EPS * Abs(Log(Z))` and is irreducible: 1.65E-13 at
the smallest positive subnormal, 5.1E-14 at 1E-100, 3.1E-16 at `Z = 0.25`. After
Phase 1 the absolute error there is dominated by `Log(Z)`, not by the series —
which is why the threshold is looser than the near-zero one despite the branch
being far more accurate in relative terms (1.2E-16 worst).
