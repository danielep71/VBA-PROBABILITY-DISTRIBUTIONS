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
5. `python3 generate_loggamma_regimes_holdout.py`
6. Still on Phase 1, run `Export_LogGammaRegimes` again and pick the **holdout**.
7. `python3 analyze_loggamma_regimes.py --baseline loggamma_regimes_baseline.csv --holdout loggamma_regimes_holdout.csv`

Steps 5-7 are needed only once, to freeze the contracts. The baseline export
(step 3) is the half that cannot be reproduced later; the holdout can always be
regenerated and re-exported.

## Files

| File | Role |
|---|---|
| `generate_loggamma_regimes.py` | Writes `loggamma_regimes_grid.csv` (142 rows, adaptive-precision refs). |
| `loggamma_regimes_grid.csv` | The grid; `arg1 = Z`. |
| `loggamma_regimes_baseline.csv` | The same grid exported before the Phase 1 edit. |
| `M_STATS_PROBDIST_LGREGIMES.bas` | Standalone export macro `Export_LogGammaRegimes` (dep: `PROB_LogGamma`). |
| `generate_loggamma_regimes_holdout.py` | Writes `loggamma_regimes_holdout.csv`; refuses to emit a point present in the fitting set. |
| `loggamma_regimes_holdout.csv` | Independent holdout, 62 points, disjoint by construction. |
| `analyze_loggamma_regimes.py` | Argument integrity, accuracy by regime, classified baseline diff, both subfindings, contract recommendation, holdout freeze decision. |

## Measured result (real VBA)

Classified baseline diff, Phase 0 export against Phase 1 export of the same
grid: **30 expected improvement, 1 expected neutral rounding, 0 unexpected,
40 unchanged**, and **no point outside `Z <= 0.25` moved at all**. That last
line is the structural check — Phase 1 edits only the small-positive branch, so
a moved point anywhere else would be a defect rather than a consequence.

Largest improvements, absolute log error:

| Z | bits | before | after | factor |
|---|---|---|---|---|
| 4.94E-324 | 1 | 4.61e-02 | 3.49e-14 | 1.32E+12 |
| 8E-323 | 5 | 5.30e-03 | 2.63e-14 | 2.01E+11 |
| 1E-322 | 5 | 2.67e-03 | 6.97e-14 | 3.83E+10 |
| 1E-320 | 11 | 6.55e-05 | 1.98e-14 | 3.30E+09 |
| 8.095E-320 | 15 | 2.84e-06 | 8.57e-14 | 3.31E+07 |
| 1E-318 | 18 | 5.72e-07 | 6.18e-14 | 9.26E+06 |
| 8.289046E-317 | 25 | 8.85e-09 | 2.96e-14 | 2.99E+05 |
| 1E-315 | 28 | 1.49e-10 | 6.60e-14 | 2.26E+03 |

Worst error by regime, after Phase 1:

| regime | n | worst absolute | worst relative |
|---|---|---|---|
| `small_positive` | 39 | 8.57e-14 @ 8.095E-320 | 1.2e-16 |
| `reflection` | 5 | 3.67e-15 | n/a |
| `near_zero` | 9 | 9.69e-15 @ 2.01 | n/a |
| `exact_zero` | 2 | 9.77e-15 @ 2.0 | undefined |
| `general` | 16 | n/a | 9.31e-14 @ 1E+50 |

Provisional contracts. The strict 1-2-5 rule gives 2E-13 / 2E-14 / 2E-13, but
all three land at only ~2x headroom, which is thin by this repository's
standard and likely to be exceeded by an independent holdout. One step looser
is the safer freeze:

| contract | strict 1-2-5 | headroom | recommended | headroom |
|---|---|---|---|---|
| `LogGamma.small_positive.log_abs` | 2E-13 | 2.3x | 5E-13 | 5.8x |
| `LogGamma.near_zero.log_abs` | 2E-14 | 2.0x | 5E-14 | 5.1x |
| `LogGamma.general.output_rel` | 2E-13 | 2.1x | 5E-13 | 5.4x |

The small-positive floor is `EPS * Abs(Log(Z))` and is irreducible: 1.65E-13 at
the smallest positive subnormal, 5.1E-14 at 1E-100, 3.1E-16 at `Z = 0.25`. After
Phase 1 the absolute error there is dominated by `Log(Z)`, not by the series —
which is why its threshold is the loosest of the three despite the branch being
by far the most accurate in relative terms.

The absolute column for `general` is meaningless and is reported only for
symmetry: at `Z = 1E+50`, `LogGamma(Z)` is about 1.1E+52, so 1.27E+36 of
absolute error is 9.3E-14 relative. `general` is the one regime a relative
contract fits.

## Independent holdout

The three contracts were set from this study's own grid, so that grid cannot
also validate them. `loggamma_regimes_holdout.csv` carries 62 points that set
none of the thresholds.

Disjointness is asserted, not eyeballed: the generator reads the fitting set and
refuses to emit any point it already contains. The points are also structurally
unlike it rather than merely different. The fitting set uses decade landmarks,
exact powers of two, and the arguments the committed Beta/F rows happen to
reach. A holdout built from neighbouring decades would test the same shape of
number, so this one uses irrational mantissas, odd binary exponents between the
fitting set's even landmarks, asymmetric offsets around both zeros, and eight
points inside the reflection interval that the fitting set samples only five
times. The subnormal points land on retained-bit counts of 2, 8, 14, 24, 30, 36,
42, 48 and 52 — none of which the fitting set measures.

| contract | threshold | holdout worst | points | margin | verdict |
|---|---|---|---|---|---|
| `LogGamma.small_positive.log_abs` | 5E-13 | 6.27e-14 @ 7.0103E-320 | 36 | 8.0x | PASS |
| `LogGamma.near_zero.log_abs` | 5E-14 | 1.13e-14 @ 2.0872 | 16 | 4.4x | PASS |
| `LogGamma.general.output_rel` | 5E-13 | 3.24e-14 @ 2.7183 | 10 | 15.4x | PASS |

This also settles the choice against the tighter 1-2-5 step. At 2E-14 the
near-zero contract would have held on the holdout at only 1.77x, below even its
2.05x on the fitting set — the looser threshold is validated by the holdout
rather than merely preferred.

The `small_positive` worst reproduces the Python mirror of the emitted code to
the digit, since that branch is pure series arithmetic. `near_zero` and
`general` drift by a few percent and the near-zero worst moves from Z = 1.9819
to Z = 2.0872, because both delegate to the Lanczos route where VBA's `Log`
differs from Python's at the ulp level. Same split seen in every earlier
checkpoint of this chain.
