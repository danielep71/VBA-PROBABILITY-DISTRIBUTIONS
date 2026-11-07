# Large-df Student t / Chi-square characterization study (P2-02)

## Why this exists

The public `StudentT_*` and `ChiSquare_*` surfaces **accept** degrees of freedom
up to the `1E100` representational bound but are only asserted "validated to
roughly `1E9`". That leaves a third state alongside *invalid input* and
*accuracy-contracted input*: **valid, computed, but not accuracy-contracted** —
where a worksheet user can receive a finite value the project makes no accuracy
claim about, with no machine-readable signal. Both surfaces now reject df above
1E8; see [Outcome](#outcome-2026-11-07).

This study measures that regime so any enforced limit (or capability claim)
rests on committed evidence rather than an assertion — the same discipline the
`f_envelope` study applied to F.

## What the regime actually looks like

By the CLT the standardized t and Chi-square both converge to the standard
normal, with `|t_cdf(x, df) - Phi(x)| ~ C(x)/df`. Consequences:

- **Student t is genuinely non-normal only up to df ~ 1E15.** Beyond that it
  equals `Phi(x)` to better than Double epsilon (`2^-53`), so the *correct
  Double answer there is the normal limit*. The question in that band is not
  "is the kernel accurate to 1E100" but **"does the kernel return the normal
  limit, or silently diverge from it?"**
- **The incomplete-beta / incomplete-gamma forms are ill-conditioned at large
  df** — fixed-precision references there are wrong. This study's references are
  self-checked (below), and any row that cannot be made solid is dropped, never
  shipped.

## References (self-checked)

- **Student t** — the beta path evaluated at escalating precision
  (80 -> 1280 digits) until two settings agree to `1E-40`. Stable and correct
  across `1E5 .. 1E100`.
- **Chi-square** — referenced only where a Wilson-Hilferty limit is *provably
  exact to Double precision*: `df >= 1E16` (WH error `~9E-3/df`, validated at
  build time against the exact incomplete gamma at low df). Chi-square is also
  bounded ABOVE at `df ~ 1E30`: beyond `~1E31` the entire non-trivial transition
  band (width `~sqrt(2 df)` around `df`) is narrower than one Double ULP of `df`,
  so no Double-representable abscissa yields a determinate Chi-square CDF - the
  accepted region there is degenerate rather than measurable, which is itself a
  finding. Within the measured band the CDF is very steep, so each abscissa is
  rounded to the exact Double the VBA will receive and the reference is taken at
  that same Double (otherwise a half-ULP mismatch would masquerade as a ~3%
  kernel error). The Chi-square **central band `1E5 .. 1E16` is reference-limited
  and intentionally omitted** - it needs a Temme uniform-asymptotic reference (a
  scoped follow-up), and shipping an under-precision reference would violate the
  project's self-check-before-ship rule.

`generate_tchi_large_df.py` asserts the WH validation and drops any
non-self-consistent row before writing the grid.

**Amendment, 2026-11-05 (#159).** Two corrections to the references above:

- **Survival precision.** The generator never set `mp.mp.dps`. Each CDF was
  computed under its own working precision, but the survival was then formed as
  `1 - CDF` at mpmath's default 53 bits, so all 75 survival references were
  accurate to only about 16 digits. The module now sets its precision
  explicitly, and each survival is computed directly inside its self-checked
  evaluation: `I_z(df/2, 1/2) / 2` for Student t.
- **Chi-square cumulative and survival.** Wilson-Hilferty is exact to Double
  precision at `df >= 1E16`, but only to 16-29 significant digits, below the
  25-digit reference floor. These rows now use Temme's uniform asymptotic
  expansion (DLMF 8.12, terms `c0` and `c1`), which gives each tail directly and
  is exact to better than `1E-38` relative at `df >= 1E16`; the generator
  validates it against the exact incomplete gamma at `df = 2E8`. The inverse
  rows keep the Wilson-Hilferty quantile.

The committed grid was re-referenced with
`python generate_tchi_large_df.py --merge-references --write`, which changes
reference cells only: 89 references moved, by at most 1.6E-16 relative, and no
observation changed. The old values are in `../study_reference_rebase.csv`.

## Running it

The grid is committed with its Excel observations. **Never run the bare
`python generate_tchi_large_df.py`:** it rewrites the grid from scratch and
drops every observation (and, before #218, the 130 seam rows). To re-reference
the committed grid, keeping every observation:

```
python generate_tchi_large_df.py --merge-references          # report only
python generate_tchi_large_df.py --merge-references --write  # patch reference cells
python analyze_tchi_large_df.py                              # worst relative error per (function, df)
```

The generator builds the seam rows from
`../student_t_large_df_study/seam_study.py`, which owns them, so
`--merge-references` covers the whole grid (#218). `test_tchi_merge.py` checks
this. Before #218 the generator did not produce the seam rows, and the merge
refused the grid.

**Observations.** The grid is exported unattended: the Excel VBA Regression
workflow's `export_tchi_large_df` input runs `ci/Export-ExcelGrids.ps1` after a
green regression, and `benchmark/apply_excel_export.py --grid tchi_large_df`
applies the artifact. The study module `M_STATS_PROBDIST_TCHI.bas` is the
exporter.

**Seam rows.** 130 rows with regime `seam` (52 survival, 52 cumulative, 26
inverse) measure the Student-t large-df seams of #210. They were appended by
`seam_study.py --append`; see
[`SEAM_STUDY.md`](../student_t_large_df_study/SEAM_STUDY.md).

## How to read the result

For each `(function, df)` the analyzer reports the worst relative error of the
observed VBA value against the true reference:

- errors that stay near Double epsilon mean the kernel **tracks the true
  (normal-limit) value** — accepted-but-uncontracted is then *safe*, and the
  right response is to extend the contracts and/or expose a capability indicator,
  not to impose a limit;
- a jump to `O(1E-6)+` marks a **silent-divergence boundary** — that is the
  committed evidence needed to justify enforcing a measured envelope there,
  exactly as `PROB_F_ValidateEnvelope` does for F.

Either way the outcome replaces the asserted "~1E9" with a measured boundary,
and no limit is introduced without evidence.

## Outcome (2026-11-07)

The `1E100` acceptance described above no longer holds. The measured envelope
(`benchmark/envelope_probe`) set `PROB_T_MAX_DF = PROB_CHI_MAX_DF = 1E8` in
`src/M_STATS_PROBDIST_TFAMILY.bas`, and the public surfaces reject larger df
with `CVErr(xlErrNum)`. Every Student-t row and every Chi-square row above
df 1E8 in this grid (146 rows) now observes `ERROR`. The rows at or below 1E8
are the Student-t rows at df 1E5..1E8, including the seam rows.
`numerical_limitations.csv` records the limit as
`IncompleteBeta.StudentTLargeDF`.
