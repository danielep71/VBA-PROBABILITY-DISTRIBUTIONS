# F accuracy-envelope study

Measures the exact df boundary where the public F functions (`F_Cumulative`,
`F_Survival`) cross their 1.1E-10 accuracy contract, so the extreme-df policy can
rest on a measured number rather than "approximately 1E7".

## Why

`F_CDF(x; d1, d2) = I_y(d1/2, d2/2)` with `y = d1*x / (d1*x + d2)`. The degrading
quantity is the incomplete-beta shape parameter `max(d1/2, d2/2)`. Beyond roughly
1E7 the continued fraction can satisfy its LOCAL increment test while the returned
value is off by up to ~4E-7 — a wrong answer returned without error
(`numerical_limitations.csv`, `IncompleteBeta.ExtremeShape`). This study finds
where that actually begins.

## Design

- **Sweep** the incomplete-beta shape parameter finely from 1E6 to 5E9 around the
  1E7 region (15 points).
- **Both orientations**: large second beta parameter (`d1=1, d2=2B`) and large
  first (`d1=2B, d2=1`) — the continued fraction can behave differently on each
  side, so both are measured.
- **Five x values** per point spanning the mass, for `F_Cumulative` and
  `F_Survival`. 300 rows total.
- **Reference**: 50-digit continued-fraction incomplete beta, validated
  self-consistent to ~1E-41 (50 vs 80 dps) across the whole region, so it is
  trustworthy ground truth exactly where the double-precision kernel is suspect.

## Files

The study has three sweeps. The strict envelope is the **minimum** boundary over
all of them, so all three matter.

| File | Role |
|---|---|
| `generate_f_envelope.py` | Writes `f_envelope_grid.csv`; base sweep, beta shape ~1E6 to 5E9. |
| `generate_f_envelope_gap.py` | Writes `f_envelope_gap_grid.csv`; fills the region below the base sweep, shape ~50 to 1E6. |
| `generate_f_envelope_bothlarge.py` | Writes `f_envelope_bothlarge_grid.csv`; both df large (`d1=d2` and `d2=3*d1`), which degrades earlier than one-large. |
| `f_envelope_grid.csv`, `f_envelope_gap_grid.csv`, `f_envelope_bothlarge_grid.csv` | The sweep grids; `arg1=x`, `arg2=d1`, `arg3=d2`. 11-column study schema, `observed_vba` at index 8. |
| `M_STATS_PROBDIST_FENV.bas` | Export macro `Export_FEnvelope` (base sweep). |
| `M_STATS_PROBDIST_FENVGAP.bas` | Export macro `Export_FEnvelopeGap` (gap sweep). |
| `M_STATS_PROBDIST_FENVBL.bas` | Export macro `Export_FEnvelopeBothLarge` (both-large sweep). |
| `analyze_f_envelope.py` | Degradation curve + measured envelope boundary. |

The reference helper is single-sourced at `benchmark/_ibeta.py`; the generators
import it from there rather than keeping a copy in this folder.

## How to run

1. Import `M_STATS_PROBDIST_FENV.bas`, `Debug > Compile`.
2. Run `Export_FEnvelope`; select `f_envelope_grid.csv`.
3. Commit the filled grid.
4. `python3 analyze_f_envelope.py`.

Repeat with `M_STATS_PROBDIST_FENVGAP.bas` / `Export_FEnvelopeGap` /
`f_envelope_gap_grid.csv`, and `M_STATS_PROBDIST_FENVBL.bas` /
`Export_FEnvelopeBothLarge` / `f_envelope_bothlarge_grid.csv`. Each macro
declares its own `VB_Name`, so the three can coexist in one workbook.

## Reading the result

The analyzer prints, per beta parameter, the worst relative error and whether it
meets the 1.1E-10 contract, then the largest beta parameter that still passes and
a conservative recommended envelope. `ERROR rows` in a band mean the VBA returned
a clean non-convergence error there (auditable); a large `worst rel err` with no
ERROR means a wrong answer returned silently (the case the policy must address).

That measured boundary is the input to the extreme-df policy decision: strict
rejection outside the envelope, or best-effort with a `Status`-channel warning.
