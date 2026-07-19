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

| File | Role |
|---|---|
| `generate_f_envelope.py` | Writes `f_envelope_grid.csv` (references). |
| `_ibeta.py` | 50-digit continued-fraction incomplete beta (shared). |
| `f_envelope_grid.csv` | The sweep grid; `arg1=x`, `arg2=d1`, `arg3=d2`. |
| `f_envelope.bas` | Export macro `Export_FEnvelope` (F CDF + survival). |
| `analyze_f_envelope.py` | Degradation curve + measured envelope boundary. |

## How to run

1. Import `f_envelope.bas`, `Debug > Compile`.
2. Run `Export_FEnvelope`; select `f_envelope_grid.csv`.
3. Commit the filled grid.
4. `python3 analyze_f_envelope.py`.

## Reading the result

The analyzer prints, per beta parameter, the worst relative error and whether it
meets the 1.1E-10 contract, then the largest beta parameter that still passes and
a conservative recommended envelope. `ERROR rows` in a band mean the VBA returned
a clean non-convergence error there (auditable); a large `worst rel err` with no
ERROR means a wrong answer returned silently (the case the policy must address).

That measured boundary is the input to the extreme-df policy decision: strict
rejection outside the envelope, or best-effort with a `Status`-channel warning.
