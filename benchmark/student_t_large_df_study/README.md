# Student-t large-df tail study (#34)

**Status: derivation and validation complete, Excel-free.** This study
produces a validated series and a measured dispatch predicate. It is not
evidence that the kernel is fixed: no VBA exists yet, and the fix lands only
inside Phase 1's real-Excel wave, atomically with source, regressions,
exports and truthful provenance.

Design: `PREREGISTRATION.md`, frozen before any value was generated, with
one disclosed amendment recorded after the first run.

## Defect

`PROB_TryStudentTTail` evaluates the tail through `I_x(df/2, 1/2)` by a
Lentz continued fraction. In binary64 at `a = df/2 ≳ 5E4` the fraction is
ill-conditioned — coefficients `1 − O(m/a)`, rounding accumulating faster
than convergence — and reaches only ~1E-9 regardless of stopping rule. The
measured onset is **df ≈ 1E5** (1.25E-12 at t = 2), a decade below #34's
committed counterexamples at 1E6–1E8.

## Result

The upper tail `S(t; df) = Q(t) + φ(t) Σ g_k(t)/df^k`, with `g_k` derived
symbolically as exact rationals (`derive_coefficients.py` → `coefficients.json`,
orders 1–8; 266 s).

| Adopted by the preregistered rule | |
| --- | --- |
| Order `K*` | **5** |
| Region | **df ≥ 1E5 and \|t\| ≤ 8** — the most permissive candidate offered |
| Bound | `2^-50` (≈ 8.9E-16) |
| Worst fitting error inside region | 4.18E-16 at t = 8, df = 1E5 (the corner) |
| Worst binary64 CF inside region | 2.88E-9 at t = 3, df = 1E8 |
| Forward holdout (21 points inside) | worst 1.28E-20 |
| Inverse round-trip holdout (48 points) | worst 2.35E-36 |
| Oracle stability, 60 vs 100 dps | ≥ 51.9 digits |
| Quadrature leg agreement | ≥ 44.2 digits |
| Validation runtime | 31 s |

Wherever the CF exceeds the bound, the expansion beats it by 3 to 35
orders of magnitude. Seams — the CF route switch at every fitting df, and
both edges of the adopted predicate — were probed below, at and above.

**Dispatch:** `df ≥ 1E5 AND |t| ≤ 8` → order-5 expansion; otherwise the
existing route. The CF is retained outside the region and is accurate there.

## Corrections made along the way

- Coefficients recalled from memory for `g_2..g_4` in an earlier diagnosis
  were **wrong**; the derived `g_2 = t(t²−3)(3t⁴+2t²+1)/96`, not
  `(5t⁵+16t³+3t)/96`. Every coefficient here is derived, and
  `test_coefficients.py` re-derives orders 1–5 from scratch on every run.
- The first validation run failed on its own oracle: built as `1 − t_cdf`,
  it lost six digits at t = 8 to complement cancellation. Recorded in
  `PREREGISTRATION.md` Amendment 1 with the precision-pair rule correction.

## Files

| File | Role |
| --- | --- |
| `PREREGISTRATION.md` | frozen design and Amendment 1 |
| `derive_coefficients.py` | SymPy derivation; writes `coefficients.json` |
| `coefficients.json` | exact rational `g_1..g_8` with factored forms |
| `test_coefficients.py` | CI fixture: re-derives 1–5, parses all 8, pins numerically |
| `validate_expansion.py` | applies the design; writes `results.json` |
| `results.json` | every measurement, decision trace, seams, holdout |

## Oracle gap

Rmpfr has no incomplete-beta primitive. The independent leg is tanh-sinh
quadrature of the density. Recorded, not worked around.

The 559-row `benchmark/holdout/holdout_grid.csv` was not read.
