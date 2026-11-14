# Student-t large-df tail study (#34)

**Status (2026-11-07).** The route shipped in #198 (merged 2026-11-06). Its
first Excel wave, run under `PRODUCTION_AMENDMENT_1.md`, falsified PT2
(`EXCEL_WAVE_RESULT.md`). The inverse polish
(`INVERSE_PREREGISTRATION.md`, #200 and #205) met every prediction and adopted
the survival, cumulative and quantile `envelope_domain` contracts at 5E-15
(`INVERSE_WAVE_RESULT.md`); #34 is closed. The seam study (`SEAM_STUDY.md`)
closed #210. #213, the inverse 1.29E-12 off near `|q| = 8` at `df = 1E7`, is
resolved by the rule of `ABOVE8_PREREGISTRATION.md` (`ABOVE8_RESULT.md`). The Defect, Result and Corrections sections record the
derivation as it was done, before any VBA existed; see
[Outcomes](#outcomes-2026-11-07) for what followed.

Design: `PREREGISTRATION.md`, frozen before any value was generated, with
one disclosed amendment recorded after the first run.

## Defect

`PROB_TryStudentTTail` evaluates the tail through `I_x(df/2, 1/2)` by a
Lentz continued fraction. In binary64 at `a = df/2 ≳ 5E4` the fraction is
ill-conditioned — coefficients `1 − O(m/a)`, rounding accumulating faster
than convergence — and reaches only ~1E-9 regardless of stopping rule. The
measured onset is **df ≈ 1E5** (1.25E-12 at t = 2 in the binary64 model;
Excel measures 3.16E-12, `PRODUCTION_AMENDMENT_1.md`), a decade below #34's
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
existing route. The CF is retained outside the region, where it is not
accurate to the bound: below df 1E5 it errs by up to 3.16E-12 (Excel), and
above |t| = 8 it erred by up to 7.8E-10 at df 1E8 (Excel) until #213 gave it
an exact-pair argument (3.4E-15 since). Both
are recorded in `numerical_limitations.csv`
(`IncompleteBeta.StudentTLargeDFFarTail`, `StudentT.LargeDFSeamMonotonicity`).

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
| `binary64_route_model.py` | binary64 model of the production route and the VBA normal-tail kernel |
| `test_binary64_route_model.py` | CI fixture pinning finding F1 |
| `model_fidelity.py` | true binary64 ULP distance, and the binding of the model to its VBA source |
| `source_binding.json` | SHA-256 of the bound VBA procedures, constants and model, and the origin of the frozen copies |
| `f1_frozen/` | the module and the 12 Excel observations F1 measured, as commit `495d0cf` held them (#81) |
| `test_model_fidelity.py` | CI fixture: ULP edge cases, three-step and source-mutation negatives |
| `PRODUCTION_FEASIBILITY.md` | finding F1: the adopted route is not buildable on the existing normal-tail kernel |
| `PRODUCTION_PREREGISTRATION.md` | the production route on the repaired kernel of #72, its VBA arithmetic, the Excel-wave predictions, the untouched holdout and the contracts, frozen before any VBA (#34) |
| `production_route_model.py` | binary64 model of that route, statement by statement, with its coefficient literal pairs |
| `predict_production.py` | measures the model against the oracle and records the predictions in `production_predictions.json` |
| `production_predictions.json` | A1–A3 and the per-row PT1/PT2 predictions |
| `test_production_route.py` | CI fixture: exact literals, reproducible predictions, A1 and A2 |
| `PRODUCTION_AMENDMENT_1.md` | Amendment 1: the outside-df control at Excel's measured accuracy |
| `check_excel_wave.py` | checks the first Excel wave against the PT1–PT4 predictions |
| `excel_wave_check.json` | that check's output |
| `EXCEL_WAVE_RESULT.md` | first wave result: PT2 falsified, no contract adopted |
| `INVERSE_PREREGISTRATION.md` | the inverse polish and contract adoption, frozen before its VBA |
| `predict_inverse.py` | records the inverse-wave predictions in `inverse_predictions.json` |
| `inverse_predictions.json` | the polish model's validation and the PI, PU and PH predictions |
| `test_inverse_polish.py` | CI fixture for the inverse polish model and its predictions |
| `test_vba_large_df_source.py` | CI fixture: the shipped VBA constants, arithmetic and wiring match the preregistered route |
| `check_inverse_wave.py` | checks the inverse wave against its predictions |
| `inverse_wave_check.json` | that check's output |
| `INVERSE_WAVE_RESULT.md` | inverse wave result: every prediction holds, contracts adopted |
| `SEAM_STUDY.md` | Excel behaviour across the df = 1E5 and \|t\| = 8 seams (#210) |
| `seam_study.py` | appends the 130 seam rows to `tchi_large_df` and checks them |
| `seam_results.json` | every seam measurement |

## Production feasibility (finding F1)

The truncation results above do not carry over to production unchanged.
Evaluated in binary64 with the existing VBA normal-tail kernel
(`PROB_NormalUpperTailPositive`), the adopted route is as inaccurate as that
kernel, up to 5.3E-10 at t = 6, and worse than the current continued fraction at
7 of the 36 fitting points inside the region. With a correctly rounded Q(t) it
meets 2^-50 everywhere. See `PRODUCTION_FEASIBILITY.md`; the production design
needs a decision and a new preregistration before VBA is written.

**2026-11-05 (#77).** The model's fidelity to Excel is now measured in true
binary64 ULPs: 9 of the 12 committed normal observations match exactly and the
other 3 are 1, 1 and 2 ULPs apart, within the unchanged two-ULP limit. The
"1.25 ulps" first recorded was relative error in units of machine epsilon. The
model is also bound to the VBA procedures and constants it mirrors, so a change to
the normal-tail kernel fails the F1 fixture until the finding is re-measured.
F1's conclusions are unchanged. See Amendment 1 of `PRODUCTION_FEASIBILITY.md`.

**2026-11-06 (#81).** #81 replaces the kernel F1 measured. F1 now stays bound to
a verbatim copy of the module at `495d0cf` and to the 12 Excel observations of
that kernel, both under `f1_frozen/`, with their origin recorded and checked.
See Amendment 2 of `PRODUCTION_FEASIBILITY.md`.

## Production route (2026-11-06)

#72 repaired the normal kernel through #81, so the route F1 could not build can
now be built. `PRODUCTION_PREREGISTRATION.md` freezes it before any VBA is
written: the study's order-5 expansion for `df ≥ 1E5`, `|t| ≤ 8`, evaluated in
VBA statement by statement on `PROB_NormalUpperTail` and
`PROB_ExpNegHalfSquare`.

- **Model measurement.** On 1670 validation points the route is within
  4.32E-16 of the oracle, below 2^-50. It is more accurate than the current
  continued fraction at every study fitting point where that misses the bound.
- **Predictions for the Excel wave.**
  - 58 forward rows: within 4 ulps and 2^-50.
  - 18 inverse rows: within a conditioning bound.
  - Every other Student-t, density and F row: unchanged.
- **Untouched holdout.** Seeded new rows are added to the independent holdout
  at the wave.
- **Contracts.** These follow if the predictions hold: 5E-15 for survival,
  cumulative and inverse over `df 1E5..1E8`, `|t| ≤ 8`.

## Outcomes (2026-11-07)

- **First wave** (run 42245883623 at `c8bb0ae`). PT1, PT3 and PT4 held; PT2
  failed at 7 of 18 inverse rows, through the inverse's stopping rule, not the
  route. No contract was adopted; the route shipped in #198 without one.
  See `EXCEL_WAVE_RESULT.md`.
- **Inverse wave** (run 42260290029 at `9cbe5d9`). Regression 955/955; PI, PU
  and PH hold. The three 5E-15 `envelope_domain` contracts were adopted (#200,
  #205), and #34 closed. See `INVERSE_WAVE_RESULT.md`.
- **Seams** (#210, closed). Survival and cumulative are not monotone across
  the seams, where the continued fraction serves the outside. See
  `SEAM_STUDY.md`.
- **Defect #213, resolved.** At `df = 1E7` near `q = −8` the inverse was
  1.29E-12 off, inside its 5E-15 contract's domain; it is within 1.2E-15
  since the tail above 8 was corrected (`ABOVE8_RESULT.md`).

## Oracle gap

Rmpfr has no incomplete-beta primitive. The independent leg is tanh-sinh
quadrature of the density. Recorded, not worked around.

The derivation did not read `benchmark/holdout/holdout_grid.csv` (559 rows
then; the production waves later appended 100 Student-t rows).
