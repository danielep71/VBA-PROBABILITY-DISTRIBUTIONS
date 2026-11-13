# Preregistration — normal upper-tail kernel repair (#72)

**Status: FROZEN before the committed derivation.** This file is the design.
The derivation commit that follows fits the coefficients, measures the kernel
against this design and records the result. It does not amend this file. If
the design proves wrong, the failure is recorded and a new preregistration is
written.

Governs: `benchmark/normal_tail_study/`. Excel-free. No `.bas` change, no
contract, no threshold, no grid, no observation, no provenance manifest.
`benchmark/holdout/holdout_grid.csv`, the independent holdout, is **not
inspected** at any point in this study.

## 0. Disclosure: exploration before this file

This design was not chosen blind. Before writing this file, an uncommitted
scratch prototype explored the design space on 2026-10-13. It used
`random.Random` seeds 1, 3, 5, 7, 11 and 20261013, a grid of `z = k/64`, and
the 36 fitting points of the #34 study. Everything it measured is design data,
not validation:

| explored | measured |
| --- | --- |
| Laplace continued fraction, 16 levels (today's depth) | truncation 5.5E-9 at `z = 3`, 4.2E-17 at `z = 7.07` |
| levels needed for truncation ≤ 2E-17 | 55 at `z = 3`, 26 at `z = 5`, 17 at `z = 7.07` |
| continued fraction, 60 levels, in binary64 | worst 1.6 ulp on `[3, 37.5]` |
| one rational for `[0, 6]`, degree (9, 9) | approximant 9.9E-20; kernel worst 4.8 ulp, 7.6E-16 relative |
| rational on `[0, 3]`, degree (7, 7) / (8, 8) / (9, 9) | approximant 1.0E-17 / 2.2E-20 / 4.1E-23; binary64 evaluation 3.4 / 4.2 / 3.5 ulp |
| rationals in `z − 1.5`, `[0, 3]` | 29–35 ulp: cancellation near `z = 0`; rejected |
| shorter rational pieces (`[2, 4]`, `[3, 6]`, `[4, 6]`) | 3.0–4.9 ulp: no better than one piece; rejected |
| unsplit `Exp(-0.5 * Z * Z)` | up to 5.7E-14 relative where `Q` is a normal Double |
| hybrid of §2 below, with the (7, 7) prototype fit | kernel worst 4.4 ulp, 5.2E-16 relative on 14,403 points; #34 route worst 3.7E-16 at its 36 fitting points |
| 60-level truncation, exact arithmetic | 2.0E-18 at `z = 3`, 9.1E-24 at `z = 4` |

The thresholds in §6 were checked for feasibility against these figures
before this file was frozen. In particular, A4 was set at `B/64` after the
60-level truncation at `z = 3` was measured.

Because of this, validation (§5) uses points generated from seeds and a grid
that the exploration did not use. The #34 route check (§6, A7) reuses the #34
fitting points, which the exploration did see. That check is a regression
requirement, not independent evidence.

## 1. Defect under study

`PROB_NormalUpperTailPositive` (`src/M_STATS_PROBDIST_NORMALFAMILY.bas`)
returns `Q(z) = P(N(0,1) > z)` for `z ≥ 0`. It has two defects:

- **Hart region.** Below `z = 7.07106781186547`, a Hart rational approximation
  is accurate in absolute terms, not relative. Measured relative error is
  1.5E-14 at `z = 3`, 5.3E-10 at `z = 6`, and 2.9E-9 near `z = 7.07`.
- **Exponential factor.** `Exp(-0.5 * Z * Z)` rounds `Z * Z` before the
  exponential. Wherever `Z * Z` is inexact, that costs up to `z²/2 · 2^-53`
  relative error, 5.7E-14 at `z ≈ 33`. `PROB_NormalPDF` has the same
  expression.

The kernel feeds:
- the Normal, Lognormal and Standard Normal cumulative, survival and interval
  functions;
- through `PROB_NormalCDF` and `PROB_NormalPDF`, the Halley step of their
  inverses.

It blocks #34: finding F1 shows the adopted Student-t route meets `2^-50` only
with an accurate `Q`.

## 2. Candidate kernel

All arithmetic is binary64 with VBA semantics. The derivation models it line
for line.

**Constants.** The VBE stores a numeric literal to 15 significant digits, so
every constant enters as one or two such literals:
- Each target constant is first rounded to the nearest Double.
- `hi` is that Double rounded to 15 significant digits.
- `lo` is the exact remainder `Double − hi`, rounded to 15 significant digits.
- The VBA constant is the expression `hi + lo`, evaluated in Double.
- The derivation must show that `hi + lo` is bit-identical to the target
  Double (A3). A constant that is exact in 15 digits has `lo = 0`.

This is the construction `PROB_INV_SQRT_TWO_PI` already uses. It assumes the
VBE converts a 15-digit literal to its nearest Double; the Excel wave checks
that through the predictions in §7.

**Exponential factor.** For `z ≥ 0`:

    zh   = Int(z * 256) / 256                      ' exact
    E(z) = Exp(-0.5 * (zh * zh)) * Exp(-0.5 * ((z - zh) * (z + zh)))

- `zh * zh` is exact, because `zh` has at most 14 significant bits below 38.5.
- `z - zh` is exact.
- The second argument is at most 0.15 in magnitude, so its rounding error is
  far below one ulp of `E`.

**Region R1, `0 ≤ z < 3`:** `Q(z) = E(z) * (P(z) / D(z))`.
- `P` and `D` have degree 7, with `D(0) = 1`.
- Both are evaluated by Horner in `z`.
- They approximate `f(z) = Q(z) · exp(z²/2)` on `[0, 3]`.

**Region R2, `3 ≤ z < 38.5`:** `Q(z) = E(z) * (C / d)`.
- `C = 1/√(2π)`, as a constant pair.
- `d` is the Laplace continued fraction with 60 levels, evaluated backwards:
  `d = z`, then `d = z + k / d` for `k = 60, 59, …, 1`.

**Unchanged:**
- `z ≥ 38.5` returns 0.
- Negative arguments, the complements in `PROB_NormalCDF` and
  `PROB_NormalSurvival`, the interval logic and the inverse are not modified.

**Density:** `PROB_NormalPDF(z) = C * E(|z|)` inside the existing underflow
guard. `PROB_NormalPDF` serves `NormalStandard_Density` and the Halley step.
`Normal_Density` and `Lognormal_Density` compute the density in log space
without it. Their own exponent rounding is outside #72.

**Fit.** The fit is deterministic and uses mpmath at 60 digits:
- linearised least squares for relative error at the 500 Chebyshev nodes of
  `[0, 3]`;
- weights `1 / (f · D_prev)`;
- Lawson reweighting by the node errors;
- 120 iterations, keeping the iterate with the smallest node maximum.

The coefficients are recorded to 40 digits and as `hi`/`lo` literals.

**Placement (for the Excel wave, not this study).** Per the architecture note
in `M_STATS_PROBDIST_CORE`:
- the kernel moves to `M_STATS_PROBDIST_SPECIALFUNCS` as a project-scoped
  helper, so the Student-t route can call it without a private copy;
- the exponential factor goes to the core numerics layer.

The product API (`K_STATS_*`) does not change.

## 3. Oracles

| Role | Method | Precision | Standing |
| --- | --- | --- | --- |
| Primary | `erfc(z/√2)/2` and `exp(−z²/2)/√(2π)` | 60 and 100 dps | evidential |
| Independent | tanh-sinh quadrature of the density on `[z, ∞)` | 100 dps | evidential, algorithm-independent |

- **Stability:** the primary at 60 and at 100 dps must agree to at least 50
  significant digits at every evaluated point.
- **Independence:** the quadrature leg is evaluated at every 25th point of the
  grid in §5. It must agree with the primary to at least 40 significant digits.

A failure of either is a study failure, not a data point.

## 4. Bound

`B = 2^-50` (≈ 8.9E-16) relative. This is the bound of the #34 study, so a
kernel that meets it cannot be the limiting term of the #34 route. Errors are
measured only where the reference value is a normal Double (`≥ 2^-1022`).
Below that, relative accuracy is not defined for a binary64 result. That band
is characterised but not bounded.

## 5. Evaluation sets

None of these seeds or grids was used in the exploration.

- **Validation V:**
  - the grid `z = k/100` for `k = 0 … 3750`;
  - 20,000 values of `random.Random(72001).uniform(0, 37.5)`;
  - the seam probes below.
- **Holdout H:** evaluated **once**, after V has passed and the result has been
  written:
  - 10,000 values of `random.Random(72002).uniform(0, 37.5)`;
  - 5,000 values of `random.Random(72003).uniform(0, 8)`.
- **Approximant set A:**
  - 20,000 values of `random.Random(72004).uniform(0, 3)`;
  - the grid `k/1000` for `k = 0 … 3000`.
- **Seams:** each seam is probed at the Double below, at it, and at the Double
  above:
  - `z = 3`;
  - `z = k/256` for `k = 1 … 9855` in steps of 97 (the `zh` steps);
  - the largest `z` with a normal `Q` (characterisation only);
  - `z = 38.5`.

## 6. Acceptance

Every condition must hold. Each must be recorded as pass or fail.

- **A1 Approximant.** `P/D`, with the 40-digit fitted coefficients and
  evaluated at 60 digits, has relative error ≤ `2^-55` against `f` on set A.
  Rounding the coefficients to Double is part of the binary64 model, so A5
  covers it.
- **A2 Positivity.** Every coefficient of `P` and `D` is positive, so Horner
  evaluation for `z ≥ 0` has no cancellation.
- **A3 Literals.** Every constant's `hi + lo` is bit-identical to its target
  Double.
- **A4 Continued fraction.** With 60 levels, truncation error (exact
  arithmetic) is ≤ `2^-56` (`B/64`) at `z = 3`, and does not increase at
  `z = 3.5, 4, …, 10`.
- **A5 Kernel.** The binary64 model meets `B` on V:
  - for `Q`, for `1 − Q` (the cumulative for `z ≥ 0`), and for `φ`;
  - at every seam probe where the reference is normal.
- **A6 Holdout.** A5 holds on H, evaluated once.
- **A7 #34 route.** The order-5 route of `binary64_route_model.py`, with the
  candidate `Q` and with `φ` formed as `C * E(t)`, meets `B`:
  - at the 36 fitting points of the #34 region;
  - at the #34 study's forward holdout points inside the region.
- **A8 No regression.** At every point of V where the current kernel meets
  `B`, the candidate also meets `B`. A5 implies this, and it is recorded
  explicitly.

**Failure conditions.** A failure of A1–A8, oracle instability, or a fit
that does not converge fails the study. A failed study is recorded as such
and produces no VBA recommendation. No parameter in §2 is tuned after this
file is committed: not the split at 3, the degrees, the depth 60, the factor
256 or the fit settings.

## 7. Predictions for the Excel wave

These are frozen now and checked when the VBA is implemented and exported with
Excel. Here "error" is relative error against the committed reference at the
row's binary64 arguments.

- **P1 Direct surfaces.** For every committed row of
  `NormalStandard_Survival`, `NormalStandard_Cumulative` and
  `NormalStandard_Density`, the Excel observation:
  - is within 4 ulps of the derivation's predicted binary64 value, which
    allows for VBA `Exp` differing from C `exp` in the last bit;
  - has error ≤ `B`.

  This applies to the main grid, `survival_boundary` and `holdout_older`.
- **P2 Standardised surfaces.** For every committed row of
  `Normal_Cumulative`, `Normal_Survival`, `Lognormal_Cumulative` and
  `Lognormal_Survival`, error ≤ `B + κ · δ`:
  - `κ` is the condition number with respect to `z`:
    - `z·φ(z)/Q(z)` for survival;
    - `|z|·φ(z)/Φ(z)` for cumulative.
  - `δ` bounds the relative error of the computed `z`:
    - `2^-52` for Normal;
    - `2^-52 + 2^-52·|ln x| / |ln x − μ|` for Lognormal, allowing one ulp in
      VBA `Log`.

  The derivation computes and records the bound for every such row.

  `Normal_Density` and `Lognormal_Density` are not affected. They evaluate the
  density in log space, as `Exp(-0.5·Z·Z − Log(σ) − …)`, and call neither the
  kernel nor `PROB_NormalPDF`. This repair does not change them, and their
  observations are predicted to be bit-identical to the committed ones.
- **P3 Interval surfaces.** For every committed row of
  `NormalStandard_IntervalProbability` and `Normal_IntervalProbability`, error
  ≤ `(Q₁·b₁ + Q₂·b₂) / P`:
  - `Q₁` and `Q₂` are the two tail terms the existing branch subtracts;
  - `b₁` and `b₂` are their P1/P2 bounds;
  - the final subtraction from one adds `2^-53 / P`.
- **P4 Inverses.** The derivation models the inverse path:
  - the Acklam seed `PROB_NormalInvCDFRaw`;
  - one Halley step on the candidate cumulative and density;
  - the affine map (Normal) or `Exp` (Lognormal).

  It records a predicted value and error for every committed
  `…InverseCumulative` and `…InverseSurvival` row. Each Excel observation must
  be within 4 ulps of its predicted value (8 for Lognormal, for VBA `Exp` and
  `Log`).
  - The `split_boundary` rows (contracts 1E-10, 2E-10 and 1E-9) fall to
    ≤ 1E-14.
  - Every row meets its current contract, unless the derivation names it in
    advance as a predicted breach.
  - `…InverseCumulativeFast` does not use the kernel and is unchanged bit for
    bit.
- **P5 Independent holdout.** The P1–P4 formulas apply unchanged to the
  affected rows of `benchmark/holdout/holdout_grid.csv` when it is next
  exported. That grid is not read by this study.

Any observation outside its predicted bound falsifies the prediction. It is
recorded against #72 before any contract is changed.

## 8. Recorded outputs

The derivation commit must record, machine-readably:
- the fitted coefficients (40 digits, Double bits, `hi`/`lo`);
- A1–A8 with the worst point for each;
- oracle stability and independence figures;
- current-kernel and candidate errors by band;
- seam results;
- the P1 and P4 predicted values, and the P2/P3 bounds, for every committed
  row they cover outside the independent holdout;
- runtime.

A fixture in the evidence-tool suite pins the coefficients, the literals and
a subsample of A5 and A7, so the result cannot silently become stale.

## 9. What this study does not do

- It writes no VBA, and changes no contract, threshold, grid, observation or
  manifest.
- It does not read the independent holdout.
- It is not evidence that the kernel is fixed.

The fix lands in an Excel wave, atomically with source, regression tests,
exports and provenance:
- `SurvivalTailRel` is retired or restated there;
- the survival contracts are widened only to the domain the Excel evidence
  supports;
- the #34 route is implemented on the new kernel under #34's own production
  preregistration.

---

## Amendment 1 — recorded after the derivation run

**Disclosure.** This amendment was written on 2026-10-13, after the
derivation run recorded in `results.json` (merged in #74) had failed A2 and
no other condition. The maintainer chose an amendment over a new
preregistration. Nothing above this section has been changed. The design
(§2), the fitted coefficients, every other condition and threshold, the
evaluation sets (§5) and the predictions (§7) all stand as frozen.

**Defect in A2.** A2 required every coefficient of `P` and `D` to be
positive, "so Horner evaluation for `z ≥ 0` has no cancellation".
Positivity is a sufficient condition for that purpose, not the purpose
itself:
- The preregistered fit produced one negative coefficient: the numerator's
  `z^7` term, −2.54E-10.
- Its effect on `[0, 3]` is a cancellation factor
  `Σ|c_k| z^k / |Σ c_k z^k|` of at most 1 + 1.5E-7.
- At that size it changes Horner's rounding-error bound by under one
  millionth of itself.

The scratch prototype had the same sign. Its positivity check skipped the
last numerator coefficient, so the defect was not caught before freezing.

**A2 is replaced by A2′:** for `P` and for `D`, with the Double coefficients,

    max over set A of  Σ|c_k| z^k / |Σ c_k z^k|  ≤  1 + 2^-20.

The threshold was chosen after seeing 1 + 1.5E-7, and that is disclosed
here. The justification is independent of the observed value: a factor
below `1 + 2^-20` inflates the Horner error bound by less than `2^-20` of
itself, which is below 0.001 ulp at the kernel's measured error. A5 measures
the binary64 result directly in any case. The failed A2 stays in
`results.json` as `A2_as_frozen`.

**Holdout.** H was not evaluated in the failed run, because §6 stops at a
failed condition. It is evaluated once, under this amendment. No parameter
in §2 was changed or refitted.

---

## Note — 2026-11-03: the one-shot holdout record is now enforced (#80)

**This note is not an amendment.** It changes nothing above: not the design,
the conditions, the thresholds, the evaluation sets or the predictions. It
records no new measurement.

- H was evaluated in the run recorded on 2026-10-13 (#75), under Amendment 1.
  Its A6 result stays in `results.json` unchanged, and the recorded status
  stays PASS.
- The validator of #75 kept the "evaluated once" of §5 only by convention.
  Every run of it evaluated H and wrote `evaluated_once: true` again, so it
  could neither prevent nor detect a repeat. The committed record reports one
  evaluation. Nothing shows a repeat, but that validator could not have shown
  one.
- From #80 on, `results.json` carries a `holdout_exposure` record with status
  `EVALUATED` and one exposure. It binds the coefficient file, the first
  15,729 bytes of this file (its text through Amendment 1), the SHA-256 of
  H's seeded points and the SHA-256 of the A6 block. `validate_kernel.py`
  refuses a second first evaluation, and its replay and regeneration modes
  never evaluate H.
- H was not evaluated to make this change.
- Later amendments must be appended after this note. The recorded design
  identity accepts appended text only.

---

## Amendment 2 — 2026-11-03: conditioning of the P2 bounds (#78)

**Disclosure.** I wrote this amendment on 2026-11-03. By then these results
had been observed:
- the holdout H, evaluated once on 2026-10-13 (#75, run started 06:45:41Z),
  A6 PASS;
- every other measurement of that run: A1–A5, A7, A8, the oracles, the seams
  and the bands;
- the §7 predicted values and bounds that run recorded for every covered row,
  all within their bounds;
- the committed Excel observations of today's code at those rows, which the
  predictions read as inputs.

No Excel observation of the candidate kernel exists, so no §7 prediction has
been tested yet. Nothing above this section has been changed. The design
(§2), the fitted coefficients, A1–A8 and their thresholds, `B`, the
evaluation sets (§5) and P1, P3 and P4 stand as frozen. Nothing was refitted,
and H was not evaluated to write or apply this amendment. No acceptance
condition depends on P2, so none can change status through it.

**Defects in P2 as frozen.** Two expressions in §7 P2 do not give an error
bound for every covered row:
- The survival condition number `z·φ(z)/Q(z)` is negative for `z < 0`. There
  `B + κ·δ` is below `B` and falls as `δ` grows, so it is not an upper bound.
  Four covered survival rows have `z < 0`: `Normal_Survival(3, 5, 2)`,
  `Normal_Survival(-2.1, 0.4, 1.7)`, `Lognormal_Survival(0.5, 0, 1)` and
  `Lognormal_Survival(0.45, 0.2, 0.6)`.
- The Lognormal `δ` divides by `|ln x − μ|`, which is zero at the centre
  `ln x = μ`. Two covered rows are there: `Lognormal_Cumulative(1, 0, 1)` and
  `Lognormal_Survival(1, 0, 1)`. Their `δ` is 0/0, so `κ·δ` is undefined.

From the first derivation run (#74) on, the validator computed `|z|·φ/F` and
dropped the Lognormal `δ` term at a zero denominator, without disclosing
either. The recorded bounds therefore came from a rule this text did not
state. This amendment states that rule and derives it.

**Conditioning for both signs.** A covered row evaluates `F` at its exact
standardised argument `z`, where `F` is the survival `S(z) = Q(z)` (for
`z < 0`, `1 − Q(|z|)`) or the cumulative `Φ(z)`. The relative condition
number of `F` with respect to `z` is

    κ_F(z) = |z · F′(z) / F(z)|,

so that, to first order, `|F(z(1 + δ)) − F(z)| / F(z) = κ_F(z)·|δ|` for
either sign of `z` and of `δ`. Since `S′ = −φ`, `Φ′ = φ`, and `φ`, `S` and
`Φ` are positive for every finite `z`:

    κ_S(z) = |z|·φ(z) / Q(z),        κ_Φ(z) = |z|·φ(z) / Φ(z).

Both are nonnegative. The frozen survival expression is `κ_S` for `z ≥ 0` and
`−κ_S` for `z < 0`; it is replaced by `κ_S`. The frozen cumulative expression
already is `κ_Φ` and is unchanged.

**Absolute form of the standardisation term.** `κ·δ` is the relative error in
`F` caused by a relative error `δ` in `z`. With `η = δ·|z|`, a bound on the
absolute error `|z̃ − z|` of the computed argument, it equals
`φ(z)·η / F(z)`. The arithmetic bounds `η` directly. With `u = 2^-53` and
`|ε₁|, |ε₂| ≤ u`:
- Normal: `z̃ = fl(fl(x + (−μ)) / σ) = z(1 + ε₁)(1 + ε₂)`, so
  `|z̃ − z| ≤ (2u + u²)|z|`, and `η = 2^-52·|z|`.
- Lognormal: `L̃ = Log(x)` with `|L̃ − ln x| ≤ 2^-52·|ln x|`, the one-ulp
  allowance of the frozen text, and `z̃ = fl(fl(L̃ + (−μ)) / σ)`. Then

      z̃ − z = (L̃ − ln x)(1 + ε₁)(1 + ε₂) / σ + z·((1 + ε₁)(1 + ε₂) − 1),

  so to first order `η = 2^-52·(|z| + |ln x| / σ)`.

Where `z ≠ 0`, `η / |z|` is exactly the frozen `δ`: `2^-52` for Normal, and
`2^-52 + 2^-52·|ln x| / |ln x − μ|` for Lognormal, because
`|z| = |ln x − μ| / σ`. P2 is therefore restated as

    error ≤ B + φ(z)·η / F(z),

which equals the frozen `B + κ·δ`, with `κ ≥ 0`, wherever `δ` is defined,
and is defined for every row. P3 takes "their P1/P2 bounds" and so inherits
this term for each tail: `η = 2^-52·|z|` for `Normal_IntervalProbability`,
and `η = 0` for `NormalStandard_IntervalProbability`.

The rule stays first order, as frozen. The terms it neglects are smaller than
the standardisation term by a factor of order `|z|·η + 2^-52`, below 1E-14 at
every covered row.

**Centre.** For a positive Double `x` and a Double `μ`, `ln x = μ` holds only
at `x = 1, μ = 0`. For any other positive rational `x`, `ln x` is
transcendental (Lindemann–Weierstrass), so it is not a Double. At `x = 1` the
frozen allowance `2^-52·|ln x|` is zero, and a faithfully rounded `Log`
returns the exact result, 0. Then `z̃ = (0 + (−0)) / σ = 0` exactly, `η = 0`,
and the bound is `B`: the row evaluates the kernel at `z = 0`, where P1 and
A5 already bound `Q` and `1 − Q` by `B`. This is the value of the absolute
form, not a special case: `φ·η / F` is continuous in `(x, μ)`, and it tends
to 0 as `x → 1` with `μ = 0`, because `η` does.

Near the centre, where `ln x ≠ μ` but `|ln x − μ| ≪ |ln x|`, the frozen `δ`
is very large and `κ` very small. Their product is finite and tends to
`√(2/π)·2^-52·|μ| / σ` as `z → 0`, since `F(0) = 1/2`. The absolute form
evaluates it without forming `0·∞`.

The term has no seam of its own. It is continuous in `z` through 0, the split
at 3 and the 37.5 and 38.5 cut-offs; the kernel's own seams are unchanged and
covered by A5. Where the reference `F` is not a normal Double (survival
above the last `z` with a normal `Q`, cumulative below its negative), the
bound is finite but, as §4 states, the error is characterised, not bounded.

**What changes in the record.** `results.json` is regenerated without
evaluating H:
- every P2 row records `z`, `κ` and `η`, and its bound as
  `B + φ(z)·η / F(z)`;
- the P2 and P3 bounds the validator recorded already followed this rule. All
  are unchanged except `Lognormal_Cumulative(0.45, 0.2, 0.6)`, which moves by
  one ulp, from evaluating the same quantity in another order;
- every predicted value and predicted error is unchanged, and every row stays
  within its bound;
- under the frozen sign, the four negative-`z` survival rows would have had
  bounds of `B − κ·δ`. These are still above their predicted errors, so no
  recorded outcome depended on the correction.

`test_conditioning.py` pins this: the recorded P2 rows, negative-`z`, exact
centre and near-centre rows, and the seams.

---

## Amendment 3 — 2026-11-03: fit convergence and result provenance (#79)

**Disclosure.** I wrote this amendment on 2026-11-03. By then these results
had been observed:
- the fit of §2 and its coefficients, recorded on 2026-10-13 (#74), with a
  node maximum of 1.03E-17 and A1 at 1.03E-17;
- the holdout H, evaluated once on 2026-10-13 (#75, run started 06:45:41Z),
  A6 PASS;
- every other measurement, prediction and bound recorded since, including
  those of Amendment 2;
- the convergence evidence of the fit, described below. I observed it before
  I fixed the tolerance, by running the frozen fit again with that evidence
  recorded.

Nothing above this section has been changed. The fit of §2 stands as frozen:
the nodes, weights, Lawson reweighting, precision, the cap of 120 iterations
and the selection of the iterate with the smallest node maximum. Nothing was
refitted: the run reproduced every field of `kernel_coefficients.json` except
its runtime, and the file was left byte-identical. A1–A8, their thresholds,
`B`, the evaluation sets of §5 and the predictions of §7 are unchanged. H was
not evaluated to write or apply this amendment.

**Defect: no convergence criterion.** §6 makes "a fit that does not converge"
a study failure, but §2 defines no criterion. The derivation ran 120
iterations, kept the best iterate and recorded no convergence decision, so no
record could show a failure. A small A1 error is not a convergence result: a
stalled or oscillating iteration can still leave an iterate that meets A1.

**Criterion.** Let `e(x) = P(x) / (D(x) f(x)) − 1` be an iterate's signed
relative error at the 500 nodes, in increasing `x`. Split the nodes into
maximal runs of one sign; each run contributes its largest `|e|`, so
consecutive extrema alternate. Let `L` be the largest `t` such that the
extrema of size at least `t` still alternate at `m + n + 2 = 16` points.

By de la Vallée Poussin's theorem for weighted rational approximation, no
rational of type (7, 7) with a denominator positive on the nodes has a smaller
node maximum than `L`. If one did, `P*·D − P·D*` would change sign at least
15 times among the 16 points, but it is a polynomial of degree at most 14 and
not identically zero. The levelling ratio `ρ = max|e| / L − 1` therefore bounds
how far an iterate is from the best node maximum of its type.

The selected iterate is:
- **CONVERGED** if `e` is finite, `D > 0` at every node, `e` alternates at 16
  or more points, and `ρ ≤ 0.01`;
- **CAP_EXHAUSTED** if it is not, and it is the 120th iterate. The node
  maximum was still falling when the cap stopped the fit;
- **NOT_CONVERGED** otherwise. The iteration stalled or oscillated without
  levelling, or produced a non-finite or nonpositive iterate.

CAP_EXHAUSTED and NOT_CONVERGED fail the study under §6, and so does a
missing record or status. The cap is not raised to rescue a fit, because §2
freezes it. For a failed fit, `derive_kernel.py` records the status, writes
no coefficient file, accepts no existing one and exits 1. The validator then
records a problem, and the study fails before any VBA recommendation.

**Tolerance.** I chose 1% after seeing the ratio of this fit, 3.34E-3. The
choice is not blind. A tolerance below 3.34E-3 would have failed this fit as
CAP_EXHAUSTED, because the selected iterate is the last one. The
justification does not depend on the observed value. At `ρ ≤ 0.01`, no (7, 7)
rational can lower the node maximum by more than 1% of itself: about 1E-19
here, below 2^-61.6 whenever A1 holds. That is under 1/3000 of `B`, so a fit
closer to the best approximation could not change a kernel result that A5 or
A6 measures.

**Observed evidence.** These figures are recorded in `fit_convergence.json`:
- the node maximum fell at every one of the 120 iterations, from 8.29E-17
  to 1.0323E-17;
- every iterate alternated at 16 points;
- the levelling ratio fell from 70.6 to 3.34E-3, and first reached 1% at
  iterate 53;
- the selected iterate is the 120th: node maximum 1.03232E-17, `L` =
  1.02888E-17, minimum `D` 1.000013;
- over the last ten iterations the node maximum moved by 2.4E-4 of itself.

The status is CONVERGED.

The validator does not take the record on trust. It requires CONVERGED for
the SHA-256 of `kernel_coefficients.json`, under this criterion, and checks
that the status follows from the recorded history and selection. It then
recomputes the evidence from the committed 40-digit coefficients and requires
the criterion there as well.

**Defect: incomplete result provenance.** `results.json` bound the
coefficient file, and the holdout record also binds the frozen design text
and H. Nothing bound the generator, the models, the validator, the full
design or the inputs the predictions read. A change to any of them could leave
recorded measurements stale while the fixture, which recomputes only a
subsample, stayed green.

**Binding.** This part is a record control, not a design change.
`results.json` now carries a `provenance` block. It holds the SHA-256 of
every file that produces or shapes the record:
- the generator `derive_kernel.py` and its record `fit_convergence.json`;
- the models `kernel_model.py`, `binary64_route_model.py` and the #34 route
  coefficients;
- the validator `validate_kernel.py`, `holdout_record.py` and
  `result_binding.py`;
- this file, whole, with its amendments;
- the three grids of the §7 predictions, `accuracy_contracts.csv`, and the
  #34 study's `results.json`, which supplies the A7 points.

The candidate stays bound by `coefficients_sha256`. The frozen prefix and H
stay bound by `holdout_exposure`, unchanged.

A fixture compares those identities with the files in CI, in well under a
second, and fails on any difference. A legitimate change to a bound file
therefore needs `validate_kernel.py --regenerate --reason TEXT`. That run
recomputes everything except H and records which identities changed. It
cannot change the candidate or the frozen text, because the holdout record
refuses both.

**What changes in the record.** `results.json` is regenerated without
evaluating H. It records:
- `acceptance.fit_convergence`, PASS;
- the `provenance` block;
- this amendment in `amendments`;
- a regeneration entry with `holdout_evaluated: false`.

Every other recorded value is unchanged, and the recorded status stays PASS.
`test_fit_convergence.py` and `test_result_binding.py` pin this. They include
deliberate nonconvergent, capped and missing-status fits, and mutations of
each bound file.

---

## Amendment 4 — 2026-11-13: grids bound by row (#302)

**Disclosure.** I wrote this amendment on 2026-11-13. By then every result
listed in the disclosure of Amendment 3 had been observed, and nothing since:
H was evaluated once, on 2026-10-13, and was not evaluated to write or apply
this amendment.

Nothing above this section has been changed. This amendment is a record
control, not a design change: the fit, A1–A8, their thresholds, `B`, the
evaluation sets of §5 and the predictions of §7 are unchanged.

**Defect: the binding was wider than the input.** Amendment 3 bound the three
grids of the §7 predictions by the SHA-256 of each file. The predictions read
only the rows of the normal and lognormal functions. Every promotion of rows
for another function therefore made the record stale and forced a
regeneration that changed only metadata: #23, #289 and #327.

**Binding.** The three grids are now bound by the rows the predictions read:
- the rows whose function is one of those `validate_kernel.predictions()`
  reads, listed in `result_binding.STUDY_FUNCTIONS`;
- each identified by its canonical key (IEEE-754 argument bits, regime,
  evidence set), with its claim, metric, reference, observation and
  expected_error;
- hashed in canonical-key order by `benchmark/_row_binding.py`.

The `provenance` entry of each grid records the row count and `rows_sha256`
in place of the file SHA-256. A row added for another function, or a
different order of rows, leaves the record current. A change to any bound
row, an added or removed bound row, or a duplicate key makes it stale. Every
other bound file stays bound by its file SHA-256, as in Amendment 3.

**What changes in the record.** `results.json` is regenerated without
evaluating H. It records the three row bindings, this amendment in
`amendments`, and a regeneration entry with `holdout_evaluated: false`. Every
other recorded value is unchanged, and the recorded status stays PASS.
`test_result_binding.py` pins the new behaviour in both directions.
