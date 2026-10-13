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
guard.

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
- **P2 Standardised surfaces.** For every committed row of `Normal_*` and
  `Lognormal_*` cumulative, survival and density, error ≤ `B + κ · δ`:
  - `κ` is the condition number with respect to `z`:
    - `z·φ(z)/Q(z)` for survival;
    - `|z|·φ(z)/Φ(z)` for cumulative;
    - `z²` for density.
  - `δ` bounds the relative error of the computed `z`:
    - `2^-52` for Normal;
    - `2^-52 + 2^-52·|ln x| / |ln x − μ|` for Lognormal, allowing one ulp in
      VBA `Log`.

  The derivation computes and records the bound for every such row.
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
- the P1 predicted values and P2/P3 bounds for every committed row they
  cover;
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
