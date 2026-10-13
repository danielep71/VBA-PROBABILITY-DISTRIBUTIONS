# Normal upper-tail kernel study (#72)

**Status: derivation and validation complete; PASS under Amendment 1.** Excel-free. This study
designs and validates a replacement for `PROB_NormalUpperTailPositive` before
any VBA is written. It is not evidence that the kernel is fixed. The fix lands
only in a real-Excel wave, together with source, regression tests, exports
and truthful provenance.

Design: `PREREGISTRATION.md`, frozen before the committed derivation. Its §0
discloses the scratch exploration that chose the design. Amendment 1, written
after the first run, replaces one acceptance condition (A2) and is disclosed
as such.

## Defect

The current kernel is accurate in absolute terms but not in relative terms:
- In its Hart region (`3 ≤ z < 7.07`), relative error reaches 2.9E-9.
- Its exponential factor rounds `z²` first, which costs up to 5.7E-14 in the
  far tail.

Because of this, the survival contracts of `Normal_Survival`,
`Lognormal_Survival` and `NormalStandard_Survival` are limited to the central
region (`SurvivalTailRel`). It also blocks the #34 Student-t route (finding
F1).

## Candidate

- A degree-(7, 7) rational for `Q(z)·exp(z²/2)` on `[0, 3)`.
- A 60-level Laplace continued fraction on `[3, 38.5)`.
- An exponential factor split so that its main argument is exact.
- Every constant built from 15-significant-digit literal pairs, as the VBE
  requires.

The acceptance bound is `2^-50` relative, the bound of the #34 study.

## Result of the derivation (2026-10-13)

`derive_kernel.py` ran the preregistered fit in 74 s. `validate_kernel.py`
applied the design and wrote `results.json`.

**First run: failed A2 only (#74).** A2 required every coefficient to be
positive. The numerator's `z^7` coefficient is −2.54E-10. A2 exists so that
Horner evaluation cannot cancel, and the term's actual effect is a
cancellation factor of only 1 + 1.5E-7 on `[0, 3]`. The scratch prototype
(§0) had the same sign; its positivity check left out the last numerator
coefficient. Under §6 the study failed, and the holdout was not evaluated.

**Amendment 1.** The maintainer chose an amendment over a new
preregistration. A2 is replaced by A2′, which bounds the cancellation factor
by `1 + 2^-20`. The coefficients were not refitted, and nothing else changed.
The frozen A2 stays recorded as `A2_as_frozen`.

**Second run: PASS.** The holdout was evaluated once:

| Condition | Result |
| --- | --- |
| A1 approximant | 1.03E-17 ≤ 2^-55 on 23,001 points |
| A2′ cancellation (Amendment 1) | numerator 1 + 1.5E-7, denominator 1 (all positive), limit 1 + 2^-20 |
| A3 literals | all 17 constants rebuild their Double exactly |
| A4 continued fraction | truncation 2.0E-18 at z = 3 ≤ 2^-56, decreasing to 4.3E-50 at z = 10 |
| A5 kernel on V (23,751 points + seams) | worst Q 5.3E-16, 1 − Q 3.9E-16, density 3.4E-16, all ≤ 2^-50 (8.9E-16) |
| A6 holdout H (15,000 points, evaluated once) | worst Q 5.5E-16, 1 − Q 3.9E-16, density 3.3E-16 |
| A7 #34 route | worst 3.7E-16 at the 36 fitting points; 1.8E-16 at its 21 holdout points |
| A8 no regression | holds at all 3,227 points where today's kernel meets the bound |
| Oracle stability | ≥ 58.0 digits; quadrature leg ≥ 98.3 digits at 151 points |

Worst relative error by band, today's kernel against the candidate:

| band | today | candidate |
| --- | ---: | ---: |
| `0 ≤ z < 3` | 1.5E-14 | 5.3E-16 |
| `3 ≤ z < 7.07` | 2.9E-9 | 4.4E-16 |
| `7.07 ≤ z < 37.5` | 5.7E-14 | 5.0E-16 |

**Predictions for the Excel wave (§7).** All are recorded in `results.json`,
and none is outside its bound:
- P1, 41 direct rows: worst 3.7E-16.
- P2, 27 standardised rows: worst 2.3E-16.
- P3, 11 interval rows: worst 1.0E-16.
- P4, 80 inverse rows: worst 2.9E-15, with no predicted contract breach.
  The `split_boundary` inverse rows fall from 2.8E-11–3.8E-10 today to the
  binary64 floor (≤ 4.1E-16).

The model of today's code reproduces every affected committed Excel
observation within 4 ulps.

**Two defects in the validator itself were corrected before the result above
was recorded.** Neither touches the design or a threshold:
- The quadrature leg integrated `exp(−t²/2)` directly over `[z, ∞)`. That
  agreed to only 4.3 digits at z = 37.5. It now integrates
  `exp(−z s − s²/2)` on a `1/(z + 1)` scale, and agrees to ≥ 98 digits.
- `Normal_IntervalProbability` rows were read as `(lower, upper, mean)`. The
  density-helper exporter calls them as `(lower, upper, 0, sd)`.

**Next.** The kernel is ready to be implemented in VBA with the recorded
literals, in the Excel wave that §9 describes. That wave tests the §7
predictions against real Excel.

## Files

| File | Role |
| --- | --- |
| `PREREGISTRATION.md` | frozen design, acceptance conditions, predictions for the Excel wave, and Amendment 1 |
| `derive_kernel.py` | the preregistered fit; writes `kernel_coefficients.json` |
| `kernel_coefficients.json` | coefficients to 40 digits, their Doubles, and VBA `hi`/`lo` literal pairs |
| `kernel_model.py` | binary64 model of the candidate, of today's kernel, and of every affected public path |
| `validate_kernel.py` | applies §3–§8; writes `results.json` |
| `results.json` | every measurement, the acceptance record and the per-row predictions |
| `test_kernel_model.py` | CI fixture pinning the coefficients, literals and recorded result |
