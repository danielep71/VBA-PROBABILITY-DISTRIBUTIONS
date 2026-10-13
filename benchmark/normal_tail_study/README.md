# Normal upper-tail kernel study (#72)

**Status: preregistered; derivation not yet run.** Excel-free. This study
designs and validates a replacement for `PROB_NormalUpperTailPositive` before
any VBA is written. It is not evidence that the kernel is fixed. The fix lands
only in a real-Excel wave, together with source, regression tests, exports
and truthful provenance.

Design: `PREREGISTRATION.md`, frozen before the committed derivation. Its §0
discloses the scratch exploration that chose the design.

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

## Files

| File | Role |
| --- | --- |
| `PREREGISTRATION.md` | frozen design, acceptance conditions and predictions for the Excel wave |
