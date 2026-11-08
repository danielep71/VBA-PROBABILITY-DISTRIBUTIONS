# Precision v2 bounded feasibility result

**250/250 independent points PASS. Feasibility only; not production acceptance.**

Protocol/candidate commit `7519fd2111e8c511c1cdf4fe5aeacfbec748d056` was published
before evaluating any of the 250 disjoint coordinates. The source-bound
`precision_v2_screen.json` retains every probability, reference and error.

- All references stable at 90/120 digits and agree with independent quadrature
  to relative 1E-80 or better.
- Maximum relative error: P 8.524E-17; Q 9.754E-17.
- Zero accuracy failures, unavailable oracles or adjacent monotonicity failures.
- Original A=.5 seam passes in the model.
- The 489 previously revealed fitting points also pass; reproduce with
  `replay_fitting_v2.py`. They are not included in the independent count.
- Four primitive fixture methods pass, including exact log(1), endpoint
  logGamma identities, two-Double residuals, range rejection and transcendental
  checks against high precision.

No runtime arbitrary-precision oracle is used by the candidate numerical path.
The degree-72 logGamma coefficient table reproduces identically at 100/160 dps.
High precision is used only by generators and verification.

## CI correction

The prior head's unit-layout failure was real: test_precision_probe.py had not
been listed in the common CI runner. Both v1 and v2 fixture modules are now
listed in benchmark/test_evidence_tools.py. All 32 local test-coverage guard
fixtures pass. The rejected v1 numerical module is unchanged; its negative
fixture explicitly reproduces log(1)=+2^-108. V2 separately requires exact zero.
This preserves the known defect rather than treating v1 as accepted. The earlier
written minus sign in the v1 report was corrected; it was not the measured sign.

## Limits and next acceptance gate

This remains a Python prototype, restricted by the frozen ranges. In particular,
A>1, extreme inputs, subnormal outputs and the complete public domain are not
certified here. The new two-Double CF stopping rule is part of this study only.
The broader VBA suite remains **1947 total / 1906 PASS / 41 FAIL** on the existing
B6 source. No v2 VBA port, numerical-source edit, export or manifest rebinding
has occurred. Main and release remain unchanged.

Before a production proposal: specify range handling and routing for the full
public domain; demonstrate arithmetic fidelity of a VBA port; cover caps,
underflow, boundary/caller behavior and the 41 retained failures; run fresh Excel
regression and exports; obtain current CI and review. The 250 independent points
are now revealed and cannot be reused as untouched holdout for a revised design.
