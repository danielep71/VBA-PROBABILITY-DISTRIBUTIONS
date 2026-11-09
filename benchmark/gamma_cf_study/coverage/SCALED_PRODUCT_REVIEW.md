# Exponent-separated multiplication study

2026-11-08; continuation of #24/#26 and PR #242 from `f1b72da`.
This is a primitive implementation and exact-rational regression, not a new
Gamma candidate or a new independent probability validation set.

## Problem and implementation

The frozen v2 primitive's unscaled splitter overflows for `(2^1000, 2^-1000)`
although the exact product is 1. It also cannot retain the residual of
`(2^-1022, nextafter(0.5, +infinity))` in an unscaled Double pair.
Both negative controls remain unchanged in `test_precision_arithmetic_audit.py`.

`scaled_product.py` now separates each operand into a binary64 mantissa and
integer exponent before using the existing product transformation. The absolute
mantissas are in `[0.5, 1)`, so splitter multiplication stays finite and the
nonzero product residual stays normal. It returns a normalized pair and the
summed exponent; neither the full product nor an unscaled residual is formed.
The representation therefore retains products below the minimum subnormal and
above the maximum finite binary64 value until final conversion.

Final conversion raises on overflow and preserves signed zero on underflow.
For subnormal results, it works in integer units of `2^-1074` and retains the
low component when deciding ties to even. Independently rounding the high and
low components at their final exponent is incorrect. The fixture with
`(hi, lo, exponent) = (0.625, 2^-56, -1072)` demonstrates this: high-only scaling
rounds to 2 subnormal units, while the represented value rounds to 3.
Malformed, nonfinite or nonnormalized inputs fail closed. These are internal
study exceptions, not a change to worksheet error policy.

The numerical path uses binary64 operations, `frexp`/`ldexp`, and an integer
exponent. Python Fractions and arbitrary-precision arithmetic appear only in
the test oracle. The integer in subnormal rounding is at most `2^52`, exactly
representable in a VBA Double; it must not be ported to a 32-bit VBA Long.
A VBA frexp/ldexp equivalent and actual host rounding still require a port and
Excel verification. This file does not claim that those steps have happened.

## Reproducible checks

Run from repository root:

```bash
python benchmark/gamma_cf_study/coverage/test_scaled_product.py
python benchmark/gamma_cf_study/coverage/test_precision_arithmetic_audit.py
```

Seven new test methods pass locally:

- Both retained v2 counterexamples, now represented exactly before final rounding.
- Every binary exponent from -1074 through 1023, with eight selected partners
  per exponent (16,784 pairs).
- 8,192 deterministic random finite binary64 bit-pattern pairs, both signs.
- Subnormal halfway cases with zero, positive and negative low components,
  even/odd parity, both signs, and the zero boundary.
- Normal/subnormal and finite/overflow boundaries, including the square of the
  largest finite value and square of the smallest positive value.
- Signed zeros and signed underflow.
- Invalid/nonfinite inputs and malformed scaled pairs.

Every product pair is compared to the exact rational product of the two
binary64 inputs; final conversion is compared with Python's rational-to-float
rounding, with explicit overflow and zero-sign checks. The fixtures are
registered in the full numerical-tooling runner; the metadata-only live path
does not execute them. This is finite regression coverage, not exhaustive
verification of every pair or a transferred R19 theorem.

## Remaining production design

| Layer | Required next step | Current scope |
|---|---|---|
| Arithmetic | Scaled addition, division, cancellation and accumulated error control | Only a product of two binary64 inputs and final pair conversion |
| Lower P/Q | Retain tiny shape/log argument information through logGamma1p, logP and expm1; bound remainder | Frozen v2 remains range-limited |
| Upper Q | Keep prefactor, shape and fraction exponent-separated until final rounding; certify total error | Existing fraction-only enclosure does not bound complete Q |
| Routing | Validate the existing x=a+1 seam and any new precision/domain transitions, including A>1 | No public routing change |
| Public families | Preserve positive ratios lost by Gamma scaling and Chi-square halving; validate inverse composition | #13/#14 remain distinct required work |
| Acceptance | Fix candidate first, freeze new disjoint validation, execute Excel and export exact-source evidence | Existing revealed sets cannot become new holdouts |

The 41 retained public Excel failures remain unresolved. No public `.bas`,
threshold, frozen v1/v2 protocol, observation, certification record or manifest
is changed by this study. PR #241/#242 remain draft and blocked.
