# Scaled two-Double arithmetic study

2026-11-08; continuation of #24/#26 and draft PR #242.

`scaled_dd.py` implements `(hi + lo) * 2^exponent`, with a normalized
binary64 pair and a separately bounded integer exponent. It adds addition,
subtraction, multiplication and division to the preceding single-product study.
The numerical path uses only binary64 operations and integer exponents.
Fractions occur only in the independent exact-rational test oracle.

## Contract and arithmetic graph

Nonzero values require `0.5 <= abs(hi) < 1`, `abs(lo) <= ulp(hi)/2`,
and `hi + lo == hi`. Zero has zero low component and exponent zero;
its high component retains the sign. Nonfinite inputs, malformed pairs,
noninteger exponents and exponents outside [-32767, 32767] are rejected.
This exponent range is an explicit study limit, not a public Gamma-domain rule.
Operations reject results whose normalized exponent leaves the study range.

Addition aligns the smaller mantissa at the larger exponent, runs the existing
`precision_probe.DD` addition graph, then renormalizes the pair. For gaps over
120, the smaller operand is omitted: each normalized magnitude is within
[0.5 - 2^-54, 1) times its exponent scale, so at gap 121 or larger the
relative omission error is less than 2^-119, also for opposite signs.
This local omission bound does not bound a sequence of operations or the
complete Gamma calculation. Alignment can underflow an extremely small low
component; the implementation does not claim exact addition.

Multiplication uses the existing DD product graph on bounded mantissas,
then adds exponents. Division uses its two residual corrections on bounded
mantissas, then subtracts exponents. Normalization uses general TwoSum before
rescaling. Large and tiny full values are never materialized in these graphs.
The previous frozen v2 implementation and its negative controls are unchanged.

`rounded()` rounds the represented pair using the previously tested conversion;
it does **not** guarantee correctly rounded results of the exact real operation.
In particular, arithmetic approximation can matter at a final halfway boundary.
No R19 error theorem is transferred to this graph. The R18/R19 methodological
review in `docs/research/PRECISION_V2_ARITHMETIC_REVIEW.md` still applies.

## Reproducible regression

Run from the repository root:

```bash
python benchmark/gamma_cf_study/coverage/test_scaled_dd.py
```

Seven test methods pass locally. They cover:

- 4,096 deterministic operand pairs and all four operations (16,384 rational
  comparisons). Half deliberately align exponents within 120 to exercise actual
  addition; the other half sample independent exponents from -2100 to 2100.
- Cancellation down to a minimum-subnormal low component, both signs, five
  exponent scales, and the normalized-mantissa exponent seam.
- Gaps around 53, 106 and the 120/121 alignment switch, including both signs
  and a separate exact-rational check of the omission bound.
- Three 256-step multiply/add sequences, checking each intermediate result.
- Products and quotients of maximum finite and minimum positive subnormal
  inputs, retaining out-of-binary64-range results until final conversion.
- Signed zeros, zero divisors, malformed inputs and exponent-range rejection.

The empirical relative-error guard is 2^-98; exact-zero references must remain
exact zero. This is finite, revealed arithmetic regression, not a proof of
98-bit accuracy throughout the domain and not an independent Gamma holdout.
The fixture is registered in the full numerical-tooling runner (69 commands),
and excluded from the metadata-only live path.

## Acceptance remains open

This commit changes no VBA source, public routing, thresholds, probability
observations, holdout or Excel certification record. The earlier exact-source
Excel result remains 7,104 assertions / 7,063 PASS / 41 public failures at
`86a7e47be621b32350fb343aa8afaa212eddddb9`. It does not certify this Python study.

Next work is a private VBA port with operation-by-operation host fidelity,
followed by complete lower/upper Gamma algorithms and accumulated error control.
Only a fixed candidate can precede a fresh disjoint probability validation set.
PRs #241/#242 remain draft; the release blockers are not closed by these tests.
