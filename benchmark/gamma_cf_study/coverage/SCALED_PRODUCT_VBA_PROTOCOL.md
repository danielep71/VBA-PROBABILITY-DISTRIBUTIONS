# Private scaled-product VBA fidelity protocol

Candidate scope: #24/#26, PR #242, after `a7e506d`. Date: 2026-11-08.

This revision ports only the exponent-separated product of two binary64 inputs
and final scaled-pair rounding. It does not implement general scaled DD
addition/division or change a public Gamma route. The frozen v1/v2 studies,
public tests and their thresholds remain intact.

## References and input identity

`scaled_product_fidelity.json` contains 216 product pairs and 18 direct rounding
cases. Expected mantissas, residuals and final rounding are generated from
Python Fractions of exact binary64 inputs, independently of `scaled_product.py`.
They reproduce deterministically. Inputs and expected values are injected as
IEEE high/low 32-bit words, using `LSet` between two eight-byte private types.
This avoids decimal-literal loss for subnormals and preserves signed zero.
The representation assumes the little-endian Windows Excel host already used
by the runner; explicit word-layout assertions guard that assumption. No new
platform compatibility is claimed.

The 216 cases cover the original splitter-overflow and subnormal-residual
counterexamples, finite/overflow and normal/subnormal boundaries, both signs,
signed zeros, and 160 deterministic finite random bit-pattern pairs. The 18
rounding cases cover low-component decisions at halfway, even/odd parity and
the half-minimum-subnormal zero boundary. These are arithmetic regression
fixtures, not an independent Gamma holdout.

A separate VBA loop visits all 2,098 input exponents -1074..1023. For each
`2^e * 1.5`, it checks the exact pair `(0.75, 0, e+1)` and final bits against
the host's single binary64 multiplication. That hardware comparison checks
rounding/fidelity and is not used to generate probability references.

## Implementation boundaries

Private helpers are generated into `tests/M_STATS_PROBDIST_TEST.bas` from
`scaled_product_vba.template` by `install_scaled_fidelity.py --write`. The
Python binding fixture requires exact agreement between template/generator,
committed JSON, generated VBA block, suite invocation and assertion policy.

- Decompose inputs by exact power-of-two multiplication/division before the
  splitter. Explicit Double stores match the Python operation graph.
- Retain the integer exponent separately. `SP_Shift` is a bounded loop used
  only where the represented result is exact; it is not a general ldexp
  substitute for arbitrary subnormal final rounding.
- Round subnormal results in units of `2^-1074`, with the low component retained
  at ties. The rounded integer can reach `2^52`, so parity uses Double and Fix,
  never a VBA Long/Mod conversion.
- Return False on invalid inputs or overflow. Preserve signed underflow zero.
  This private Boolean result is not a new public worksheet error contract.

## Frozen expected counts and acceptance

The previous complete suite contains 2,447 assertions. This port adds:

| Group | Assertions |
|---|---:|
| 216 exact-rational product pairs, two checks each | 432 |
| Direct halfway rounding cases | 18 |
| 2,098 exponents, two checks each | 4,196 |
| Invalid inputs and bit-layout checks | 11 |
| **New assertions** | **4,657** |
| **Complete suite** | **7,104** |

Require compilation, execution of every new assertion and zero new failures.
The existing 41 public failures must remain visible with identical failure
labels/details. If these expectations hold, the total is 7,063 PASS / 41 FAIL,
which is still a failed release regression. Do not relabel it as green, export
grids, rebind manifests, close #24/#26, or merge #241/#242 on that basis.
Any port failure is retained and corrected in a separately identified commit;
this document is a protocol, not a record of an Excel run.

## Local checks

Seven scaled-product fixture methods and four oracle/generator binding methods
pass. The new fixture is registered in the full numerical-tooling runner and
its routing test. Static checks cannot establish VBA compilation or arithmetic
fidelity; those require the exact-source Excel run and retained job evidence.
