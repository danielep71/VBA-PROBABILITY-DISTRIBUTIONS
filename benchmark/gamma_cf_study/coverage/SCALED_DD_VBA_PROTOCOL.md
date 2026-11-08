# Private scaled two-Double VBA fidelity

2026-11-08; #24/#26, draft PR #242.

The generated `SD_*` block in the private test module ports `scaled_dd.py`.
It reuses the unchanged bounded `PV2_*` mantissa graphs and `SP_*` bit helpers.
No public source or routing changes. `install_scaled_dd_fidelity.py --write`
reproduces the block and `scaled_dd_fidelity.json`; old generated SP fixtures
remain frozen. Rebuilding the earlier standalone SP installation intentionally
refuses the newer aggregate assertion count.

The 611 new assertions compare all four operations to the Python model's high
word (including signed zero), numeric low word and exponent, or check refusal
with cleared outputs. Low-word zero signs are not contractual. There are 588
successful cases and 23 refusal cases. The total policy count is 7,715.

Coverage includes seeded operands, nearby and widely separated exponents,
cancellation with minimum-subnormal low components, exponent seams, alignment
gaps 120/121, signed zeros, nonfinite inputs, malformed pairs, zero divisors,
unknown operations and exponent-range overflow. An independent Fraction test
checks every successful stored pair against the exact rational operation with
the unchanged empirical 2^-98 relative-error guard (exact zero required).
This is revealed arithmetic regression, not independent Gamma acceptance.

Alignment computes a power-of-two factor exactly, then multiplies the low word
once. Repeated halving of that word could double-round subnormals and would not
match `math.ldexp`. Normalization similarly uses a single exact scale factor;
when the TwoSum high word is subnormal, its nonoverlapping low word is zero.
All exponents and gaps fit VBA Long within the explicit +/-32767 study range.
Try failures clear all output fields. Actual host fidelity still requires an
Excel run on the exact source commit; Python tests and static checks alone do
not establish this. The previous 7,104-assertion evidence is not rebound.

Local commands:

```bash
python benchmark/gamma_cf_study/coverage/test_scaled_dd.py
python benchmark/gamma_cf_study/coverage/test_scaled_dd_vba_port.py
python benchmark/gamma_cf_study/coverage/test_scaled_vba_port.py
```

The tests pin generation, independent rational comparisons, refusal coverage,
policy counts and byte-equivalent restoration of the earlier VBA suite after
removing the new block/call. The full tooling runner has 70 registered commands.
The expected public failure baseline remains 41; no release acceptance claim
is made even if all new arithmetic fixtures pass.
