# Private scaled Gamma VBA fidelity

2026-11-08; #24/#26, draft PR #242 remains open.

`scaled_gamma_vba.template` ports the complete bounded `scaled_gamma.py`
experiment into a private `SG_*` block in the test module. It uses Variant
triples (high, low, exponent) with private SGD operations derived from SD.
The frozen SD/PV2 primitives and their historical evidence remain unchanged.
No public VBA kernel or routing changes. Lower series, logGamma1p polynomial,
log/log1p, exp/expm1, upper forward convergence screen and backward fraction
reconstruction follow the Python operation order. Integer exponential reduction
uses ties-to-even Round. Explicit Double stores remain in the SD/PV2 primitives.

Domain: finite positive binary64 a <= 1 and x <= 32, including subnormal inputs.
Out-of-domain inputs and internal failures return False and clear P/Q. Series
screens are heuristic; this port does not establish accumulated-error bounds.

The generator `install_scaled_gamma_fidelity.py --write` reproduces 129 pairs:
69 in-domain revealed B6 points, 36 original development points and 24 adjacent
routing points (intentional overlaps retained). Each checks success and both
output bit patterns, including zero signs. Eight invalid cases check refusal
and cleared outputs. Eight three-point seams add four monotonicity assertions
each. Another 48 exact-dyadic normalization cases exercise signs, half-ULP
boundaries, ties-to-even and binade transitions. This adds 475 assertions to
7,715, giving 8,190 total.

The JSON references come from the Python model, not an independent oracle;
existing `test_scaled_gamma.py` separately compares revealed values to mpmath
and checks 160/320-digit rounding stability at the seams. None is a fresh
holdout. Old generator installations refuse the newer aggregate count rather
than overwrite it. Isolation fixtures prove the previous VBA suite and public
sources unchanged after removing the new SG block/call. Prior Excel artifacts
remain bound to their original sources and assertion totals.

Run the six generation/source/evidence tests with:

```bash
python benchmark/gamma_cf_study/coverage/test_scaled_gamma_vba_port.py
```

The full tooling runner has 72 commands. Actual host fidelity requires a run
on the exact port commit; local tests alone do not establish Excel execution.
The 41 public failures remain the expected baseline. Closing or merging either
PR is outside this checkpoint: full-domain/caller/inverse integration, error
analysis, fixed-candidate independent validation and release acceptance remain.

## Excel diagnosis and canonicalization

The initial run (42511055298, `da14d50`) had 8,142 assertions and 139 failures:
98 private Gamma failures plus the unchanged 41 public failures. Runs
42511468913 (`1a3854b`), 42511903287 (`974c661`) and 42512236419 (`6754fe7`)
reproduced those totals with increasingly precise diagnostics. Original artifact
bytes and SHA-256 digests for all four runs are retained in the adjacent JSON
files. Merely constructing the half-ULP constant from its bits did not fix it.

The bit diagnostic captured, for example, H=`3FE1FFFF:FFFFFFFE`,
L=`3C900000:00000001`, exponent 2, with the host sum still equal to H.
The low term exceeds half an ULP. This pair represents the same exact dyadic
value as the neighboring high component plus a negative residual, but host
addition alone did not produce that canonical pair. This observation does not
by itself identify the processor/runtime cause or certify general DD bounds.

Private `SGD_Canonicalize` selects the adjacent binary64 high component using
its words, then subtracts the exact one-ULP step from the low component. It
handles ties using the significand parity, the shorter step toward zero at
|H|=0.5, and binade renormalization. It refuses residuals beyond one adjacent
step. `SGD_Valid` uses those exact dyadic boundaries and parity instead of
relying on a repeated host sum. `SGD_Round` preserves the frozen SP conversion
graph with this private validation. No tolerance is widened. The 48 expected
pairs are independently generated with Python Fraction exact arithmetic.
Exact-source Excel run 42512900145 on `f0e8ef0` passes all 475 new private
assertions; all 41 public failure details remain identical to the earlier
7,715-assertion run. See [the result](SCALED_GAMMA_VBA_RESULT.md).
