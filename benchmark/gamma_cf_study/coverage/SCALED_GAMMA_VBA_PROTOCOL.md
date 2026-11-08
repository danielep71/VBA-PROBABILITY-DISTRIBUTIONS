# Private scaled Gamma VBA fidelity

2026-11-08; #24/#26, draft PR #242 remains open.

`scaled_gamma_vba.template` ports the complete bounded `scaled_gamma.py`
experiment into a private `SG_*` block in the test module. It uses Variant
triples (high, low, exponent) with the already host-checked SD operations.
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
each. This adds 427 assertions to 7,715, giving 8,142 total.

The JSON references come from the Python model, not an independent oracle;
existing `test_scaled_gamma.py` separately compares revealed values to mpmath
and checks 160/320-digit rounding stability at the seams. None is a fresh
holdout. Old generator installations refuse the newer aggregate count rather
than overwrite it. Isolation fixtures prove the previous VBA suite and public
sources unchanged after removing the new SG block/call. Prior Excel artifacts
remain bound to their original sources and assertion totals.

Run the four new generation/source tests with:

```bash
python benchmark/gamma_cf_study/coverage/test_scaled_gamma_vba_port.py
```

The full tooling runner has 72 commands. Actual host fidelity requires a run
on the exact port commit; local tests alone do not establish Excel execution.
The 41 public failures remain the expected baseline. Closing or merging either
PR is outside this checkpoint: full-domain/caller/inverse integration, error
analysis, fixed-candidate independent validation and release acceptance remain.
