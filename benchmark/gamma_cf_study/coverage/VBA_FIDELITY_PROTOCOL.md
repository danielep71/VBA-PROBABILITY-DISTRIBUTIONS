# Precision v2 VBA arithmetic fidelity protocol

Prospective execution specification: run the bounded test-only port against all
250 revealed Python v2 output pairs in precision_v2_screen.json. Require all
500 returned binary64 P/Q values to match exactly, and successful computation
for each pair. Failure status cannot pass an output comparison. hi;lo diagnostic
encoding retains failed output values. Expected suite size: 2447 (1947 previous
assertions plus 500 exact-output assertions).

No reference probabilities are regenerated for this comparison. This is runtime
fidelity evidence on revealed coordinates, not an independent accuracy screen.
The port stays private in the test module; no src module or public routing is
changed. The full-domain and convergence obligations in the research note stay
open even if all 500 comparisons pass. Existing 41 B6 failures remain blockers.

The diagnostic implementation uses Variant arrays to transport pairs, but every
load-bearing arithmetic operation is assigned to Double variables or returned
as Double. This representation has not been selected for production performance.
All 73 polynomial coefficients and ln(2) parts are generated as arithmetic sums
of 15-digit literals. No handwritten subnormal VBA literals are introduced.

Generate with install_v2_fidelity.py from precision_v2_vba.template. Frozen v2
Python algorithms, coefficient artifact and old rejected protocols are unchanged.
Do not turn a failed Excel run into a fresh observation export or manifest rebind.
