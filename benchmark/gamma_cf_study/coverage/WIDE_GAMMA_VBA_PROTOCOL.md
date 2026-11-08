# Private wide Gamma port: frozen host protocol

This port extends only the test module. The six production VBA modules are
byte-for-byte unchanged from the extended-Gamma Excel checkpoint 924ff800.
It therefore cannot resolve the 41 public regression failures by itself.

`wide_complete_v2_screen.json` completes the original 198-point, 14-shape
screen: no accuracy or monotonicity violations. The initial interrupted run
and the four oracle errors from its first completion are retained unchanged.
`wide_outer_replay.json` passes all 93 previously revealed boundary points,
including the composition rejected in `routing_extended/result.json`.
These remain Python model observations until separately reproduced in Excel.

Generate the private WG_ routines with
`python benchmark/gamma_cf_study/coverage/install_wide_gamma_fidelity.py --write`.
The dataset has 290 unique pairs (wide screen, outer replay and exact x=0
endpoints), seven invalid inputs, four routing controls and five cap/recovery
checks: 886 added assertions; aggregate policy 10848 = 9962 + 886.

Require bit-for-bit Python/VBA agreement on both outputs of every pair.
Reproduce the rejected composition in real Excel at a=128 and its binary64
successor, x=1024, then require both new wide outputs to be monotone. Test the
actual 100000-iteration upper and lower caps at a=1e14, clearing both outputs
and reporting the applicable cap. A subsequent ordinary call must recover.
Do not replace a natural cap with a test-only lowered iteration limit.

The retained aggregate baseline is 9962 total, 9921 pass, 41 fail. If all WG_
checks pass and no old outcome changes, the expected new aggregate is 10848
total, 10807 pass, 41 fail. Record actual totals, failed labels, import/compile/
cleanup, exact candidate SHA, seven imported-source hashes, run/artifact IDs,
ZIP digest and original artifact bytes. A red aggregate is not certification.
No benchmark exports or release eligibility may be inferred from private
fidelity; public/caller/inverse integration and fresh exports remain pending.

`wide_resource_probe.json` separately records development-only Python cost:
a=5e7 at its mean needs 67920 lower iterations, a=1e8 needs 96404. The positive
lower remainder screen is deliberately stricter than B6. These observations
are neither a global domain guarantee nor an Excel performance measurement.

Earlier source-bound evidence stays attached to its historical commit. The
replay fixtures remove only the exactly generated WG_ addition when proving
that the earlier test module is preserved; they do not claim that the old
Excel run imported the new module. The live aggregate count includes all 886
new assertions and is independently checked by the generator fixture.
