# Fresh validation of the fixed public Gamma candidate

The public source at 1440b180 is frozen before this cohort is evaluated.
`public_fresh_protocol.json` binds the three affected distribution modules,
wide model and oracle driver by LF-normalized SHA-256. It lists 206 forward
points on 14 shapes disjoint from the audited 99-shape / 46-file inventory,
and 24 inverse references on six of those shapes. Coordinates and scripts
must be committed before evaluation; do not tune the candidate to outcomes.

Forward reference evaluation uses direct lower and upper mpmath at 160/320
digits, with the previously frozen density-quadrature fallback on convergence
exceptions. Both precisions must round identically. Keep the original P/Q
3e-15 / 5e-15 normal-result screens and exact rounded bits for subnormal/zero
results. Check sorted-X monotonicity without an allowance. Record every row,
method, failure and interruption. A finite screen is not a global error bound.

Inverse references use high-precision Newton iteration in log X, driving the
smaller direct tail to its exact binary64 input target. SciPy supplies only
an initial guess. Require a high-precision residual check and identical rounded
references at 160/320 digits. This prepares references, not inverse-model PASS.
In Excel, exercise public Gamma and Chi-square inverses against those references
with the unchanged 2e-14 relative development criterion; retain failures.

The forward Python candidate is identical to the previously verified private
operation graph. Require actual public P/Q and Gamma/Chi-square CDF/SF output
bits to equal the frozen Python predictions. Public host composition is tested
separately; neither the earlier private port nor the new Python screen alone
certifies those callers. Keep the original public-composition and old 41
regression assertions active. Current exports and final review remain separate.
