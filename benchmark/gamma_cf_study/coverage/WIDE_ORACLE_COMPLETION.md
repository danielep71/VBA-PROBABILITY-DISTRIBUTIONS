# Wide-screen oracle completion

2026-11-08; no candidate, coordinates or thresholds change.

The original invocation of the frozen `screen_wide_gamma.py` aborts at index
171 (zero based), a=65521.125, x=131042.25. mpmath's direct lower-gamma
hypergeometric evaluation raises `NoConvergence`, which is an `Exception` but
not a `RuntimeError`. The original log is retained verbatim in
`wide_initial_attempt.txt`; that invocation is INCOMPLETE, not PASS. No complete
result file was produced, and no successful rows are invented for it.

Before re-evaluation, freeze the completion harness in `screen_wide_complete.py`.
It retains the unchanged candidate/protocol hash checks, first tries the same
direct lower/upper evaluations at 160/320 digits, and captures all ordinary
exceptions as failures. Only for lower-oracle `NoConvergence` at a>128,x>a, it
integrates the normalized Gamma density directly from 0 to x, using
y=a+t*sqrt(a) and breakpoints around the mode. This does not subtract Q from
one and does not call the candidate's Stirling/deviance or series code. Q
continues to use direct mpmath upper Gamma. Record each oracle method per row.

This is completion for the same fixed candidate and frozen shape-disjoint
sample, not a new second holdout. References must still have stable binary64
rounding at both precisions; exceptions and numerical/monotonicity breaches
still reject. Preserve the original script, protocol, candidate and incomplete
attempt. A passing completion would remain finite model evidence, not Excel
or public acceptance, and would not erase the oracle failure.
