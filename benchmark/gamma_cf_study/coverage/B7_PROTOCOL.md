# B7 small-shape series compensation: preregistered model screen

Status: frozen before B7 independent evaluation or any B7 VBA change.
Parent B6 implementation: bf1f122. Coverage baseline: cb45b92.

## Diagnosis and candidate

The source-derived model at A=.5, x=predecessor(1.5) attributes the three-ulp
CDF overshoot to the ascending-series accumulation. Replacing only the prefactor
with a correctly rounded reference leaves the result unchanged; replacing the
sum or using Kahan accumulation removes the overshoot in this counterexample.
Exact Excel values must confirm model fidelity before a production patch.

Candidate B7 changes only the ascending-series addition for 0<A<=1 to explicit
Double-store Kahan accumulation. Recurrence terms, existing stopping criterion,
100000 cap, prefactor, route boundary X=A+1, B6 CF and the entire A>1 path stay
unchanged. No clamp, boundary shift, tolerance change or extra iterations.

Predicted movement: lower P improves at the A=.5 junction, restoring the two
new B6 monotonicity failures. Lower Q still uses 1-P and remains subject to #24;
large-A and inverse residual defects are not claimed fixed. This is a bounded
candidate screen, not permission to close #26 or #24.

## Fitting and untouched model validation

Fitting: the already-revealed 89-pair coverage design and the A=.5 junction.
Independent shape values, interpreted as exact binary64 values of these decimal
literals: 1e-72, 9e-9, .0071, .021, .083, .19, .34, .61, .88, .987.
For each shape, let s=binary64(1+a). Evaluate x=s*.17, s*.43, s*.81 (one
binary64 multiplication), plus 17 adjacent binary64 x values centred on s,
from eight predecessors through eight successors. Total 200 points, shape-major.
Do not inspect these outputs until this file is committed. Once evaluated they
are revealed and cannot be retuned or reused as untouched evidence.

References: direct regularized incomplete gamma at 90 and 120 decimal digits;
reject unstabilized values (relative disagreement above 1e-80). Independently
cross-check P by positive-integrand quadrature using
P=x^a/Gamma(a+1) * integral_0^1 exp(-x*u^(1/a)) du,
and Q by upper-tail quadrature; report convergence/disagreement rather than
silently scoring an unavailable reference. The transformed lower integral may
require shape-aware splitting near u=1. Numerical reference difficulties do
not count as candidate successes.

## Decision rules

- Require exact agreement of the model and observed counterexample before VBA.
- Report unchanged frozen P (3e-15 relative) and Q (5e-15 relative) scores for
  B6 and B7, including existing breaches. No missing/failed oracle is a PASS.
- Reject B7 if either newly introduced A=.5 seam failure remains, if a previously
  passing fitting/independent accuracy assertion becomes failing, or if B7
  introduces a new adjacent-X monotonicity failure against B6.
- Preserve and report every common failure and any changed magnitude. Passing
  this differential screen is not full numerical acceptance.
- If the screen fails, stop and report before selecting or tuning a replacement.
- If it passes, add permanent VBA seam regressions, obtain paired real-Excel
  evidence and new exact-source exports; CI/review remain mandatory. The
  revealed prior public holdout cannot certify a retuned implementation.
