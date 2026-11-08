# #26 — B1 independent feasibility checkpoint

Frozen on 2026-11-08 before evaluating any of the new holdout coordinates.
Baseline: `23a56f2ea60ab9391a93367431dabeefc51e5d37`.
This checkpoint screens a candidate in the source-derived model. It does not
authorize a production accuracy claim or close #26. No VBA or evidence binding
changes are part of this checkpoint.

## Candidate and comparator

B1 is exactly the existing `cf(a, x, eps=2**-52)` followed by
`backward(a, x, n)` in model.py at `5185d50`. The depth n is the first Lentz
iteration meeting that stop. There are no extra iterations, consecutive-hit
rules, alternative tolerances or fallback candidates. Keep the 100000 cap,
existing prefactor and final Double multiplication unchanged. Exhaustion,
non-finite results or invalid denominators fail; never score a partial value.
Compare with the original `cf(a, x)` and the tighter stop without backward
reevaluation. This is one candidate, not a search over stopping rules.

## Frozen fitting and independent coordinates

Fitting comprises the twelve committed X=2 ladder rows and the 72 already
inspected expanded fitting rows with shape <=0.5 in isolation.json. These are
not independent evidence. Keep all their original coordinates and results.

The independent set has 84 points. Shapes are 2E-6, 7E-6, 2E-5, 7E-5,
2E-4, 7E-4, 0.002, 0.007, 0.02, 0.07, 0.2 and 0.4. For each binary64 shape,
form s = binary64(a + 1); use s, nextafter(s,+infinity), and binary64(s*m)
for m in {1.5, 3, 6, 10, 14}. candidate_protocol.json stores the exact hexadecimal
binary64 inputs. No coordinate may overlap the fitting set. No existing global
holdout observations are read for design or candidate selection.

This window spans the seam and moderate upper tails. It is a feasibility
sample, not a newly declared production routing domain. Tiny/subnormal shapes,
subnormal outputs, the lower side of the series/CF seam and the moderate/large
shape public routes are outside this model checkpoint and remain mandatory
production-validation work if B1 survives. No claim of their validation follows
from passing these 84 points.

## Oracles and measurements

Use the exact binary64 coordinates as arbitrary-precision inputs. Evaluate
regularized upper incomplete gamma at 60 and 90 decimal digits. Independently
integrate exp(-t)*t**(a-1) from x to infinity and divide by Gamma(a), at 90 dps.
Require relative precision agreement <=1E-55 and quadrature agreement <=1E-70.
Unavailable, non-finite, nonpositive or unstable references invalidate the run;
they cannot be counted as passes. All selected reference outputs must be normal
binary64; unexpected underflow invalidates the checkpoint rather than dropping
a row. Record iteration counts, signed relative output errors, failing IDs,
maximum errors and source/input hashes. No averaging or rounded-value scoring.

## Prediction and decision

Prediction: B1 satisfies the unchanged Gamma survival bound, relative error
<=5E-15, on every fitting and every independent point, with no convergence
failures. This prediction is deliberately falsifiable despite the narrow
4.72E-15 fitting margin. The two comparator arms need not pass.

One B1 violation rejects this candidate as a sufficient repair. Preserve all
results and stop before VBA implementation; do not change this protocol, prune
points, relax the bound, or tune B1 on the revealed set. A later candidate must
use a new preregistration and a fresh independent set; this set then becomes
diagnostic evidence. A report verifier may pass for an honestly recorded
rejection: it verifies reproducibility, not numerical acceptance.

If B1 passes, that permits preparation of a separate production protocol only.
That protocol must freeze the routing domain and boundary behavior, equivalent
Gamma/Chi-square calls, P/Q consistency, inverse and Poisson controls,
non-convergence behavior, permanent regressions, fresh public-function holdout,
and exact-source Excel exports. Existing moderate/large-shape prefactor concerns
remain separately visible; neither their thresholds nor the global 5E-15
contract may be weakened to accommodate this candidate.
