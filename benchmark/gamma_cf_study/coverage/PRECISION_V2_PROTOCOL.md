# Precision v2 bounded feasibility protocol

The 489 revealed A<=1 points pass after completing two-Double arithmetic in
logGamma1p, the upper prefactor and CF, as well as the lower path. Earlier in
exploration, retaining the old rounded logGamma1p left two seam failures
(A=1E-100 and .01). This is fitting history, not independent validation.

Freeze v2 code, degree-72 coefficient artifact and 250 disjoint coordinates
before evaluating any new coordinate. The old v1 and its rejected evidence
remain unchanged. log(1) is exact. Coefficients of
g(a)=logGamma(1+a)/(a(a-1)) are expanded around .5, stored as two binary64
parts, and reproduce identically at 100/160 dps. Degree 72 is selected for
additional precision relative to the original degree 32, not fitted pointwise.

Require P/Q relative errors <=3E-15/5E-15, no wrong-direction adjacent pairs,
valid finite probabilities and 90/120-dps oracle plus quadrature agreement
<=1E-80. Any unavailable or unconverged value rejects. Do not change routing,
contract thresholds or frozen points after evaluation. CF cap remains 100000;
the study convergence test is 2^-100 with backward reevaluation in two-Double
arithmetic. This changed convergence rule is part of this frozen prototype,
not a production change.

The prototype has explicit limited ranges in JSON. Passing this screen does
not cover the whole public domain, VBA arithmetic fidelity, Excel regression,
source-bound exports or merge acceptance.
