# Extended scaled Gamma candidate

2026-11-08; #24/#26. Development candidate; public routing remains unchanged.

The frozen bounded-routing preflight rejected the simple a<=1, x<=32 switch:
two accuracy failures remain just above a=1 and one wrong-direction Q step
remains at x=predecessor(2). None is a newly introduced monotonicity failure.
The model observations in `routing/result.json` are not Excel observations.

This separately identified candidate extends the numerical path to finite
positive a<=128 and x<=1024. It uses the frozen scaled-DD operations and
degree-72 polynomial. For a>1, logGamma(1+a) is reduced by recurrence to that
polynomial's interval, with scaled logarithms retaining both mantissa terms.
The positive lower series avoids cancellation in the alternating series at
larger x. The small-shape lower path is otherwise retained. Direct Q uses
expm1 on a halved log probability, followed by q(2z)=q(z)*(2-q(z)); it never
forms lower-region Q as 1-P. The upper path retains scaled backward CF.

For final lower log probabilities below -750, return P=0, Q=1: this lies well
below log(2^-1075), the binary64 rounding-to-zero boundary. This experimental
classification still needs accumulated-error control, like the rest of the
candidate; it is not justified merely by an exponential helper's range limit.
No general theorem is transferred from the literature. Convergence screens
remain heuristic and all explicit iteration caps remain fail-closed.

The nine revealed routing-failure coordinates now match direct oracle bits in
the Python model. The 89 already revealed B6 pairs pass the two-ULP development
screen. These are fitting checks, not independent validation.

Freeze `extended_gamma.py`, its numerical dependencies and the 220 coordinates
in `extended_protocol.json` before evaluating those coordinates. Their 20 shape
bit patterns are disjoint from the earlier Gamma study JSON inventory recorded
in the protocol. Disjointness is claimed against that explicit inventory only.
Cover subnormal inputs, extremely small lower tails, x=a+1 neighbors and upper
tails through x=1024. Record every outcome, with direct independent mpmath
P/Q references at 160/320 digits. Require identical rounded references, normal
P/Q relative errors <=3e-15/5e-15, correctly rounded subnormal/zero values, and
no wrong-direction P/Q step across each sorted shape group. Unavailable or
unconverged values reject; do not change the candidate or thresholds after
evaluation under this identity.

After a model pass, port it privately and verify exact Python/VBA fidelity on
the original plus newly revealed coordinates, retain failures and source
identity, and test the public-routing counterexamples in the same Excel run.
Full public-domain routing, fresh public-candidate validation, caller/inverse
composition, error analysis, exports and current CI/review remain separate gates.
Neither passing this finite screen nor private Excel fidelity closes #24/#26.
