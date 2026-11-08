# Extended Gamma: model screen and subnormal correction

2026-11-08; #24/#26; PR #242 remains open and draft.

## Frozen model evidence

Candidate/protocol commit `64430e088ac7ed64ef701049b8f7f46a996ebcd8` precedes
evaluation of the 220 coordinates in `extended_protocol.json`. All 220 pass
the declared bounded-model screen. Direct lower/upper mpmath references at
160/320 digits round identically; no wrong-direction step occurs across the
20 sorted shape groups. The shape bit patterns are disjoint from the explicitly
audited Gamma study inventory. The coordinates are now revealed.

Maximum relative errors on normal outputs: P `9.748004e-17`, Q `1.007328e-16`.
All 72 subnormal/zero reference outputs match the correctly rounded bits.
This finite evidence does not certify correct rounding throughout the domain.
The complete source-bound rows remain in `extended_screen.json`.

The preceding bounded-switch preflight at `57f7cb7` remains rejected and intact
in `routing/result.json`: two old-path accuracy failures just above a=1 and one
old-path Q monotonicity failure at x=predecessor(2), with no newly introduced
monotonicity failure. The extended model matches the oracle bits at all nine
coordinates involved in those three groups. Its private VBA tests include an
explicit control of that rejected switch and the repaired shape seam.

## Additional counterexample and separate v2 revision

After the 220-coordinate design was frozen, a separate development probe found
that v1 returns the minimum positive binary64 value for a=2, x=2^-537. The true
lower probability is strictly below 2^-1075 and must round to zero. The retained
v1 code is unchanged and a negative fixture reproduces its result. Passing the
independent finite screen did not detect this midpoint case.

`extended_gamma_v2.py` changes only a=2, 0<x<=2^-512. It rounds the exact scaled
product x*x/2 and returns Q=1. All other calls delegate unchanged to frozen v1.
No newly untouched holdout claim is made for v2; the 220 cases are replayed as
revealed checks, and this branch is verified separately with exact rationals.

### Why the product gives the correct rounding in this region

From P(2,x)=integral_0^x t*exp(-t) dt and 1-x<=exp(-t)<=1:

```
(x*x/2)*(1-x) < P(2,x) < x*x/2,  for 0<x<=2^-512.
```

Write the binary64 input as x=M*2^q, with integer M<2^53. The exact leading
product has dyadic spacing delta=2^(2q-1). The omitted correction is less than
(x*x/2)*x, and its ratio to delta is M*M*x < 2^106*2^-512 = 2^-406.
Subnormal rounding midpoints are integer multiples of delta in the only
relevant nonzero-output range. Consequently the correction cannot cross a
midpoint unless the leading product is exactly at one.

At such a midpoint x=J*2^-537 for odd J. Since J*J is 1 modulo 8, the lower
neighbor in units of 2^-1074 is even. Ties-to-even already selects that lower
neighbor, as does the true probability strictly below the midpoint. Smaller
leading products round to zero. Q rounds to one because P<=2^-1025.

The numerical branch uses the existing exact scaled binary64-product primitive,
private canonicalization and one final conversion; it does not compute x*x in
ordinary binary64 first. Rational fixtures check both enclosing endpoints and
the candidate on 8,291 values, including 96 midpoint neighbors and the domain
edges. No general bound for other shapes or the log/CF arithmetic is inferred.

## Private VBA checkpoint

The separate EG block covers 587 distinct pairs: prior B6/scaled cases, the
newly revealed 220 points, the routing preflight, domain edges and rational
shape-two probes. It adds 1,772 assertions: 1,761 success/P/Q-bit comparisons,
eight invalid/cleared-output checks and three routing-control assertions.
Aggregate policy: 9,962 assertions. Actual host results must be retained before
claiming Excel fidelity for this extension.

Generation/source/evidence fixtures preserve the complete preceding 8,190-test
module after removing only the new EG block/call, and every public source byte.
Earlier artifacts remain bound to their own counts and sources. No export,
manifest update or public routing change is included.

## Remaining scope

The extended domain is still bounded: finite positive a<=128, x<=1024. Public
routing outside and across that domain, accumulated-error control, inverse and
ratio composition, fresh public-candidate validation and final release checks
remain. The 41 public failures remain pending public integration. Neither the
220-point model pass nor private port fidelity closes the issues or PRs.
