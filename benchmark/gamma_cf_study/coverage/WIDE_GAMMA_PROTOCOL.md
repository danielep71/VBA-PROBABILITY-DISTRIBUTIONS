# Wide Gamma candidate and validation protocol

2026-11-08; #24/#26; development in draft PR #242, not public acceptance.

The frozen outer-routing preflight rejects the extended-v2/B6 composition:
13 pointwise failures, one newly introduced and two retained wrong-direction
shape steps. In particular Q drops at a=128 to its successor, x=1024, although
the B6 comparison increases. Original rows remain in
`routing_extended/result.json`. Those are model results, not Excel observations.

## Separate design

`wide_gamma.py` preserves the entire extended-v2 interior a<=128, x<=1024.
For a<=128 above that argument range it extends the same upper fraction and
uses a conservative endpoint bound at x>=1280. For larger a it replaces the
recurrence-based log-Gamma evaluation with the stable expression

```
d = x-a-a*log(x/a)
log(x^a exp(-x)/Gamma(a+1)) = -d-log(a)/2-log(2*pi)/2-StirlingError(a).
```

Near x=a, d is evaluated as a times the series for u-log(1+u), u=(x-a)/a,
starting at u^2/2. The first-order terms are never subtracted there. Scaled
two-Double arithmetic keeps large intermediate values out of binary64 range.

Ten rational Stirling terms are used only above a=128. For positive real a,
the analytic remainder is bounded by the first omitted term (DLMF
[5.11(i)](https://dlmf.nist.gov/5.11.i) and
[5.11(ii)](https://dlmf.nist.gov/5.11.ii)). Here its magnitude is below
(854513/63756)/128^21, hence below 2^-142. This bounds the mathematical
truncation only, not all floating-point operations. The log(sqrt(2*pi)) pair
is independently regenerable at 110/160 decimal digits.

The positive lower series uses t_0=1, t_n=t_(n-1)*x/(a+n). With
r=x/(a+n+1)<1, the uncomputed tail is at most t_n*r/(1-r) in exact arithmetic.
The candidate screens four times the computed tail against 2^-68 times the
computed sum. The factor four provides arithmetic headroom; it is not a
theorem for the complete scaled implementation. P and direct Q share log P;
Q is reconstructed by expm1 and doubling, never from rounded lower P.

Both wide lower and upper loops retain an explicit 100000-iteration limit.
The upper convergence screen remains heuristic. A stricter lower remainder
test can change reachability near the resource limit and must be measured,
not silently presented as identical to the old last-term criterion.

Chernoff's Gamma tail inequality gives the small tail <=exp(-d), on its
respective side of the mean. The numerical d>750 guard has a margin over
1075*log(2), but still requires arithmetic validation. For a<=128, x>=1280,
monotonicity of this bound and log(10)<2.5 give log Q<-832; returning P=1,Q=0
there is analytically safe in binary64. Exact endpoint x=0 returns P=0,Q=1.
Invalid nonfinite/negative inputs and detected numerical/cap failures refuse.

## Frozen screen

Freeze source hashes, coordinates and this protocol before evaluating the
candidate. The 14 new shapes are disjoint from the explicitly audited JSON
inventory in `wide_protocol.json`; this is not a claim about every imaginable
prior experiment. Keep the complete initial screen, including failures.

Use direct lower and upper mpmath references at 160/320 decimal digits. Require
identical reference rounding, P/Q relative errors <=3e-15/5e-15 for normal
outputs, exact reference bits for subnormal/zero outputs and no wrong-direction
adjacent step in each sorted shape group. Oracle exceptions reject the screen.
These are unchanged development thresholds, not a release contract expansion.

Separately replay the revealed 93 outer-routing points, earlier Gamma cases,
analytic underflow endpoints and resource-limit controls. Test public caller
and inverse composition, then validate a VBA port in real Excel. A passing
finite screen does not establish a general error bound or waive the existing
41 public failures, full regression, fresh exports and release gates.
