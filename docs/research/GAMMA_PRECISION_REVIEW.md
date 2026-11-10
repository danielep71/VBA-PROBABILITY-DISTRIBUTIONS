# Gamma precision: research applicability and remaining gates

Related: #24, #26, #241, #242; research catalogue in
[PR #267](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/pull/267),
reviewed at `85e799522e7e48da5190b36f1db5025d69bf6c14`.
This note does not amend any frozen candidate, contract or holdout.

## Sources actually inspected

**R11, Johansson.** Full text of the 2021 preprint
[arXiv:2109.08392v1](https://arxiv.org/pdf/2109.08392), particularly sections
2.2–2.3, Algorithm 4 and section 5.1; publication DOI
[10.5206/mt.v3i1.14591](https://doi.org/10.5206/mt.v3i1.14591), 2023.
The publisher PDF was unavailable through the retrieval route; the inspected
preprint is not represented as an audited copy of the final publication.
Relevant principles: separate truncation and arithmetic error, account for
conditioning near logGamma zeros, and use interval/ball propagation for a
certified result. The reciprocal-Gamma coefficient theorem is not a theorem
about this project's factored logGamma polynomial. No formula is transplanted
on that basis. Agreement of coefficients at 100/160 dps is useful evidence,
but not an outward-rounded coefficient enclosure.

**R10, Pinelis.** Full publisher PDF
[10.7153/mia-2020-23-95](https://files.ele-math.com/articles/mia-23-95.pdf),
Theorem 1.1, Proposition 2.8 and their supporting proof sections. For 0<a<1,
(1.3)/(1.5) provide an upper bound on unregularized Gamma(a,x), while (2.12)
provides a lower bound. Both require division by positive Gamma(a) to bound Q.
At a=1 use the exponential identity separately. Inequality directions outside
this shape range must not be inferred from these formulas. The numerical
implementation below is supplementary and not an interval certificate.

## Stable normalization of the supplementary bounds

For 0<a<1 and x>0, let L=logGamma(1+a), b=exp(L/(a-1)) and
C=exp(-x+a log(x)-L). Then

```
Q_lower = C*a*(x+3-a)/(x*x+(4-2*a)*x+(a-1)*(a-2))
Q_upper = C*expm1(a*log1p(b/x))/b
```

The second expression avoids subtracting two nearly equal powers and cancels
Gamma(a)=Gamma(1+a)/a algebraically. Tiny-a reference L is computed without
forming rounded 1+a. `research_checks.py` evaluates both at 90/120 dps on
739 already revealed A<=1 points: all references lie within the bounds, with
stable evaluations. This does not create independent holdout evidence.
The negative fixture demonstrates the limit: both known erroneous A=.39 B6
values lie within these bounds, despite their wrong-direction one-ulp jump.
A deliberately large error is detected. Binary64 conversion margins do not
make the underlying mpmath evaluation an outward-rounded enclosure.

## Project-derived truncation bound for the exact degree-72 polynomial

This derivation is specific to our polynomial; it is not attributed as a
Johansson theorem. Let f(a)=logGamma(1+a), g(a)=f(a)/(a(a-1)), with removable
endpoint values. Since f(0)=f(1)=0, g(a) is the second divided difference
f[0,1,a]. Its simplex integral representation integrates f'' over a triangle
of area 1/2. On |a-1/2|<=7/5, every argument w of the trigamma function in
that representation has Re(w)>=1/10. Therefore

```
|trigamma(w)| <= sum(k>=0) 1/(k+1/10)^2
              <= 100 + integral(0,infinity) dt/(t+1/10)^2 = 110,
|g(a)| <= 55.
```

The disk remains away from the first nonremovable singularity at a=-1.
Cauchy's coefficient estimate and |a-1/2|<=1/2 give the exact-real remainder

```
B = 55*(5/14)^73/(1-5/14) < 1.949E-31,
|logGamma(1+a) - a*(a-1)*g_72(a)| <= a*(1-a)*B   (0<=a<=1).
```

The rational B is reproduced exactly with Python Fraction. This excludes
coefficient rounding, Horner roundoff and all subsequent arithmetic errors.
It also excludes any claim of correctly rounded P or Q.

For the lower region 0<x<a+1, combine P<=x^a/Gamma(1+a) with the lower bound
above. The perturbation in Q caused *only* by this logGamma truncation is at
most `7*exp(2)*B*exp(B/4)` in relative terms: use x<=2, denominator polynomial
<=14, x+3-a>=2, and |exp(delta)-1|<=|delta|exp(|delta|).
This separates a small analytic approximation contribution from the still
unbounded floating-point contributions. Endpoint identities remain exact in
the factored representation.

## Full-domain routing contract: target, not implemented

The public Gamma wrappers accept finite X, positive finite scale and
0<shape<1E100; that magnitude guard is not an accuracy/convergence guarantee.
The internal standardized X may be zero after underflow or overflow on division.
Those existing wrapper semantics and Chi-square's distinct half-X policy remain.

| Region/condition | Required route or behavior | Present evidence/gap |
| --- | --- | --- |
| X<=0 after existing wrapper handling | P=0, Q=1 | Preserve existing support semantics |
| Standardized-ratio overflow | Existing Gamma limit P=1, Q=0 | Preserve wrapper behavior; not a kernel failure |
| 0<a<=1, 0<x<a+1 | Shared accurate logP; direct small Q via expm1 | v2 bounded feasibility passes; tiny inputs outside its range are not covered |
| 0<a<=1, x>=a+1 | Direct upper Q; complementary P | v2 bounded feasibility passes only for its explicit ranges |
| a>1 | Existing series/Loader/CF path until separately diagnosed | Existing A=100 and other seam failures remain; no automatic certification |
| Subnormal a, x or result | Scale intermediate arithmetic; round only the final probability | Two-Double product identities need protection from underflow; not implemented |
| Exp argument beyond prototype [-64,8] | Exponent-separated evaluation and explicit underflow classification | Do not use an arbitrary prototype range cutoff as probability zero |
| Nonconvergence/invalid arithmetic | False with diagnostic, public #NUM!, then successful status recovery | Cap 100000 is retained; new convergence criterion requires error control |

In particular, a bare Lentz delta tolerance, even 2^-100, is not a rigorous
bound on the remaining CF. It must be checked independently, including depth
extension and any available remainder estimate. Supplementary Pinelis bounds
are too wide to replace this check. No clamp may turn an invalid intermediate
or unconverged result into a successful probability.

Before production integration, the coefficient/evaluation error budget, full
range handling and VBA arithmetic fidelity must be established jointly. The
250-point v2 set is revealed. A revised algorithm needs a newly frozen disjoint
validation set, and then complete Excel regression, callers, exports, CI and
review. The current B6 result remains 1906/1947; neither paper closes a gate.
