# Independent fraction enclosure for the bounded v2 study

This is an offline, post-hoc diagnostic on the revealed v2 set, not a new
holdout, a VBA kernel, a probability contract or a production acceptance.
Frozen v1/v2 numerical code and thresholds are unchanged.

## Mathematical construction

[DLMF 8.9.2](https://dlmf.nist.gov/8.9.E2) expresses
H = x^(-a) exp(x) Gamma(a,x) as a continued fraction. For 0<a<1 and x>0,
its coefficients are positive. An equivalent generalized fraction has b0=x,
a_(2k-1)=k-a, b_(2k-1)=1, a_(2k)=k, b_(2k)=x; H is its reciprocal.
All convergent numerators and denominators are positive. The determinant
identity for consecutive convergents alternates sign. Replacing the positive
remaining tail by zero or infinity gives the two consecutive convergents;
the fractional-linear map is monotone in that tail. Hence the infinite limit
is enclosed between them. At a=1, the first numerator vanishes and H=1/x.
These positivity arguments are not transferred to a>1.

`cf_enclosure.py` uses 120-digit Decimal contexts with directed floor/ceiling
rounding on each addition, multiplication, division and coefficient formation.
Inputs are exact Decimal conversions of binary64 values. The hull of two
consecutive interval convergents encloses the real H. The width is rounded
upward and the relative tolerance downward. A cap of 4096 raises an error;
no partial enclosure is returned as a converged result. This is exponent-safe
for the inspected study inputs, not an assertion about arbitrary Decimal or
VBA inputs. No high-precision code is introduced into the VBA runtime.

## Measurement

`check_v2_cf_enclosure.py` inspects the return-frame H of the unchanged frozen
v2 `upper` function; it does not substitute an alternative arithmetic path.
All 120 upper-branch rows in the already revealed 250-point protocol were
measured. The enclosure width is at most 2^-105 relative; maximum depth 737.
The candidate's maximum Lentz depth is 310. Exact Fraction comparisons bound
its represented hi+lo H error using both enclosure endpoints and the positive
lower endpoint. Maximum relative error upper bound: 6.91220972029415E-30;
all rows satisfy the explicitly diagnostic 2^-94 guard. Runtime about 2.86 s
in this environment. Full source hashes and intervals are retained in
`precision_v2_cf_enclosure.json`.

Three fixture methods cover outward arithmetic against exact rationals,
comparison with 150-digit incomplete-gamma evaluation, the a=1 identity,
invalid domains and deliberate cap exhaustion. The high-precision comparisons
are tests of implementation; the positivity argument supplies the enclosure.

This separates fraction truncation/arithmetic from prefactor error. It does
not enclose logGamma coefficients, log/exp evaluation, the upper prefactor,
P/Q final rounding or the lower branch. In particular it neither proves
adjacent-input monotonicity nor certifies the full public Gamma domain.
