# Precision v2: arithmetic and Gamma research review

Review date: 2026-11-08. Candidate: e23e8007bf7707ff82661278f88bec730d1f2f6e
on PR #242. Research index: PR #273 at
867328d5102bf41e503d919d92e41b932d077c93. This note records a study review,
not acceptance, a changed contract, or permission to merge #241/#242.

## Sources and access depth

- R18: [author manuscript](https://ir.cwi.nl/pub/20433/20433B.pdf),
  sections 2.1–2.4, formulas 2.1–2.28 inspected.
- R19: [author-uploaded HAL v3 text](https://www.researchgate.net/publication/320313271_Tight_and_Rigorous_Error_Bounds_for_Basic_Building_Blocks_of_Double-Word_Arithmetic),
  algorithms 6, 12, 17 and surrounding assumptions inspected. HAL and publisher
  direct retrieval failed. This is not a complete proof replication.
  The author's publication list also identifies the 2022 Muller–Rideau follow-up,
  DOI 10.1145/3484514; its full proof was not accessible in this review, so no
  bound is transferred as a certified implementation guarantee.
- R20: [author paper](https://www.tuhh.de/ti3/paper/rump/OgRuOi05.pdf),
  algorithms 3.1–3.5 and theorem 3.4 inspected.
- R21: [revised author preprint](https://helios2.mi.parisdescartes.fr/~moisan/papers/download.php?file=gammainc_revised_preprint.pdf),
  normalization (5), representation and conversion discussion inspected.
- R22: [2000 technical report](https://www.davidhbailey.com/dhbpapers/quad-double.pdf),
  section 2, algorithms 3–7 and normalization definition inspected; this is the
  longer report, not a claim to have audited the conference version or QD code.

## Decisions from the source comparison

R18 selects the primary tail using alpha(x), rather than the library's x=a+1
series/CF seam. Its small-shape Q construction avoids subtraction of a rounded
P from one. Its D prefactor and the CF numerator k(a-k) are algebraically
related to the current route. Its auxiliary g for reciprocal Gamma is **not**
our logGamma1p/[a(a-1)] polynomial. Retain the cancellation-aware v2 logP/expm1
construction; do not copy the published region thresholds or coefficients.
In particular, x=a+1 is not a proof that the lower branch computes the smaller
tail. Future routing changes require their own adjacent-input seam validation.

R19 distinguishes accurate addition from sloppy addition and supplies different
bounds for different multiplication/division operation graphs. Our addition
resembles algorithm 6 but uses general TwoSum for both renormalizations.
Our multiplication has separate rounded cross-products, not algorithm 12's
FMA steps. Our division performs two residual corrections using our own
multiplication, not algorithm 17. Consequently none of those numerical bounds
is imported verbatim. Rounding, normalized operands and every intermediate's
exponent range must be checked independently.

R20 separates exact sum transformations from products whose residual may be
lost to underflow. Keep the general TwoSum ordering and test it with reversed
operands, cancellation and gradual underflow. The multiplication's operation
order must be compared with its own algorithm, rather than transferring a
bound merely because both use Dekker splitting. Accurate accumulation cannot
repair an already inaccurate prefactor or argument.

R21's G equals an unregularized selected integral multiplied by exp(x-a log x).
For the corresponding tail, conversion is G*exp(a log x-x-logGamma(a)).
This still needs an accurate, exponent-safe prefactor and stable complementation.
The normalization is a useful comparator and motivates retaining exponent
information until final rounding; it is not itself a replacement P/Q oracle.
No generalized-integral API or C implementation is adopted.

R22 assumes binary64 operations and ties-to-even and makes normalization
explicit. Its non-FMA Two-Prod ordering matches the current high-product
calculation. Keep explicit Double temporaries in VBA and test the actual host.
Four components are not adopted; two components also do not imply binary128's
exponent range or correctness guarantees.

## Exact-rational checks of the actual code

`test_precision_arithmetic_audit.py` adds four fixture methods, registered in
`benchmark/test_evidence_tools.py`. They check TwoSum exactness, cancellation,
and 2,048 deterministic normalized operand pairs for addition, multiplication,
division and output normalization. References are exact Python Fractions of
binary64 inputs; no mpmath rounding is involved. The 2^-98 relative guard is
an empirical regression threshold, not a theorem attributed to R19.
All four methods pass locally, including two deliberately retained limitations:

- product(2^1000,2^-1000) has exact value 1 but the unscaled splitter overflows;
- product(2^-1022,nextafter(0.5,+infinity)) loses an exact 2^-1075 residual.

These examples are outside the bounded Gamma candidate and prevent future
claims that its unscaled arithmetic is universally error-free. They do not
justify changing frozen v1/v2 code or weakening probability contracts. A
production extension needs exponent scaling and final subnormal rounding.

## Real Excel port evidence

Run [42473052390](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/42473052390)
used PR head e23e8007bf7707ff82661278f88bec730d1f2f6e, tested merge
5f54a43364c07f9e4b8de99acca0547743f89a03, Excel 16.0 build 20430, 64-bit.
Result: 2,447 assertions, 2,406 PASS, 41 FAIL. All 500 private v2 fidelity
assertions passed: both outputs exactly match the frozen Python results on
250 already revealed points. The failure list contains only the retained B6
and prior precision-counterexample assertions. The private port does not
replace production Gamma P/Q. Thus this verifies bounded arithmetic fidelity,
not independent validation, a production repair or full-domain acceptance.

The first port failed compilation because Scale was a reserved identifier.
Commit e23e800 renamed it to ScaleFactor in both generator template and test
module. The corrected port compiled and ran in the real Excel run above.
No grid export occurred; no evidence manifest was rebound.

## Remaining gates

- Exponent-safe full-domain routing, including subnormals and out-of-range
  exp arguments. The prototype's guards are not a public-domain policy.
- A convergence/remainder argument; a small Lentz update is not that argument.
- A production candidate resolving all retained failures and caller seams.
- Freeze a new independent set **after** that candidate is fixed; the 250
  fidelity points are already revealed and cannot be relabelled independent.
- Exact-source real Excel, fresh exports, CI and final review before merging.
