# Annotated bibliography

[Research index](README.md) · [Issue map](ISSUE_MAP.md)

Sources checked on 2026-11-08. Dates below distinguish first online publication
from the journal volume year. These are research candidates, not adopted code.

## R01 — Incomplete beta using Deuflhard's algorithm

- **Authors:** Toshio Yoshida and Yoshinori Adachi.
- **Title:** Computation of incomplete beta function ratios I_x(a,b) with Deuflhard's algorithm.
- **Publication:** Numerical Algorithms 97, 373–390 (2024); first online 2023-12-07.
- **Primary source:** [DOI 10.1007/s11075-023-01707-6](https://doi.org/10.1007/s11075-023-01707-6).
- **Access depth:** publisher abstract, bibliographic record and exposed appendix; full subscription text not reviewed.

The authors analyze recurrence errors and use Deuflhard's method to avoid
repeated fixed-length Miller recurrences, reporting fewer iterations than
Gautschi's method for equivalent accuracy. That comparator is not this VBA library.

**Project hypothesis:** an alternative recurrence may help selected incomplete-beta
regimes or provide a structurally different cross-check. Compare against the
current Lentz route only after obtaining enough algorithm detail to reproduce it.
The diagnosed balanced-beta prefactor problem is not automatically repaired by
changing the recurrence. Track under #175; retain #163/#173 as defect owners.

## R02 — Regularized incomplete beta by combined methods

- **Authors:** Vera Egorova, Amparo Gil, Javier Segura and Nico Temme.
- **Title:** Computation of the regularized incomplete Beta function.
- **Publication:** Dolomites Research Notes on Approximation 16(3), 10–16 (2023).
- **Primary source:** [Publisher page and paper](https://drna.padovauniversitypress.it/2023/3/3); DOI 10.14658/PUPJ-DRNA-2023-3-3.
- **Access depth:** publisher abstract and metadata; linked PDF not audited equation by equation.

The method combines continued fractions with series and large-parameter
asymptotics. The authors report better accuracy than MATLAB betainc over their
tested parameter ranges; this does not establish superiority over current MATLAB
versions or our supported domain.

**Project hypothesis:** compare regime selection and prefactor treatment with
the existing implementation. Measure balanced/unbalanced shapes, endpoint tails
and dispatch seams. An alternate kernel remains conditional on evidence under
#175, not a prerequisite for closing the current defects.

## R03 — Quantile error and characteristic-function inversion

- **Author:** Gero Junike.
- **Title:** Precise quantile function estimation from the characteristic function.
- **Publication:** Statistics & Probability Letters 222, 110395 (July 2025).
- **Primary sources:** [DOI 10.1016/j.spl.2025.110395](https://doi.org/10.1016/j.spl.2025.110395), [author preprint](https://arxiv.org/abs/2502.13537), [institutional published-version record](https://epub.ub.uni-muenchen.de/129631/).
- **Access depth:** publisher and institutional abstracts/metadata; no full proof or code audit.

The paper provides quantile error bounds and a COS method, with numerical
examples on normal-inverse Gaussian and tempered stable distributions. The
reported exponential convergence assumes smooth densities and semi-heavy tails;
it must not be transferred indiscriminately to Student-t or every current family.

**Project hypothesis:** strengthen inverse diagnostics by separating probability
residual from quantile error and local conditioning. Review existing contracts
under #7 and apply the diagnostic to #213/#35 without replacing their frozen
protocols. COS implementation is deferred; no new distribution is approved here.

## R04 — Gaussian quantiles from log probabilities

- **Author:** Martin Mächler, ETH Zurich.
- **Title:** Asymptotic Tail Formulas For Gaussian Quantiles.
- **Publication type/date:** technical note dated November 2022; consulted PDF typeset 2025-12-16. The typesetting date is not a new 2025 journal publication.
- **Primary source:** [Author technical note distributed with DPQ](https://r-forge.r-universe.dev/DPQ/doc/qnorm-asymp.pdf).
- **Access depth:** PDF abstract and selected explanatory text; not a complete derivation or code audit.

The note develops asymptotic approximations for Gaussian quantiles with
log-probability arguments in extremely deep tails where ordinary probability
inputs underflow. It discusses R/DPQ implementations and their region selection.

**Project hypothesis:** use as a design reference for #194's log-space operations.
A function returning log probabilities is different from an inverse accepting
log probabilities. Assess the latter explicitly; if it needs a new kernel or
exceeds #194's existing scope, split a future proposal before implementation.
Do not compute `Exp(log_p)` before inversion when that discards the input.
Existing inverse-survival functionality is not evidence that log-input inversion
is already implemented or certified.

## R05 — Noncentral beta CDF

- **Authors:** Vera Egorova, Amparo Gil, Javier Segura and Nico M. Temme.
- **Title:** A numerical algorithm for the computation of the noncentral beta distribution function.
- **Publication:** Numerical Algorithms 99, 1791–1804 (2025); first online 2024-09-10.
- **Primary source:** [DOI 10.1007/s11075-024-01931-8](https://doi.org/10.1007/s11075-024-01931-8).
- **Access depth:** open publisher text, abstract and selected algorithm/series sections; accompanying MATLAB implementation not audited.

The paper supplies a noncentral-beta algorithm and MATLAB implementation, using
representations suited to different parameter regions and complementary tails.

**Project hypothesis:** candidate numerical foundation for the noncentral F
work in #197 through the beta/F transformation. Benchmark central limits,
large noncentrality, transition regions, convergence and direct upper tails.
Independent reference algorithms must be identified before treating agreement
with the authors' implementation as validation.

## R06 — Noncentral beta inversion

- **Authors:** Vera Egorova, Amparo Gil, Javier Segura and Nico M. Temme.
- **Title:** Algorithms for the inversion of the noncentral beta distribution function.
- **Publication:** Numerical Algorithms 103, 1549–1569 (2026); first online 2025-12-02.
- **Primary source:** [DOI 10.1007/s11075-025-02274-8](https://doi.org/10.1007/s11075-025-02274-8).
- **Access depth:** open publisher text, abstract and selected introduction/formula sections; no complete implementation audit.

Two distinct inversion problems are addressed: finding the quantile and finding
the noncentrality parameter. The algorithms combine asymptotic initialization
and iterative solution. The paper gives the relationship to the noncentral F
distribution and reports implementation comparisons in its tested domains.

**Project hypothesis:** support noncentral F quantile design within #197.
A public noncentral-beta API or noncentrality estimator requires an explicit
inventory/scope decision; neither is silently added to the existing roadmap.
R05 supplies a companion forward-function candidate, not an independent oracle.

## R07 — Negative-binomial fitting near the Poisson limit

- **Authors:** Yingying Yang, Niloufar Dousti Mousavi, Zhou Yu and Jie Yang.
- **Title:** From Poisson observations to fitted negative binomial distribution.
- **Publication:** Statistical Papers 67, article 62 (2026); published 2026-05-14.
- **Primary source:** [DOI 10.1007/s00362-026-01842-x](https://doi.org/10.1007/s00362-026-01842-x).
- **Access depth:** open publisher abstract and selected sections on APMA, likelihood and parameterization; supplemental R code not audited.

The paper studies unstable negative-binomial maximum-likelihood fitting for data
near the Poisson limit. It proposes adaptive profile maximization and an extended
parameterization containing Poisson distributions as a special case.

**Project hypothesis:** assess fitting as a future capability under #266. This
is not a replacement for the present probability kernels. Separate optimization
and statistical-estimation error from PMF/CDF error; review real-valued shape
parameters versus the library's current API restrictions. Numerical limit tests
may be useful even if fitting is ultimately deferred.

## Review boundary

No article above has been fully replicated in this review. Access, license,
algorithm details, versions and errata must be rechecked before implementation.
Research relevance is not a numerical accuracy contract. Subsequent study results
must link to immutable source-bound evidence rather than overwrite these initial
hypotheses as though they had always been established.

## R08 — Asymptotic inversion of discrete distributions

- **Authors:** Amparo Gil, Javier Segura and Nico M. Temme.
- **Title:** Asymptotic inversion of the binomial and negative binomial cumulative distribution functions.
- **Publication:** Electronic Transactions on Numerical Analysis 52, 270–280 (2020); preprint 2020-01-12.
- **Primary sources:** [Publication or author manuscript](https://doi.org/10.1553/etna_vol52s270); [Supporting primary source](https://arxiv.org/abs/2001.03953).
- **Access depth:** Preprint introduction and selected formula passages; not a complete implementation audit.

**Project hypothesis and limitations:** Asymptotic continuous inversion can supply starting values for discrete quantile searches. Evaluate total cost after correcting to the exact discrete generalized inverse: F(k-1) < p <= F(k) for interior probabilities. Approximate seeds alone do not satisfy this condition. Compare large parameters, tails and existing fallbacks under #269; this is distinct from fitting in #266.

## R09 — Student-t asymptotic representation and inversion

- **Authors:** Amparo Gil, Javier Segura and Nico M. Temme.
- **Title:** A new asymptotic representation and inversion method for the Student’s t distribution.
- **Publication:** Integral Transforms and Special Functions 33(8), 597–608 (2022); online 2021-12-05; preprint 2020.
- **Primary sources:** [Publication or author manuscript](https://doi.org/10.1080/10652469.2021.2007906); [Supporting primary source](https://arxiv.org/abs/2012.09789).
- **Access depth:** Publisher abstract/metadata and preprint landing page; derivation not audited.

**Project hypothesis and limitations:** The complementary-error-function representation offers an asymptotic route to Student-t probabilities and inverses. Under #213, first establish how it differs from the existing large-df approximation, then compare dispatch seams, quantile error and probability residual. It is a candidate diagnostic or alternative, not evidence that the current seam defect is fixed.

## R10 — Bounds for the incomplete gamma function

- **Authors:** Iosif Pinelis.
- **Title:** Exact lower and upper bounds on the incomplete gamma function.
- **Publication:** Mathematical Inequalities & Applications 23(4), 1261–1278 (2020); preprint 2020-05-13.
- **Primary sources:** [Publication or author manuscript](https://doi.org/10.7153/mia-2020-23-95); [Supporting primary source](https://arxiv.org/abs/2005.06384).
- **Access depth:** Abstract and institutional bibliographic metadata; proof not audited.

**Project hypothesis and limitations:** The paper develops bounds for the upper incomplete gamma function, including asymptotically exact bounds. For positive shape, transfer to regularized Q requires stable normalization by Gamma(a). Check theorem domains and rounding before using bounds as diagnostics in #24/#26 or the audit #7. A wide bound is not a 1e-14 accuracy certificate; rigorous floating-point enclosures require controlled rounding.

## R11 — Arbitrary-precision gamma reference methods

- **Authors:** Fredrik Johansson.
- **Title:** Arbitrary-precision computation of the gamma function.
- **Publication:** Maple Transactions 3(1), article 14591 (2023); published 2023-02-01; preprint 2021-09-17.
- **Primary sources:** [Publication or author manuscript](https://doi.org/10.5206/mt.v3i1.14591); [Supporting primary source](https://arxiv.org/abs/2109.08392).
- **Access depth:** Publisher/preprint abstracts and bibliographic metadata; algorithms not audited.

**Project hypothesis and limitations:** The survey addresses arbitrary-precision gamma and related functions with rigorous error control. Its possible role is external reference construction for Gamma/logGamma and prefactor diagnostics (#7/#175), not arbitrary-precision arithmetic inside the VBA runtime. Record precision, exact binary64 inputs, enclosure widths and algorithm lineage; two wrappers around the same implementation do not provide independent evidence.

## R12 — Third-order central beta inversion

- **Authors:** Dhivya Prabhu K, Sanjeev Singh and Antony Vijesh V.
- **Title:** A third-order iterative algorithm for inversion of cumulative central beta distribution.
- **Publication:** Numerical Algorithms 94, 1331–1353 (2023); online 2023-05-10.
- **Primary sources:** [Publication or author manuscript](https://doi.org/10.1007/s11075-023-01537-6).
- **Access depth:** Subscription abstract and metadata only; full convergence proof not reviewed.

**Project hypothesis and limitations:** The proposed third-order method has stated sufficient nonlocal convergence conditions. Compare those conditions and total evaluation cost with safeguarded Newton/bisection before any adoption. Under #35, an alternative beta solver does not by itself repair cancellation in reconstructing the F inverse from a beta pair. Preserve bracket, convergence-failure and output-error requirements.

## R13 — Accurate Poisson-binomial tails

- **Authors:** Noah Peres, Andrew Lee and Uri Keich.
- **Title:** Exactly computing the tail of the Poisson-Binomial Distribution.
- **Publication:** Preprint 2020-04-16; ACM Transactions on Mathematical Software 47(4), article 35 (2021).
- **Primary sources:** [Publication or author manuscript](https://arxiv.org/abs/2004.07429); [Supporting primary source](https://stat.ethz.ch/CRAN/web/packages/ShiftConvolvePoibin/refman/ShiftConvolvePoibin.html).
- **Access depth:** Preprint abstract and official package documentation; implementation not audited.

**Project hypothesis and limitations:** ShiftConvolvePoibin combines exponential tilting with FFT convolution to address relative errors in tiny tails. Evaluate direct recursion for moderate n before taking on a VBA FFT implementation. The distribution counts independent heterogeneous Bernoulli outcomes, not correlated defaults or weighted losses. Feasibility #270 must verify tail error, limiting cases and resource limits; exact in the title does not establish correct rounding for every input.

## R14 — Bivariate normal evaluation through Owen’s T

- **Authors:** Janez Komelj.
- **Title:** The Bivariate Normal Integral via Owen’s T Function as a Modified Euler’s Arctangent Series.
- **Publication:** American Journal of Computational Mathematics 13(4), 476–504 (2023); revised preprint v2 2025-06-19.
- **Primary sources:** [Publication or author manuscript](https://doi.org/10.4236/ajcm.2023.134026); [Supporting primary source](https://arxiv.org/abs/2312.00011).
- **Access depth:** Abstract and package documentation; no complete derivation/code audit.

**Project hypothesis and limitations:** The modified series is a candidate alongside Genz and Drezner-Wesolowsky in #196. Use the corrected 2025 preprint when evaluating the method. The reported maximum absolute error of 3.45e-16 over one million tested triples is an empirical absolute-error result, not a relative-tail guarantee. Independently test extreme thresholds, correlation approaching both ±1 limits and degenerate cases.

## R15 — Gamma median approximations and candidate bounds

- **Authors:** Richard F. Lyon.
- **Title:** On closed-form tight bounds and approximations for the median of a gamma distribution.
- **Publication:** PLOS ONE 16(5), e0251626; published 2021-05-13.
- **Primary sources:** [Publication or author manuscript](https://doi.org/10.1371/journal.pone.0251626); [Supporting primary source](https://pmc.ncbi.nlm.nih.gov/articles/PMC8118309/).
- **Access depth:** Abstract and bibliographic metadata; complete derivations not reviewed.

**Project hypothesis and limitations:** The paper supplies median approximations and bounds, with conjectures that must be distinguished from proved results. Read together with R16. For #14/#7, these can provide supplemental shape-dependent checks at p=0.5, not references for all gamma quantiles. Evaluate small-shape expressions in log space where necessary.

## R16 — Proved gamma median bounds

- **Authors:** Richard F. Lyon.
- **Title:** Tight bounds for the median of a gamma distribution.
- **Publication:** PLOS ONE 18(9), e0288601; published 2023-09-08.
- **Primary sources:** [Publication or author manuscript](https://doi.org/10.1371/journal.pone.0288601); [Supporting primary source](https://journals.plos.org/plosone/article?id=10.1371/journal.pone.0288601).
- **Access depth:** Selected open full-text passages; not an equation-by-equation audit.

**Project hypothesis and limitations:** This follow-up proves some conjectures from the 2021 paper; do not treat every earlier conjecture as proved. Bounds concern the unit-scale gamma median for positive shape. Supplemental checks in #14/#7 must handle scale and small-shape underflow explicitly. They do not certify arbitrary probabilities or resolve the final-scaling defect by themselves.

## R17 — Compound frequency-severity distributions

- **Authors:** Stephen Mildenhall.
- **Title:** Aggregate: fast, accurate, and flexible approximation of compound probability distributions.
- **Publication:** Annals of Actuarial Science 19(2), 193–232 (2025); online 2024-11-15.
- **Primary sources:** [Publication or author manuscript](https://doi.org/10.1017/S1748499524000216).
- **Access depth:** Publisher abstract and selected article passages; Python implementation not audited.

**Project hypothesis and limitations:** The Aggregate framework uses FFT-based methods for compound distributions. Feasibility #271 should distinguish discretization, truncation and aliasing, evaluate exponential tilting and document VBA cost. Matching the first three moments does not certify extreme tails, VaR or ES. This is a separate research direction, not an approved extension of the current dependence kernels.

## R18 — Gamma ratios and inverses

- **Authors:** Amparo Gil, Javier Segura and Nico M. Temme.
- **Title:** Efficient and Accurate Algorithms for the Computation and Inversion of the Incomplete Gamma Function Ratios.
- **Publication:** SIAM Journal on Scientific Computing 34(6), A2965–A2981 (2012); arXiv posting 2013.
- **Primary sources:** [Publication](https://doi.org/10.1137/120872553); [Author or official supporting source](https://ir.cwi.nl/pub/20433/20433B.pdf).
- **Access depth:** Selected manuscript sections on direct computation, region selection and inversion; no code audit.

**Project hypothesis and limitations:** The paper computes regularized P and Q, selecting the smaller probability first and obtaining its complement, and provides inverse initialization methods and the Fortran IncgamFI module. Compare formulas, prefactors and dispatch regions with #24/#26 before changing algorithms. Test seams and direct tails independently; a published double-precision algorithm does not establish the repository contract or monotonicity at adjacent binary64 inputs. Also relevant to gamma inverse work #14.

## R19 — Double-word arithmetic error bounds

- **Authors:** Mioara Joldes, Jean-Michel Muller and Valentina Popescu.
- **Title:** Tight and Rigorous Error Bounds for Basic Building Blocks of Double-Word Arithmetic.
- **Publication:** ACM Transactions on Mathematical Software 44(2), 2017.
- **Primary sources:** [Publication](https://doi.org/10.1145/3121432); [Author or official supporting source](https://jncf2017.lip6.fr/program/abs-popescu.pdf).
- **Access depth:** Author conference abstract and institutional publication metadata; full-paper retrieval from HAL was blocked.

**Project hypothesis and limitations:** Use this as a priority reference for auditing two-Double addition, multiplication, division and normalization in the precision work under #24/#26. Obtain the full algorithm and proof before adopting a bound. Check rounding assumptions, operand ordering, normalization, FMA requirements and exponent-range exclusions against actual VBA behavior. Primitive bounds do not automatically prove an entire Gamma evaluation or monotonicity.

## R20 — Accurate sums and dot products

- **Authors:** Takeshi Ogita, Siegfried M. Rump and Shin’ichi Oishi.
- **Title:** Accurate Sum and Dot Product.
- **Publication:** SIAM Journal on Scientific Computing 26(6), 1955–1988 (2005).
- **Primary sources:** [Publication](https://doi.org/10.1137/030601818); [Author or official supporting source](https://www.tuhh.de/ti3/paper/rump/OgRuOi05.pdf).
- **Access depth:** Publisher metadata and author-hosted abstract; no complete proof or code audit.

**Project hypothesis and limitations:** Error-free transformations and accurate accumulation offer comparisons for series and polynomial evaluation under #24/#26 and numerical audit #7. Verify the preconditions of the selected algorithm. Improving summation alone does not remove errors already introduced by log arguments, logGamma, exponentiation or a prefactor; retain the rejected sum-only experiments as evidence.

## R21 — Generalized incomplete gamma evaluation

- **Authors:** Rémy Abergel and Lionel Moisan.
- **Title:** Fast and Accurate Evaluation of a Generalized Incomplete Gamma Function.
- **Publication:** ACM Transactions on Mathematical Software 46(1), 2020.
- **Primary sources:** [Publication](https://doi.org/10.1145/3365983); [Author or official supporting source](https://helios2.mi.parisdescartes.fr/~moisan/research.php).
- **Access depth:** Author-hosted abstract and publication/software listing; direct full PDF retrieval returned a service error.

**Project hypothesis and limitations:** The generalized integral uses explicit normalization, continued fractions, integration by parts and Romberg integration for close endpoints. Assess the representation as a comparator for #24/#26, with stable conversion to regularized P/Q. Reported gains are against the paper’s comparator, not this VBA library. Interval probabilities are a possible future application, not a new approved API.

## R22 — Floating-point expansions and QD

- **Authors:** Yozo Hida, Xiaoye S. Li and David H. Bailey.
- **Title:** Algorithms for Quad-Double Precision Floating Point Arithmetic.
- **Publication:** Technical report 2000; ARITH-15 proceedings (2001), 155–162.
- **Primary sources:** [Publication](https://doi.org/10.1109/ARITH.2001.930115); [Author or official supporting source](https://escholarship.org/uc/item/69q5t2mj).
- **Access depth:** Institutional abstract and QD project documentation; implementation not audited.

**Project hypothesis and limitations:** The work describes arithmetic with unevaluated floating-point expansions, including algebraic and transcendental operations. Use it for design comparison and primitive tests in #24/#26; QD offers double-double and quad-double implementations. It does not establish that four components are required or that two components equal IEEE binary128. Audit runtime arithmetic assumptions and licensing before any code reuse.

## R23 — Bivariate and trivariate normal and t probabilities

- **Authors:** Alan Genz.
- **Title:** Numerical Computation of Rectangular Bivariate and Trivariate Normal and t Probabilities.
- **Publication:** Statistics and Computing 14(3), 251–260 (2004).
- **Primary sources:** [Publication](https://doi.org/10.1023/B:STCO.0000035304.20635.31); [Author or official supporting source](https://cran.r-project.org/web/packages/mvtnorm/refman/mvtnorm.html).
- **Access depth:** Official mvtnorm references/documentation; publisher full text not retrieved.

**Project hypothesis and limitations:** Make this a concrete baseline candidate in #196 alongside R14. Verify the actual algorithm, rectangle conventions, correlation limits and direct-tail handling before implementation. Two wrappers around the same Genz routine are not independent reference legs. Trivariate operations remain outside the current committed scope unless separately approved.

## R24 — Poisson-binomial distribution computation

- **Authors:** Yili Hong.
- **Title:** On Computing the Distribution Function for the Poisson Binomial Distribution.
- **Publication:** Computational Statistics & Data Analysis 59, 41–51 (2013); related institutional report 2011.
- **Primary sources:** [Publication](https://doi.org/10.1016/j.csda.2012.10.006); [Author or official supporting source](https://vtechworks.lib.vt.edu/items/7f2c25bf-ad02-43c5-a200-a2ac65464b46).
- **Access depth:** Publisher abstract, institutional report metadata and official poibin documentation; code not audited.

**Project hypothesis and limitations:** Compare characteristic-function/DFT evaluation against recursion and the tilted convolution approach R13 under #270. The independent heterogeneous Bernoulli assumptions remain essential. Exact finite formulas do not guarantee tiny-tail relative accuracy after floating-point evaluation; test tail probabilities and deterministic and homogeneous limits explicitly.

## R25 — Marcum Q and noncentral gamma

- **Authors:** Amparo Gil, Javier Segura and Nico M. Temme.
- **Title:** Algorithm 939: Computation of the Marcum Q-Function.
- **Publication:** ACM Transactions on Mathematical Software 40(3), article 20 (2014); preprint 2013.
- **Primary sources:** [Publication](https://doi.org/10.1145/2591004); [Author or official supporting source](https://arxiv.org/abs/1311.0681).
- **Access depth:** Author preprint abstract and ACM algorithm catalogue; full implementation not audited.

**Project hypothesis and limitations:** The algorithm combines series, integrals, asymptotics and recurrences for complementary Marcum P/Q, corresponding to noncentral gamma/chi-square probabilities. Under #197, map parameter conventions and tested domains first. Reported accuracy is about 1e-12 in the stated domain and degrades for larger parameters; it must not be presented as meeting tighter repository contracts. This is not a noncentral F algorithm by itself.

## R26 — Central beta inverse using Schwarzian–Newton

- **Authors:** Amparo Gil, Javier Segura and Nico M. Temme.
- **Title:** Efficient Algorithms for the Inversion of the Cumulative Central Beta Distribution.
- **Publication:** Author preprint and online publication 2016.
- **Primary sources:** [Publication](https://doi.org/10.1007/s11075-016-0139-2); [Author or official supporting source](https://arxiv.org/abs/1605.03503).
- **Access depth:** Author preprint abstract and publisher search metadata; proof and implementation not audited.

**Project hypothesis and limitations:** Combines fourth-order Schwarzian–Newton iteration, asymptotic inversion and sharp tail bounds. Compare convergence conditions and total work with the safeguarded inverse and R12 under #35. An improved beta solver alone does not repair cancellation when reconstructing F from the beta pair. Preserve the frozen defect protocol and distinguish local convergence order from global reliability.
