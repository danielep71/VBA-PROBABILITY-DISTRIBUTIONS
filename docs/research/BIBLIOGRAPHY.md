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
