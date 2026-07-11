# Code Review — `M_STATS_PROBDIST_*` VBA Statistics Library

**Reviewer:** Claude (static inspection + cross-module symbol sweep)
**Date:** 2026-07-11
**Scope:** Six delivered modules, house-style-normalized final versions:
`CORE` (758 lines), `SPECIALFUNCS` (1451), `NORMALFAMILY` (2680), `TFAMILY` (3289), `CONTINUOUS` (4931), `TEST` (4465). ~17,500 lines total.

---

## 1. What this review is — and is not

This is a **structural and design review** based on close reading of the source plus an automated cross-module symbol-resolution sweep. It is deliberate about the limits of that method:

- It is **not** a compile. I cannot load the project into the VBE, so I cannot certify it builds clean of "Variable not defined" / type errors.
- It is **not** an independent numerical audit. I did **not** re-derive the reference constants or accuracy claims against mpmath/scipy in this pass. Statements about numerical *correctness* rest on (a) the soundness of the visible algorithms and their provenance, (b) internal consistency, and (c) the existence of your own Python-validated regression harness — not on my re-computation of every literal.
- Depth of reading varied. Read in full or near-full: `CORE` primitives, all of `SPECIALFUNCS`, `CONTINUOUS`, the `TFAMILY` numerical kernels, the `NORMALFAMILY` tail/inverse kernels, the `TEST` assertion layer. The public wrappers in `NORMALFAMILY`/`TFAMILY` were sampled rather than read line-by-line; they follow one validated pattern and the sweep confirms their wiring.

Treat the scores as "quality of the artifact as written," with the understanding that a green harness run and a clean VBE compile are the two things that would convert "excellent as inspected" into "verified."

---

## 2. Verdict

This is a **professional-grade numerical library** — well above the quality bar for typical VBA, and competitive with what you'd expect from a curated C/Fortran special-function library wrapped for a spreadsheet. The defining characteristic is an unusually disciplined **failure contract**: iterative kernels return `Boolean` and never pass off a non-converged partial sum as an answer, and arithmetic that can overflow is routed through guarded `Try*` primitives that classify the result rather than letting a `Double` silently become a fault. That single decision is what separates this from almost every hand-rolled stats library.

**Overall (as inspected): 9.3 / 10.**

The gap to 10 is not sloppiness — it's the handful of consumer-facing parameterization surprises, a small amount of residual documentation drift, and the irreducible fact that this review cannot itself run the numbers.

---

## 3. Scoring method

Each module is scored 0–10 on six dimensions, then combined with weights:

| Dimension | Weight | What it measures |
|---|---|---|
| Numerical method & correctness | 30% | Are the algorithms right, well-chosen, and cancellation-aware? |
| Robustness & error handling | 25% | Overflow/underflow classification, the Try-contract, edge cases |
| API & interface design | 10% | Signature consistency, least-surprise, worksheet parity |
| Documentation | 15% | Accuracy, honesty, provenance |
| Style & house-consistency | 10% | Post-normalization formatting discipline |
| Test coverage | 10% | Reflected via the `TEST` suite exercising that module |

Scores in the **9+** band mean "excellent, ship it"; they are not rounded up from lower work.

---

## 4. Module-by-module

### 4.1 `CORE` — infrastructure & arithmetic primitives

The foundation. Its most important idea is the **split between finiteness and supported magnitude**: `PROB_IsFinite` (is this a real, non-fault `Double`?) is a distinct question from `PROB_IsWithinSupportedMagnitude` (is it inside the range the kernels are allowed to trust?), and the two are not conflated. `PROB_TryExp` keys off the true floating-point boundary rather than a hard-coded `MAX_EXP`, and the `TryAdd/TryMultiply/TryDivide` trio give every downstream module a way to do arithmetic that *reports* overflow instead of producing infinity. The re-architecture that lets huge `X` saturate CDFs to exactly 0/1 is mathematically correct, not a hack.

**Notable:** the underflow-is-a-valid-zero handling in `PROB_TryExp` (negative argument that underflows returns `True` with result 0) is exactly the contract the density/tail code needs, and it's threaded consistently.

**Watch-item:** `CORE` and `SPECIALFUNCS` are `Option Private Module`. That's correct — they're kernels, not worksheet surface — but it means their `Public` members are project-visible only. The sweep confirms nothing outside expects them to be worksheet-callable.

| Dim | Score |
|---|---|
| Method/correctness | 9.3 |
| Robustness | 9.7 |
| API | 9.2 |
| Documentation | 9.3 |
| Style | 9.8 |
| Testing | 8.8 |
| **Composite** | **9.4** |

### 4.2 `SPECIALFUNCS` — log-gamma, log-beta, incomplete beta/gamma & inverses

The numerical heart, and the strongest module. It reduces the entire non-normal distribution stack to the regularized incomplete beta `I_x(a,b)` and incomplete gamma `P/Q(a,x)`, and does the hard parts once, well:

- **Never-return-a-partial-sum.** Every Lentz continued fraction and every series returns `False` on non-convergence instead of its current accumulator. This is the single most valuable safety property in the library, and it is applied without exception.
- **Two-argument incomplete beta.** `PROB_TryBetaRegularized(X, Y, …)` takes both `X` and `Y = 1−X` from forms that don't cancel and never re-derives one by subtraction — the reason Student-t is exact near zero and F-quantiles reach 1E+34.
- **Loader's Stirling error** for `PROB_LogChoose`, assembling `log C(N,K)` from three small `δ` corrections instead of subtracting two `N·log N`-sized log-gammas.
- **`PROB_LogGammaHalfDiff`** and the **unbalanced-argument asymptotic** in `PROB_LogBeta` both target specific, documented cancellation failures with correct magnitude thresholds.
- Inverses (`TryBetaInvRegularized`, `TryGammaInvP`) solve on the **smaller tail** and use safeguarded Newton bracketed by bisection that cannot diverge.

The one real defect found this session — an unguarded `Exp(2·W)` in the AS 109 seed — is now fixed (routed through `PROB_TryExp`, with the seed clamp as a second net). Provenance is documented per function with honest, specific error figures.

| Dim | Score |
|---|---|
| Method/correctness | 9.5 |
| Robustness | 9.7 |
| API | 9.3 |
| Documentation | 9.6 |
| Style | 9.7 |
| Testing | 8.8 |
| **Composite** | **9.5** |

### 4.3 `NORMALFAMILY` — standard normal, general normal, lognormal

Direct-tail throughout: `PROB_NormalUpperTailPositive` computes `Q(Z)` with a Hart rational approximation below the split and a 16-term Laplace continued fraction in the far tail, with an underflow cutoff at the point (`Z ≈ 38.49`) where the one-sided probability genuinely rounds to zero — so survival probabilities are never recovered as `1 − CDF`. The interval-probability kernel branches on which tail the bounds fall in, protecting same-tail intervals from collapsing (the N5 regression concern). The inverse (`PROB_NormalInvCDF`) is an Acklam seed plus a **single Halley step**, and — crucially — the refinement is **guarded against saturated regions** (`If PdfX > 0 And CdfX > 0 And CdfX < 1`), so it doesn't corrupt a tail quantile where the residual is no longer informative.

**Highlight — documentation honesty.** The `ACCURACY` block states plainly that tail quantiles carry ~1E-10 *relative* error because the Hart/West arrangement is 1E-15 *absolute*, not relative. Very few libraries are this candid about where their precision actually is.

**Minor:** the module retains a few intentionally-terse legacy kernels (`PROB_NormalPDF/CDF/Survival`), and a couple of comment lines run slightly over width after the `M_STATS_CORE → M_STATS_PROBDIST_CORE` rename expanded them. Cosmetic only.

| Dim | Score |
|---|---|
| Method/correctness | 9.3 |
| Robustness | 9.2 |
| API | 9.2 |
| Documentation | 9.5 |
| Style | 9.3 |
| Testing | 8.8 |
| **Composite** | **9.3** |

### 4.4 `TFAMILY` — Student t, chi-square, F

The most algorithmically intricate module, and it holds up. The Student-t tail forms are built so **`X` is never squared directly**: `DF=1` uses arctangent and its reciprocal-tail form, `DF=2` uses algebraic forms that avoid an overflowing square, and the general case forms the complementary beta arguments `DF/(DF+X²)` and `X²/(DF+X²)` through a **logistic/softplus pair on `2·Log|X| − Log(DF)`** — no squaring, no subtraction. That's what lets extreme F ratios assemble without intermediate overflow. The t-inverse uses a third-order Cornish-Fisher seed into a safeguarded Newton loop whose **every** arithmetic step (division, addition, bracket expansion) is guarded and falls back to bisection, with a stable `Low + 0.5·(High−Low)` midpoint and no artificial quantile ceiling.

Chi-square and F are thin, correct wrappers onto the incomplete-gamma/beta kernels, reusing the cross-family machinery.

**Minor:** Student-t survival accepts negative `X` (unlike Excel `T.DIST.RT`) — a deliberate, documented divergence, but a divergence a consumer should know about.

| Dim | Score |
|---|---|
| Method/correctness | 9.4 |
| Robustness | 9.5 |
| API | 9.1 |
| Documentation | 9.4 |
| Style | 9.6 |
| Testing | 8.9 |
| **Composite** | **9.4** |

### 4.5 `CONTINUOUS` — Gamma, Beta, Exponential, Weibull, Uniform

Broad and consistent. Densities are assembled in the log domain with the scale terms arranged to cancel cleanly; left tails use `Expm1`; the Weibull variance factor is formed as `2·logΓ(1+ε) + log(e^Δ−1)` with a small-Δ `Expm1` branch and a large-shape ζ(2)-leading asymptotic to avoid the catastrophic `Γ(1+2/k) − Γ(1+1/k)²` cancellation; Uniform uses scaled-width and convex-combination forms so opposite-sign bounds whose true width exceeds `Double` max are still handled. The `Beta` two-shape validator correctly constrains *both* shapes (the earlier asymmetry is gone), and density poles return `#NUM!` as specified.

**Biggest API caveat in the library lives here.** Two parameterization choices will surprise a consumer who assumes internal uniformity:
1. **Exponential is rate-parameterized** (matching `EXPON.DIST`), while its sibling continuous distributions (Gamma, Weibull) are scale-parameterized. This is now a settled, documented choice — but it *is* an inconsistency across the module.
2. **Shape parameters are capped at "supported magnitude"** while scale/rate parameters get the full finite `Double` range. Intentional (shapes drive the iterative kernels; scales don't), and documented — but asymmetric.

Neither is a bug; both are legitimate design calls. They're the main thing I'd want a downstream user to read the header about.

| Dim | Score |
|---|---|
| Method/correctness | 9.3 |
| Robustness | 9.5 |
| API | 8.9 |
| Documentation | 9.3 |
| Style | 9.6 |
| Testing | 8.9 |
| **Composite** | **9.3** |

### 4.6 `TEST` — consolidated regression harness

A genuine test suite, not a smoke test: five argument-less entry points, silent passes, one-line detailed failures, a consolidated verdict, and a documented **regression registry** (C1–C3, N1–N6, T1–T7, D1–D6) that ties each named risk to concrete assertions. The assertion layer is itself defensively written — `AssertClose`/`AssertRelClose` reject `CVErr` and non-numeric returns *before* comparing, form their difference/ratio through the guarded `Try*` primitives so an extreme mismatch can't overflow the harness, and `AssertRelClose` falls back to absolute comparison at a zero reference. Error-code assertions distinguish `#NUM!` from `#VALUE!`, enforcing the failure-classification contract rather than just "did it error."

**The honest ceiling on my confidence sits here.** I cannot run this harness, and its reference constants are validated by *your* external mpmath/scipy pipeline, not by me. The suite also, by design, tests reachable numerical cases rather than forcing artificial non-convergence — a reasonable choice, but it means the "never return a partial sum" property is verified by construction and inspection more than by a triggered timeout. Hence a deliberately lower "independence/verifiability" sub-score.

| Dim | Score |
|---|---|
| Coverage & design | 9.0 |
| Harness robustness | 9.4 |
| Usability | 9.2 |
| Documentation | 9.2 |
| Style | 9.2 |
| Independence / verifiability (by me) | 7.5 |
| **Composite** | **8.9** |

---

## 5. Cross-cutting strengths

- **The Try-contract is applied without exception.** Non-convergence and overflow are first-class, reported states everywhere — the property most numerical libraries get wrong.
- **Cancellation-awareness is systemic**, not spot-applied: two-argument beta, `LogGammaHalfDiff`, `Expm1`/`Log1p` in tails, never-square-X, scaled-width Uniform, log-domain moments.
- **Direct-tail evaluation** (survival via `Q`, swapped beta) preserves probabilities that `1 − CDF` would round away.
- **Clean cross-module wiring**, confirmed by the sweep: 234 definitions, 139 cross-module reference resolutions, 354 dependency-header entries — zero undefined references, zero private-scope leaks, zero name collisions, correct `Option Private Module` scoping.
- **Documentation is honest about its own limits** — the tail-accuracy caveats and per-function provenance are a cut above.
- **Uniform house style** post-normalization: consistent banners, rulers, `Dim` alignment, `Err_Handler`/`Fail_Num` scaffolding, and error policy (`#NUM!` vs `#VALUE!`).

## 6. Weaknesses, risks & caveats (honest)

Ordered roughly by how much they'd matter to a consumer:

1. **Parameterization asymmetries in `CONTINUOUS`** (rate-vs-scale Exponential; shape magnitude-cap vs full-range scale). Documented, intentional, but the most likely source of downstream confusion.
2. **Accuracy claims are self-reported.** The header grids (e.g., "Student t quantile ≤ 3.0E-12 relative") are the contract, but this review did not independently reproduce them. They live or die by your regression run.
3. **Verification gap by construction.** No compile and no harness execution happened here; the sweep is static only.
4. **Non-short-circuit `And`.** A known VBA hazard in your notes. In the code I read (`CONTINUOUS`, `TFAMILY` — none; the `And` guards in `SPECIALFUNCS`/`NORMALFAMILY` I saw are division-free and safe), I found no live instance — but I did not exhaustively audit every `And` across all ~17,500 lines. Worth a one-line grep-and-eyeball as a standing check.
5. **Very large iteration budgets** (1E5). Correct and documented, but a genuinely pathological input would spin a long time before correctly reporting failure.
6. **Minor documentation drift.** A few over-width comment lines in `NORMALFAMILY` from the module-name expansion; some deliberately-verbose VALIDATE/COMPUTE prose in `CONTINUOUS` left un-flattened; undocumented (but now aligned) throwaway locals in `TEST`.

## 7. Recommendations (prioritized)

1. **Keep the green harness run as the gate.** Nothing in this review substitutes for it; it's what closes items 2–3 above.
2. **Add a standing "non-short-circuit `And`" lint** — a trivial grep in your scraper that flags any `… And …` whose right operand contains `/` or a call that could fault. Cheap insurance against the one hazard class you've already been bitten by.
3. **Consider a one-paragraph "parameterization at a glance" table** in a top-level README or the `CONTINUOUS` header, spelling out rate-vs-scale and the shape-cap rule in one place, so a consumer doesn't have to infer it per function.
4. **Optional polish:** re-wrap the handful of over-width comment lines in `NORMALFAMILY`; add short descriptions to the bare aligned locals in `TEST` if you want the fully-documented look everywhere.
5. **Optional:** if you ever want the "never return a partial sum" property covered by an *executed* test rather than by inspection, a single fixture with a deliberately tiny iteration cap on one kernel would triggered-test the `False` path.

## 8. Score summary

| Module | Composite |
|---|---|
| `CORE` | 9.4 |
| `SPECIALFUNCS` | 9.5 |
| `NORMALFAMILY` | 9.3 |
| `TFAMILY` | 9.4 |
| `CONTINUOUS` | 9.3 |
| `TEST` | 8.9 |
| **Library (as inspected)** | **9.3** |

**Bottom line:** this is code a numerical-methods reviewer would be happy to sign off on, pending a clean compile and a green regression run. The engineering discipline — especially the uncompromising failure contract and the systematic cancellation-avoidance — is the real asset, and it's consistent across all six modules.
