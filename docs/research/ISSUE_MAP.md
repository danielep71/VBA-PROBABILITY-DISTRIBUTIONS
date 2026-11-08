# Research-to-issue map

[Research index](README.md) · [Bibliography](BIBLIOGRAPHY.md)

Planning snapshot: 2026-11-08. Live issues own status, dependencies, dates and
acceptance criteria. This map is not a duplicate completion or release tracker.

| Research | Owning work | Intended use |
| --- | --- | --- |
| R01, R02: incomplete beta | [#175](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/175) | Compare candidate methods after diagnosing prefactor versus iteration error; v1.0.2 numerical maintenance |
| Existing balanced-beta defects | [#163](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/163), [#173](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/173) | Preserve current v1.0.0 evidence/acceptance scope; link optional research through #175 |
| R03: inverse conditioning | [#7](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/7) | Review metric adequacy during the post-release contract-margin audit; avoid duplicating existing inverse contracts |
| Current inverse defects | [#213](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/213), [#35](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/35) | Supplement existing diagnosis with output-error and residual separation; do not overwrite frozen protocols |
| R04: Gaussian log probabilities | [#194](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/194) | Inform log-space design; separately assess log-input inverse scope under v1.1.0 Symmetry |
| R05, R06: noncentral beta and inversion | [#197](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/197) | Support noncentral F design under v1.4.0; decide explicitly whether a public Beta API or noncentrality estimator belongs |
| R07: Poisson-limit fitting | [#266](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/266) | New feasibility/design issue in v1.0.2; implementation requires separate future scope |
| R08: discrete inverse seeds | [#269](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/269) | Bounded feasibility; exact discrete correction and total cost |
| R09: Student-t asymptotics | [#213](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/213) | Optional seam/large-df comparison; preserve defect acceptance |
| R10: incomplete-gamma bounds | [#24](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/24), [#26](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/26), [#7](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/7) | Supplementary bounds with domain and normalization checks |
| R11: arbitrary-precision gamma | [#7](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/7), [#175](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/175) | External reference feasibility and algorithm independence |
| R12: third-order beta inverse | [#35](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/35) | Optional solver comparison; reconstruction remains separate |
| R13: Poisson-binomial tails | [#270](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/270) | Independent Bernoulli feasibility, direct recursion versus tilted FFT |
| R14: bivariate normal series | [#196](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/196) | Alternative candidate; revised source and tail-specific validation |
| R15, R16: gamma median bounds | [#14](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/14), [#7](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/7) | Median-only supplementary checks, including small-shape log scaling |
| R17: compound frequency-severity | [#271](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/271) | Separate feasibility; discretization, aliasing and tail validation |
| R18: Gamma P/Q and inverses | [#24](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/24), [#26](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/26), [#14](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/14) | Compare formulas, prefactors, dispatch and inverse seeds |
| R19/R20/R22: floating-point expansions and accumulation | [#24](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/24), [#26](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/26), [#7](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/7) | Audit two-Double primitive assumptions and total error; informs PRs #241/#242 |
| R21: generalized incomplete gamma | [#24](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/24), [#26](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/26) | Alternative representation and normalization comparison |
| R23: bivariate normal and t | [#196](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/196) | Genz baseline alongside R14; no trivariate scope expansion |
| R24: Poisson-binomial DFT | [#270](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/270) | Compare with recursion and R13; verify tiny-tail errors |
| R25: Marcum Q | [#197](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/197) | Noncentral gamma/chi-square parameter and accuracy assessment |
| R26: central beta inverse | [#35](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/35) | Compare convergence/cost; F reconstruction remains separate |
| Research documentation | [#265](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/265) | Bibliography, methodology, template and discoverable navigation |

## Scope safeguards

No new numerical implementation issue duplicates the existing Beta, inverse,
log-space or noncentral work. Literature links do not alter v1.0.0 release
criteria, approve contract relaxation or imply a new required dependency.
The fitting and discrete-seed issues are bounded feasibility decisions, not
commitments to ship new production functionality in a maintenance release.
Poisson-binomial and compound-distribution studies are exploratory work planned
alongside v1.3.0; they do not expand #196 or become its prerequisites.

No COS engine, public noncentral-beta API or inverse-log-probability API is
silently added by this map. Record the explicit inventory decision in the owning
issue before implementation; split work that exceeds its existing scope.
