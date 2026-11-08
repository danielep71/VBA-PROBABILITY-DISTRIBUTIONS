# Methodology for evaluating numerical research

[Research index](README.md) · [Bibliography](BIBLIOGRAPHY.md)

This protocol complements existing repository requirements. It does not amend a
frozen study, waive a release gate or retroactively preregister observed results.
The research review itself has produced no numerical measurements.

## 1. Define the decision before implementation

Record the owning issue, baseline SHA, public operations, supported domains,
existing contracts and affected callers. State a falsifiable hypothesis: improved
worst error, convergence, supported range or runtime, with a specified comparator.
Distinguish a confirmed defect from an optional improvement or new capability.
Keep the option to reject a newer algorithm if its benefit is not demonstrated.

Inspect the complete paper and required supplementary algorithms before porting.
Record first-publication date, revision, errata, access depth and code license.
Map parameterization and tail conventions explicitly; identical function names
do not imply identical inputs or outputs.

## 2. Diagnose before replacing a kernel

Separate prefactor error, recurrence/series error, complement cancellation,
dispatch error and inverse stopping/rounding error. For Beta, an alternative
continued fraction need not fix a defective prefactor. Measure components with
faithful operation order and explicit binary64 rounding where load-bearing.

Use model results to design experiments, not to claim Excel certification. Mark
idealized surrogates and explain which actual VBA operations they omit.

## 3. Establish independent references

- Evaluate references at the exact binary64 values supplied to VBA, recording
  hexadecimal/bit identities where useful, not just the original decimal tokens.
- Form complementary reference arguments at sufficient precision from those
  exact inputs; do not silently round an exact complement or evaluate at a
  different decimal input. Follow each public wrapper's argument contract.
- Compute small survival probabilities directly, not by subtracting a rounded
  CDF from one. Use log references for underflowing probabilities.
- Record reference algorithm, precision, software version and convergence status.
  Raise precision and use structurally independent checks where feasible.
- Two wrappers of one implementation are one evidence leg. Agreement at two
  precisions does not prove convergence after an iteration cap.
- Keep unresolved oracle gaps visible. Analytic identities and monotonicity are
  useful checks but are not a substitute for independent accuracy evidence.

## 4. Preregister domains, metrics and holdouts

Before viewing candidate results, freeze the comparison protocol, expected caller
movement, fitting grid, acceptance thresholds and an untouched holdout strategy.
Include ordinary inputs and adversarial seams: balanced/unbalanced shapes,
parameters below one, extreme degrees of freedom, endpoints, subnormals,
overflow/underflow boundaries, dispatch switches and non-convergence paths.

Preserve every existing preregistration. A failed experiment stays failed; a new
candidate needs a new prospective comparison. Do not tune on a holdout and then
continue calling it untouched. Follow existing one-shot holdout safeguards.

## 5. Measure the returned quantity

| Quantity | Required diagnostic |
| --- | --- |
| Density or ordinary probability | Absolute/relative/ULP errors as appropriate to its contract; worst input and regime |
| Small tail | Relative tail error and direct-tail reference; underflow classification |
| Log probability | Error on log scale, including where the ordinary probability underflows |
| Continuous quantile | Error in the returned quantile against an independent reference, plus probability residual |
| Discrete quantile | Generalized-inverse inequality F(k-1) < p <= F(k), with endpoint and support rules |
| Iterative computation | Iteration count, convergence criterion, exhaustion and invalid-state behavior |

For continuous quantiles, the first-order diagnostic

\[
\Delta x \approx \frac{\Delta F}{f(x)}
\]

explains why a small CDF residual can correspond to a larger quantile error when
the density is small. It is a local sensitivity approximation, not a universal
rigorous bound or a replacement for measured output error. Near zero, use an
explicit absolute/scaled metric rather than undefined relative error. Apply
representability-aware rules at endpoints. R03's COS convergence assumptions do
not transfer automatically to all distributions.

Round trips, complement identities, symmetry and monotonicity supplement these
measurements. Correlated errors in forward and inverse paths can make a round
trip look good while both outputs are inaccurate.

## 6. Validate shared callers and the actual Excel runtime

Enumerate reachable callers from source, including transformed inputs and both
tails. For Beta changes, inspect Beta, F, Student-t and relevant discrete callers;
for Gamma changes, inspect Gamma, Chi-square and Poisson paths. Do not assume all
operations take the same route. Compare predicted with observed movement.

Run applicable deterministic regressions, main-grid contracts, independent
holdout and failure controls. Compile and execute changed VBA on the exact source
candidate. Record Excel version/build/bitness, source hashes, exporter identity,
grid digests, environment and cleanup. Do not rebind old measurements to new code.
See [provenance](../../benchmark/PROVENANCE.md) and
[Excel certification](../EXCEL_CERTIFICATION.md).

## 7. Measure cost fairly

Compare equivalent accuracy and supported domains on the same host. Separate
reference-generation time, VBA computation and worksheet-call overhead. Report
warm-up policy, repetitions, median/dispersion, worst iteration counts and
failures. A fast approximation with weaker guarantees is not an equivalent
replacement. No performance claim follows from publication date alone.

## 8. Record an adoption decision

Publish a reproducible comparison under `benchmark/<study>/`, linked from the
owning issue and this research index. Include source SHA, scripts, frozen inputs,
reference provenance, metrics, full failures, limitations and Excel evidence.

- **Adopt:** improvement is demonstrated across the claimed domain, with no
  unexplained supported-domain regressions and all applicable gates satisfied.
- **Defer:** plausible value, but missing oracle, access, scope agreement, runtime
  evidence or acceptable cost. Name the missing condition and owner.
- **Reject:** evidence or maintenance cost does not justify adoption; retain the
  result to avoid repeating the same investigation.

Threshold changes need dedicated evidence and review; never weaken a threshold
just to accept a candidate. New public APIs require the additive-baseline and
evidence procedures tracked by #192/#193. New release-blocking checks require
negative proofs under [verification depth](../VERIFICATION_DEPTH.md).

## Family-specific decisions

- **Log-input normal inverse:** specify log-p/log-q, sign, endpoints, invalid
  values, supported range and failure behavior. Returning LogCDF/LogSurvival is
  a different API. Do not round-trip through an underflowing exponential.
- **Noncentral families:** distinguish quantile and noncentrality inversion,
  prove or bound series truncation, test the central limit and direct tails,
  and define explicit iteration/resource limits.
- **Poisson/negative-binomial fitting:** separate numerical likelihood accuracy
  from optimizer convergence and statistical estimation. Include boundary
  optima, all-zero/underdispersed samples and real-shape parameter compatibility.
  A fitting feasibility study does not change an evaluation-only API.
