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
| Research documentation | [#265](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/265) | Bibliography, methodology, template and discoverable navigation |

## Scope safeguards

No new numerical implementation issue duplicates the existing Beta, inverse,
log-space or noncentral work. Literature links do not alter v1.0.0 release
criteria, approve contract relaxation or imply a new required dependency.
The new fitting issue is a bounded feasibility decision, not a commitment to
ship parameter estimation in a maintenance release.

No COS engine, public noncentral-beta API or inverse-log-probability API is
silently added by this map. Record the explicit inventory decision in the owning
issue before implementation; split work that exceeds its existing scope.
