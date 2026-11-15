# Fix comparison implementation preflight — 2026-11-12

Baseline: `a3f018bf002702c020910c6491ef5ae1d1f3d1cd`. Status: implementation prerequisites incomplete; no candidate evaluated.

The frozen FIX_COMPARISON_PREREGISTRATION.md remains unchanged. The original
NONE isolation result is not confirmation. No independent holdout observations
were inspected and no numerical source, grid, threshold or manifest changed.

## Reuse assessment

The existing `beta_balanced_survival/beta_route_model.py` is not a complete
forward model for this study. Its `log_beta` deliberately raises `NotModelled`
for the production `PROB_LogGammaDelta` branch. This is a safe refusal, not an
incorrect result. A direct local probe returned that exception for all three
retained shape pairs (50, 5e9), (5e7, 5), and (5e5, 1.5). These correspond to
F degrees of freedom (100, 1e10), (1e8, 10), and (1e6, 3).

Production `PROB_LogBeta` uses the stable delta when Large >= 1 and
Small/Large < 0.1. Substituting a rounded high-precision log-beta or the
three-log-gamma identity would not satisfy the frozen operation-order model.
The old `f_inverse_model.py` also uses a rounded high-precision forward CDF;
it cannot serve as the new baseline unchanged.

## Arithmetic semantics to distinguish

The frozen protocol specifies Python float binary64 operations. The balanced
Beta model instead documents extended intermediates and Double storage
boundaries, supported by its own historical observations. Those observations
are not proof of the inverse path's arithmetic semantics. Keep the prescribed
binary64 baseline and evaluate its frozen faithfulness gate before selecting
any candidate. If it fails, record the failure and stop as required; do not
silently substitute another arithmetic model or revise the frozen criteria.

## Remaining implementation sequence

1. Port current LogGammaDelta and its dependencies, including guards and the
   half-integer shortcuts. Bind all reached source procedures and constants.
2. Port the forward prefactor/CF and raw normal seed with explicit arithmetic
   order; do not use an oracle as a production component.
3. Port inverse bracket/Newton termination, supplied complement, pair unwind,
   both-shapes-below-one representability check and public F reconstruction.
4. Record exact inputs, seed, iterations, residual, density, bracket fallback,
   settled root and pair. Check the retained-row faithfulness criteria first.
5. Only after that gate, run F-A/F-B on the prescribed fitting rows. The
   holdout and shared-path invariance sets remain gated by the frozen rules.

This preflight is not a completed model, a candidate comparison, an Excel
validation or a resolution of issue #35. No acceptance checkbox is satisfied
by this document alone.
