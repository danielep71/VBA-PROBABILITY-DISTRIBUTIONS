# The Loader deviance as a pair of Doubles (#498)

**Status (2026-11-26): model only.** No file under `src/` changes, and nothing
here is registered yet. The registration, with its holdout and rules, comes
before any VBA.

After #175 five densities and mass functions far from the mean are still less
accurate than Excel by more than 1E-13 (`../beta_prefactor_study/RESULT_2.md`).
Extending the existing series does not remove the cause: a Double of size D
carries D * 2^-53, and that is the relative error of the density. The model
forms the deviance as a pair of Doubles and carries the pair to the final
exponential.

In the model, at the five points of the issue the error falls from between
1.1E-13 and 4.4E-13 to below 1E-15. Without the rule the model returns the
errors Excel measured at the Poisson and Binomial points (4.4E-13, 1.9E-13,
1.1E-13); at the two Gamma points it gives 2.4E-13 and 2.5E-13 where Excel
measured 3.0E-13 and 1.9E-13, so it is not faithful there.

Scope the model supports: `Poisson_PMF`, `Binomial_PMF` and the Gamma density
(and so the chi-square density). The Binomial needs its two means, `N p` and
`N (1 - p)`, as pairs: rounded to one Double they alone cost up to 1E-13.
The Beta, F and Student t densities are not modelled here; the second holdout
of #175 has none of them above 1E-13.

| File | Role |
| --- | --- |
| `pair_model.py` | the pair arithmetic, the deviance, the three functions with and without the rule |
| `test_pair.py` | the transforms, the deviance against 60-digit values, the five points |
