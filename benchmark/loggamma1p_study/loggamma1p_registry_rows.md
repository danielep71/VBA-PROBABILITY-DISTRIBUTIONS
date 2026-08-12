# Draft registry rows — PROB_TryLogGamma1p (issue #12)

Schemas taken from the committed headers. **Do not paste until the full Phase 0
sweep is clean** — thresholds are measured but not holdout-validated, so
`provenance` says so.

---

## `benchmark/accuracy_contracts.csv`

Header: `contract_id,function,regime,measure,metric,threshold,domain,provenance,status,evidence,notes`

```csv
LogGamma1p.small.scaled_abs,PROB_TryLogGamma1p,small,scaled_absolute_error,absolute,5E-16,"X from 3.854839696505424E-308 (PROB_MIN_NORMAL / EulerGamma) through PROB_LG1P_SERIES_MAX = 0.25",measured; pending independent holdout,active,loggamma1p_study,"Metric is Abs(observed - reference) / X, not ordinary absolute or relative error: the scaled Gamma inverse computes [LogProbability + LogGamma1p(Shape)] / Shape, so the scaled error IS the relative error of the quantile it produces. Worst measured 1.36E-16 at X = 1E-13, headroom 3.7x. The floor is the rounding of the final Acc * X, about two ulps; it is not coefficient-limited, so widening further coefficients would buy nothing."
LogGamma1p.series_seam.scaled_abs,PROB_TryLogGamma1p,series_seam,scaled_absolute_error,absolute,1E-13,"one-ulp neighbourhood of PROB_LG1P_SERIES_MAX = 0.25",measured; characterization,active,loggamma1p_study,"Kept out of the small-series contract because the point one ulp ABOVE the seam runs the delegated PROB_LogGamma route and inherits that kernel's error, roughly two orders of magnitude above the series floor. Measured 2.24E-14 at 0.25+1ulp against 4.31E-17 at 0.25; step across the seam 5.68E-14, inside PROB_LogGamma's own published bound. Pooling the two sides would either loosen the series contract 100x or fail the seam for no reason."
```

---

## `benchmark/numerical_limitations.csv`

Header: `limitation_id,affected_functions,domain,cause,observed_effect,evidence,status,notes`

```csv
LogGamma1p.SubnormalResultRepresentability,PROB_TryLogGamma1p,"X below 3.854839696505424E-308, that is PROB_MIN_NORMAL / EulerGamma","the leading term EulerGamma * X is itself subnormal, so the returned Double has no grid point near the answer","scaled absolute error grows toward the half-ulp bound 0.5 * minimum_subnormal / X: measured 1.92E-16 at X=1E-308, 1.29E-14 at 1E-310, 1.41E-04 at 1E-320 and 4.18E-01 at the smallest positive subnormal",loggamma1p_study,characterized,"Representability limit of the OUTPUT, not an algorithm defect: no evaluation order can place a value on a grid that has no point near it. The boundary is mathematical rather than empirical - measured error stays inside the small-series contract somewhat below it, but the quantization begins exactly where EulerGamma * X leaves the normal range. Measured is at or below the closed-form bound at every point, worst gap 11.7x at X=1E-315; the bound is a half-ulp ceiling, not a prediction."
```

---

## Not yet a row: `PROB_LogGamma` relative claim

Confirmed in real VBA during this study, but it belongs to `PROB_LogGamma`, not
to this kernel, so it is deliberately excluded from the rows above.

The header publishes *"Relative error below 6.1E-14 across Z in [1E-8, 1E+50]"*.
Measured 9.31E-14 at Z = 1.75. `Log(Gamma)` has zeros at Z = 1 and Z = 2, so
relative error there is unbounded by construction while absolute error stays at
or below about 1.5E-14.

Either restate the claim as absolute, or domain-exclude a neighbourhood of the
zeros. Belongs on issue #7 (contract margin audit) or its own hygiene issue —
and it should be settled **before** Phase 1, because Phase 1 changes
`PROB_LogGamma` and a wrong published claim would then look like a regression
caused by that edit.
