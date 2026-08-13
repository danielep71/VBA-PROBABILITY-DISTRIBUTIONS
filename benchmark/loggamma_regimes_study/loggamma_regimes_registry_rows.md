# Registry rows — LogGamma regimes (#15) and LogGamma1p kernel (#12)

> **SUPERSEDED — do not action.** The authoritative registry is
> `benchmark/accuracy_contracts.csv`. The final contract freeze landed in
> `1fac483` and the domain correction in `189ac25`. The rows and
> instructions below are retained only as pre-freeze design history and no
> longer describe the shipped contracts.

Holdout exported and passed on real VBA (62/62 points, 0 shared with the fitting set, all three contracts PASS). Ready to paste. The `provenance` and `status` values below
match the 162 existing rows exactly; no new vocabulary is introduced, because the
holdout makes `validated and frozen` an honest claim rather than an aspiration.

---

## `benchmark/accuracy_contracts.csv`

Header: `contract_id,function,regime,measure,metric,threshold,domain,provenance,status,evidence,notes`

```csv
LogGamma.small_positive.log_abs,PROB_LogGamma,small_positive,output_error,absolute,5E-13,"0 < Z <= 0.5, covering the LogGamma1p series branch (Z <= PROB_LG1P_SERIES_MAX = 0.25) and the reflection branch above it",validated and frozen,active,loggamma_regimes_study,"Absolute error in the logarithm, not relative: downstream callers propagate Exp(v), whose relative error is approximately the absolute error of v. Fitting worst 8.57E-14 at Z = 8.095E-320; holdout worst 6.27E-14 at Z = 7.0103E-320 over 36 disjoint points, margin 8.0x. The floor is EPS * Abs(Log(Z)) and is irreducible: 1.65E-13 at the smallest positive subnormal, 5.1E-14 at 1E-100, 3.1E-16 at Z = 0.25. Replaces the withdrawn global claim of relative error below 6.1E-14 across Z in [1E-8, 1E+50]."
LogGamma.near_zero.log_abs,PROB_LogGamma,near_zero,output_error,absolute,5E-14,"neighbourhoods of the zeros of Log(Gamma) at Z = 1 and Z = 2, taken as 0.9 <= Z <= 1.1 and 1.9 <= Z <= 2.1, including the two exact zeros",validated and frozen,active,loggamma_regimes_study,"Absolute only. Log(Gamma(Z)) is exactly zero at Z = 1 and Z = 2, so a relative contract is undefined there and ill-conditioned nearby: 9.31E-14 relative at Z = 1.75 is about 7.9E-15 absolute on a value of magnitude 0.084. Fitting worst 9.77E-15 at Z = 2.0; holdout worst 1.13E-14 at Z = 2.0872 over 16 disjoint points, margin 4.4x. The tighter 1-2-5 step of 2E-14 would have held on the holdout at only 1.77x, below its own 2.05x fitting margin, so 5E-14 is validated by the holdout rather than merely preferred."
LogGamma.general.output_rel,PROB_LogGamma,general,output_error,relative,5E-13,"Z >= 0.5 excluding the neighbourhoods of Z = 1 and Z = 2, where Abs(LogGamma(Z)) is bounded away from zero",validated and frozen,active,loggamma_regimes_study,"The one regime in which a relative contract is meaningful. Fitting worst 9.31E-14 at Z = 1E+50; holdout worst 3.24E-14 at Z = 2.7183 over 10 disjoint points, margin 15.4x."
LogGamma1p.small.scaled_abs,PROB_TryLogGamma1p,small,scaled_output_error,absolute,5E-16,"3.854839696505424E-308 <= X <= PROB_LG1P_SERIES_MAX = 0.25; the lower bound is PROB_MIN_NORMAL / EulerGamma",validated and frozen,active,loggamma1p_study,"Metric is Abs(observed - reference) / X, not ordinary absolute or relative error: the scaled Gamma inverse computes [LogProbability + LogGamma1p(Shape)] / Shape, so the scaled error IS the relative error of the quantile it produces. Worst measured 1.36E-16 at X = 1E-13, margin 3.7x. The floor is the rounding of the final Acc * X, about two ulps; it is not coefficient-limited, so widening further coefficients would buy nothing."
LogGamma1p.series_seam.scaled_abs,PROB_TryLogGamma1p,series_seam,scaled_output_error,absolute,1E-13,"one-ulp neighbourhood of PROB_LG1P_SERIES_MAX = 0.25",validated and frozen,active,loggamma1p_study,"Kept out of the small-series contract because the point one ulp ABOVE the seam runs the delegated PROB_LogGamma route and inherits that kernel's error, roughly two orders of magnitude above the series floor. Measured 2.24E-14 at 0.25 + 1 ulp against 4.31E-17 at 0.25; step across the seam 5.68E-14. Pooling the two sides would either loosen the series contract 100x or fail the seam for no reason."
```

---

## `benchmark/numerical_limitations.csv`

Header: `limitation_id,affected_functions,domain,cause,observed_effect,evidence,status,notes`

```csv
LogGamma1p.SubnormalResultRepresentability,PROB_TryLogGamma1p,"X below 3.854839696505424E-308, that is PROB_MIN_NORMAL / EulerGamma","the leading term EulerGamma * X is itself subnormal, so the returned Double has no grid point near the answer","scaled absolute error grows toward the half-ulp bound 0.5 * minimum_subnormal / X: measured 1.92E-16 at X = 1E-308, 1.29E-14 at 1E-310, 1.41E-04 at 1E-320 and 4.18E-01 at the smallest positive subnormal",loggamma1p_study,characterized,"Representability limit of the OUTPUT, not an algorithm defect: no evaluation order can place a value on a grid that has no point near it. The boundary is mathematical rather than empirical - measured error stays inside the small-series contract somewhat below it, but the quantization begins exactly where EulerGamma * X leaves the normal range. Measured is at or below the closed-form bound at every point, worst gap 11.7x at X = 1E-315; the bound is a half-ulp ceiling, not a prediction."
```

---

## Order of operations

1. Regenerate the holdout: `python3 generate_loggamma_regimes_holdout.py`
2. With **Phase 1** loaded, `Export_LogGammaRegimes`, pick the holdout
3. `python3 analyze_loggamma_regimes.py --baseline loggamma_regimes_baseline.csv --holdout loggamma_regimes_holdout.csv`
4. Only if section 6 reports **PASS on all three**, paste the rows above
5. `cd benchmark && python3 refresh_evidence.py`
6. `benchmark/README.md` is generated from `numerical_limitations.csv`, so it will
   change; that is expected and `refresh_evidence.py` handles it

If any threshold fails on the real export, adjust **that single threshold** to the
honest holdout-inclusive worst and record why. Do not adjust the others to match.

## What this closes

- **#15** — the three `LogGamma.*` rows are the last item; subfinding A is fixed
  and the header claim already withdrawn.
- **#12** — the two `LogGamma1p.*` rows plus the limitation entry. Its remaining
  checklist items (`PROB_StirlingError`, forward branches, inverse) belong to
  Phase 2, #13 and #14 respectively and should be moved to a comment rather than
  held open here.
