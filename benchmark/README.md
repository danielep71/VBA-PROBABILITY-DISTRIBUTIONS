# Accuracy benchmarks

Reproducible evidence for the accuracy claims published in the VBA source
comments — so those claims are artifacts a contributor can regenerate, not
statements that have to be taken on trust.

Each function that publishes a measured error level is evaluated on a fixed
input grid, compared against a 50-digit
mpmath reference, and checked against its own claim.

## Why two phases

The reference values are computed in Python (mpmath). The library under test is
VBA and can only run inside Excel. So the harness is split: Python owns the
reference and the error analysis; a small Excel macro owns the observed values.
Neither side trusts the other's numbers — Python never sees the library's code,
and the macro never reads the reference column.

## Two test tiers, two roles

Accuracy is protected at two levels with deliberately different tolerances, and
the distinction is intentional:

- **The VBA suite** (`tests/M_STATS_PROBDIST_TEST.bas`) is a fast, deterministic
  regression and public-contract smoke test. It runs entirely inside Excel with
  no high-precision oracle, so it uses broad tolerances (typically 1E-10
  absolute/relative, 1E-9 tail, 1E-6 loose). Its job is to catch regressions and
  confirm the public contract holds across many inputs — not to certify the
  tightest achievable accuracy.
- **This external benchmark** is the measured high-precision accuracy gate. It
  compares the library against 50-digit mpmath references and enforces the
  machine-readable contracts in `accuracy_contracts.csv` (down to 5E-15 where
  claimed). This is where the tight, per-regime accuracy levels are certified.

The VBA tolerances are therefore *broader by design* than several benchmark
contracts. A green VBA suite means "no regression and the public contract holds";
a green benchmark gate means "the measured accuracy contracts are met". Both are
required, and they answer different questions.

## Files

| File | Role |
|---|---|
| `generate_reference_values.py` | Builds the input grid and 50-digit mpmath reference values. Phase 1. |
| `probability_accuracy_grid.csv` | The grid: inputs, reference, and an empty `observed_vba` column. |
| `M_STATS_PROBDIST_ACCURACYEXPORT.bas` | Excel macro that fills `observed_vba` by calling the library. Phase 2. |
| `compute_errors.py` | Joins observed vs reference, finds max-error locations, checks each claim, writes the summary. Phase 3. |
| `accuracy_summary.md` | The generated verdict table. |
| `environment.txt` | Python and dependency versions, reference precision, date. |

## Running it

```
# Phase 1 — reference (Python)
python generate_reference_values.py            # -> probability_accuracy_grid.csv

# Phase 2 — observed (Excel)
#   Import M_STATS_PROBDIST_ACCURACYEXPORT.bas into the workbook,
#   put probability_accuracy_grid.csv beside the workbook,
#   run Export_Accuracy_Observations. It fills the observed_vba column.

# Phase 3 — analysis (Python)
python compute_errors.py                       # -> accuracy_summary.md
```

`compute_errors.py` degrades honestly: any row whose `observed_vba` is still
empty is reported as *not measured* and excluded from the pass/fail check, so a
partial export produces a partial — not a misleading — summary.

> **Observed-value format.** VBA writes each observation as a two-part sum `hi;lo` (two 15-digit numbers), because VBA cannot emit more than ~15 significant digits in one literal. `compute_errors.py` sums the parts to recover the full-precision Double, so the harness measures accuracy below the 15-digit floor. A single number is also accepted for backward compatibility.

## Claims under test

Taken verbatim from the `ACCURACY` comments in the source:

The harness covers five modules: the special-function kernels
(`SPECIALFUNCS`), the test-statistic families (`TFAMILY`), the normal and
lognormal family (`NORMALFAMILY`), the continuous distributions
(`CONTINUOUS`), and the discrete distributions (`DISCRETE` — Binomial,
Poisson, Geometric, Negative Binomial, and Hypergeometric). The
machine-readable contract table below is the authoritative list.

**Special functions**

| Function | Published claim |
|---|---|
| `PROB_LogGamma` | relative error < 6.1E-14 for Z in [1E-8, 1E+50] |
| `PROB_LogGammaHalfDiff` | relative error <= 2E-14 for Z > 0 (tested range) |
| `PROB_StirlingError` | absolute error <= 3E-17 for N >= 0.5 |
| `PROB_LogChoose` | relative error <= 3.2E-16 for N in [2, 2^53], all K |

**Test-statistic families**

| Function | Published claim |
|---|---|
| Student t density | relative error <= 2E-14 (tested range) |
| Student t cumulative / survival | relative error <= 1.3E-12 |
| Student t quantile | relative error <= 3.0E-12 |
| Chi-square cumulative / survival | relative error <= 2.6E-10 |
| Chi-square quantile | relative error <= 4.7E-12 |
| F cumulative / survival | relative error <= 1.1E-10 |
| F quantile | relative error <= 5.9E-13 |

**Normal and lognormal family**

| Function | Published claim |
|---|---|
| Standard normal density / CDF / survival / inverse / inverse survival / interval | relative error <= 5E-15 |
| Standard normal fast inverse (raw Acklam) | relative error <= 5E-9 |
| General normal density / CDF / survival / inverse / inverse survival / z-score | relative error <= 5E-15 |
| Lognormal density / CDF / survival / inverse / inverse survival | relative error <= 5E-15 |
| Lognormal mean / variance / std dev / parameter conversion | relative error <= 5E-15 |

The normal-family `~1E-15` source comments are interpreted as a hard bound of
5E-15 for the harness. `NormalStandard_Survival` measures about 1.5E-14 at
moderate tail values, above that bound; it is reported as *below harness
precision* rather than a failure, and its claim is left as documented.

**Continuous distributions**

Bounds were set from the measured worst-case error over the tested grid, not
from source comments (the module publishes none). Exponential is parameterized
by rate (Lambda), not scale.

<!-- BEGIN generated: accuracy_contracts.csv via render_contract_table.py. Do not hand-edit. -->

| Contract | Function | Regime | Measure | Metric | Threshold | Provenance |
|---|---|---|---|---|---|---|
| Beta_Cumulative.balanced.output_rel | Beta_Cumulative | balanced | output_error | relative | 2E-14 | validated and frozen |
| Beta_Cumulative.unbalanced.output_rel | Beta_Cumulative | unbalanced | output_error | relative | 1E-10 | validated and frozen |
| Beta_Density.balanced.output_rel | Beta_Density | balanced | output_error | relative | 1E-14 | validated and frozen |
| Beta_Density.unbalanced.output_rel | Beta_Density | unbalanced | output_error | relative | 4E-12 | validated and frozen |
| Beta_InverseCumulative.balanced.quantile_rel | Beta_InverseCumulative | balanced | quantile_error | relative | 5E-15 | validated and frozen |
| Beta_InverseCumulative.unbalanced.quantile_rel | Beta_InverseCumulative | unbalanced | quantile_error | relative | 1E-10 | validated and frozen |
| Beta_InverseCumulative.unbalanced.tail_rel | Beta_InverseCumulative | unbalanced | tail_probability_residual | relative | 1E-9 | validated and frozen |
| Beta_Mean.all.output | Beta_Mean | all | output_error | relative | 5E-15 | validated and frozen |
| Beta_StdDev.all.output | Beta_StdDev | all | output_error | relative | 5E-15 | validated and frozen |
| Beta_Survival.balanced.output_rel | Beta_Survival | balanced | output_error | relative | 5E-15 | validated and frozen |
| Beta_Survival.unbalanced.output_rel | Beta_Survival | unbalanced | output_error | relative | 2E-10 | validated and frozen |
| Beta_Variance.all.output | Beta_Variance | all | output_error | relative | 5E-15 | validated and frozen |
| Binomial_Cumulative.all.output_rel | Binomial_Cumulative | all | output_error | relative | 5E-8 | validated and frozen |
| Binomial_InverseCumulative.all.output_abs | Binomial_InverseCumulative | all | output_error | absolute | 1E-9 | validated and frozen |
| Binomial_LogPMF.all.output_rel | Binomial_LogPMF | all | output_error | relative | 1E-13 | validated and frozen |
| Binomial_Mean.all.output_rel | Binomial_Mean | all | output_error | relative | 5E-15 | validated and frozen |
| Binomial_PMF.all.output_rel | Binomial_PMF | all | output_error | relative | 1E-12 | validated and frozen |
| Binomial_StdDev.all.output_rel | Binomial_StdDev | all | output_error | relative | 5E-15 | validated and frozen |
| Binomial_Survival.all.output_rel | Binomial_Survival | all | output_error | relative | 5E-8 | validated and frozen |
| Binomial_Variance.all.output_rel | Binomial_Variance | all | output_error | relative | 5E-15 | validated and frozen |
| ChiSquare_Cumulative.all.output | ChiSquare_Cumulative | all | output_error | relative | 2.6E-10 | validated and frozen |
| ChiSquare_Density.all.output_rel | ChiSquare_Density | all | output_error | relative | 1E-13 | measured provisional |
| ChiSquare_InverseCumulative.all.output | ChiSquare_InverseCumulative | all | output_error | relative | 4.7E-12 | validated and frozen |
| ChiSquare_Survival.all.output | ChiSquare_Survival | all | output_error | relative | 2.6E-10 | validated and frozen |
| DiscreteUniform_Cumulative.all.output_rel | DiscreteUniform_Cumulative | all | output_error | relative | 5E-15 | validated and frozen |
| DiscreteUniform_InverseCumulative.all.output_abs | DiscreteUniform_InverseCumulative | all | output_error | absolute | 1E-9 | validated and frozen |
| DiscreteUniform_LogPMF.all.output_rel | DiscreteUniform_LogPMF | all | output_error | relative | 5E-15 | validated and frozen |
| DiscreteUniform_Mean.all.output_rel | DiscreteUniform_Mean | all | output_error | relative | 5E-15 | validated and frozen |
| DiscreteUniform_PMF.all.output_rel | DiscreteUniform_PMF | all | output_error | relative | 5E-15 | validated and frozen |
| DiscreteUniform_StdDev.all.output_rel | DiscreteUniform_StdDev | all | output_error | relative | 5E-15 | validated and frozen |
| DiscreteUniform_Survival.all.output_rel | DiscreteUniform_Survival | all | output_error | relative | 5E-15 | validated and frozen |
| DiscreteUniform_Variance.all.output_rel | DiscreteUniform_Variance | all | output_error | relative | 5E-15 | validated and frozen |
| Exponential_Cumulative.all.output | Exponential_Cumulative | all | output_error | relative | 5E-15 | validated and frozen |
| Exponential_Density.all.output | Exponential_Density | all | output_error | relative | 5E-15 | validated and frozen |
| Exponential_InverseCumulative.all.output | Exponential_InverseCumulative | all | output_error | relative | 5E-15 | validated and frozen |
| Exponential_Survival.all.output | Exponential_Survival | all | output_error | relative | 5E-15 | validated and frozen |
| F_Cumulative.validated.output_rel | F_Cumulative | validated | output_error | relative | 1.1E-10 | validated and frozen |
| F_Density.all.output_rel | F_Density | all | output_error | relative | 1E-13 | measured provisional |
| F_InverseCumulative.validated.quantile_rel | F_InverseCumulative | validated | quantile_error | relative | 2E-10 | validated and frozen |
| F_InverseCumulative.validated.tail_rel | F_InverseCumulative | validated | tail_probability_residual | relative | 2E-10 | validated and frozen |
| F_Survival.validated.output_rel | F_Survival | validated | output_error | relative | 1.1E-10 | validated and frozen |
| Gamma_Cumulative.all.output | Gamma_Cumulative | all | output_error | relative | 2E-14 | validated and frozen |
| Gamma_Density.all.output | Gamma_Density | all | output_error | relative | 2E-14 | validated and frozen |
| Gamma_InverseCumulative.all.output | Gamma_InverseCumulative | all | output_error | relative | 2E-14 | validated and frozen |
| Gamma_Mean.all.output | Gamma_Mean | all | output_error | relative | 5E-15 | validated and frozen |
| Gamma_StdDev.all.output | Gamma_StdDev | all | output_error | relative | 5E-15 | validated and frozen |
| Gamma_Survival.all.output | Gamma_Survival | all | output_error | relative | 5E-14 | validated and frozen |
| Gamma_Variance.all.output | Gamma_Variance | all | output_error | relative | 5E-15 | validated and frozen |
| Geometric_Cumulative.all.output_rel | Geometric_Cumulative | all | output_error | relative | 5E-15 | validated and frozen |
| Geometric_InverseCumulative.all.output_abs | Geometric_InverseCumulative | all | output_error | absolute | 1E-9 | validated and frozen |
| Geometric_LogPMF.all.output_rel | Geometric_LogPMF | all | output_error | relative | 5E-15 | validated and frozen |
| Geometric_Mean.all.output_rel | Geometric_Mean | all | output_error | relative | 5E-15 | validated and frozen |
| Geometric_PMF.all.output_rel | Geometric_PMF | all | output_error | relative | 2E-15 | validated and frozen |
| Geometric_StdDev.all.output_rel | Geometric_StdDev | all | output_error | relative | 5E-15 | validated and frozen |
| Geometric_Survival.all.output_rel | Geometric_Survival | all | output_error | relative | 5E-15 | validated and frozen |
| Geometric_Variance.all.output_rel | Geometric_Variance | all | output_error | relative | 5E-15 | validated and frozen |
| Hypergeometric_Cumulative.all.output_rel | Hypergeometric_Cumulative | all | output_error | relative | 5E-14 | validated and frozen |
| Hypergeometric_InverseCumulative.all.output_abs | Hypergeometric_InverseCumulative | all | output_error | absolute | 1E-9 | validated and frozen |
| Hypergeometric_LogPMF.all.output_rel | Hypergeometric_LogPMF | all | output_error | relative | 2E-13 | validated and frozen |
| Hypergeometric_Mean.all.output_rel | Hypergeometric_Mean | all | output_error | relative | 5E-15 | validated and frozen |
| Hypergeometric_PMF.all.output_rel | Hypergeometric_PMF | all | output_error | relative | 2E-13 | validated and frozen |
| Hypergeometric_StdDev.all.output_rel | Hypergeometric_StdDev | all | output_error | relative | 5E-15 | validated and frozen |
| Hypergeometric_Survival.all.output_rel | Hypergeometric_Survival | all | output_error | relative | 2E-14 | validated and frozen |
| Hypergeometric_Variance.all.output_rel | Hypergeometric_Variance | all | output_error | relative | 5E-15 | validated and frozen |
| LogChoose.all.output | LogChoose | all | output_error | relative | 3.2E-16 | validated and frozen |
| LogGamma.all.output | LogGamma | all | output_error | relative | 6.1E-14 | validated and frozen |
| LogGammaHalfDiff.all.output | LogGammaHalfDiff | all | output_error | relative | 2E-14 | validated and frozen |
| Lognormal_Cumulative.all.output | Lognormal_Cumulative | all | output_error | relative | 5E-15 | validated and frozen |
| Lognormal_Density.all.output | Lognormal_Density | all | output_error | relative | 3E-14 | validated and frozen |
| Lognormal_InverseCumulative.all.output | Lognormal_InverseCumulative | all | output_error | relative | 5E-15 | validated and frozen |
| Lognormal_InverseSurvival.all.output | Lognormal_InverseSurvival | all | output_error | relative | 5E-15 | validated and frozen |
| Lognormal_Mean.all.output | Lognormal_Mean | all | output_error | relative | 5E-15 | validated and frozen |
| Lognormal_ParametersFromMeanStdDev.param_meanlog.output_rel | Lognormal_ParametersFromMeanStdDev | param_meanlog | output_error | relative | 5E-15 | measured provisional |
| Lognormal_ParametersFromMeanStdDev.param_stddevlog.output_rel | Lognormal_ParametersFromMeanStdDev | param_stddevlog | output_error | relative | 5E-15 | measured provisional |
| Lognormal_StdDev.all.output | Lognormal_StdDev | all | output_error | relative | 5E-15 | validated and frozen |
| Lognormal_Survival.all.output | Lognormal_Survival | all | output_error | relative | 5E-15 | validated and frozen |
| Lognormal_Variance.all.output | Lognormal_Variance | all | output_error | relative | 5E-15 | validated and frozen |
| NegativeBinomial_Cumulative.all.output_rel | NegativeBinomial_Cumulative | all | output_error | relative | 2E-11 | validated and frozen |
| NegativeBinomial_InverseCumulative.all.output_abs | NegativeBinomial_InverseCumulative | all | output_error | absolute | 1E-9 | validated and frozen |
| NegativeBinomial_LogPMF.all.output_rel | NegativeBinomial_LogPMF | all | output_error | relative | 2E-13 | validated and frozen |
| NegativeBinomial_Mean.all.output_rel | NegativeBinomial_Mean | all | output_error | relative | 5E-15 | validated and frozen |
| NegativeBinomial_PMF.all.output_rel | NegativeBinomial_PMF | all | output_error | relative | 2E-12 | validated and frozen |
| NegativeBinomial_StdDev.all.output_rel | NegativeBinomial_StdDev | all | output_error | relative | 5E-15 | validated and frozen |
| NegativeBinomial_Survival.all.output_rel | NegativeBinomial_Survival | all | output_error | relative | 5E-11 | validated and frozen |
| NegativeBinomial_Variance.all.output_rel | NegativeBinomial_Variance | all | output_error | relative | 5E-15 | validated and frozen |
| NormalStandard_Cumulative.all.output | NormalStandard_Cumulative | all | output_error | relative | 5E-15 | validated and frozen |
| NormalStandard_Density.all.output | NormalStandard_Density | all | output_error | relative | 5E-15 | validated and frozen |
| NormalStandard_IntervalProbability.all.output | NormalStandard_IntervalProbability | all | output_error | relative | 5E-15 | validated and frozen |
| NormalStandard_InverseCumulative.all.output | NormalStandard_InverseCumulative | all | output_error | relative | 5E-15 | validated and frozen |
| NormalStandard_InverseCumulativeFast.all.output | NormalStandard_InverseCumulativeFast | all | output_error | relative | 5E-9 | validated and frozen |
| NormalStandard_InverseSurvival.all.output | NormalStandard_InverseSurvival | all | output_error | relative | 5E-15 | validated and frozen |
| NormalStandard_Survival.all.output | NormalStandard_Survival | all | output_error | relative | 2E-14 | validated and frozen |
| Normal_Cumulative.all.output | Normal_Cumulative | all | output_error | relative | 5E-15 | validated and frozen |
| Normal_Density.all.output | Normal_Density | all | output_error | relative | 5E-15 | validated and frozen |
| Normal_IntervalProbability.all.output_rel | Normal_IntervalProbability | all | output_error | relative | 1E-14 | measured provisional |
| Normal_InverseCumulative.all.output | Normal_InverseCumulative | all | output_error | relative | 5E-15 | validated and frozen |
| Normal_InverseSurvival.all.output | Normal_InverseSurvival | all | output_error | relative | 5E-15 | validated and frozen |
| Normal_Survival.all.output | Normal_Survival | all | output_error | relative | 5E-15 | validated and frozen |
| Normal_ZScore.all.output | Normal_ZScore | all | output_error | relative | 5E-15 | validated and frozen |
| PROB_LogBeta.all.log_abs | PROB_LogBeta | all | log_absolute_error | absolute | 2E-13 | validated and frozen |
| Poisson_Cumulative.all.output_rel | Poisson_Cumulative | all | output_error | relative | 2E-10 | validated and frozen |
| Poisson_InverseCumulative.all.output_abs | Poisson_InverseCumulative | all | output_error | absolute | 1E-9 | validated and frozen |
| Poisson_LogPMF.all.output_rel | Poisson_LogPMF | all | output_error | relative | 2E-15 | validated and frozen |
| Poisson_Mean.all.output_rel | Poisson_Mean | all | output_error | relative | 5E-15 | validated and frozen |
| Poisson_PMF.all.output_rel | Poisson_PMF | all | output_error | relative | 1E-14 | validated and frozen |
| Poisson_StdDev.all.output_rel | Poisson_StdDev | all | output_error | relative | 5E-15 | validated and frozen |
| Poisson_Survival.all.output_rel | Poisson_Survival | all | output_error | relative | 5E-10 | validated and frozen |
| Poisson_Variance.all.output_rel | Poisson_Variance | all | output_error | relative | 5E-15 | validated and frozen |
| StirlingError.all.output | StirlingError | all | output_error | absolute | 1E-13 | validated and frozen |
| StudentT_Cumulative.all.output | StudentT_Cumulative | all | output_error | relative | 1.3E-12 | validated and frozen |
| StudentT_Density.all.output | StudentT_Density | all | output_error | relative | 2E-14 | validated and frozen |
| StudentT_InverseCumulative.all.output | StudentT_InverseCumulative | all | output_error | relative | 3.0E-12 | validated and frozen |
| StudentT_Survival.all.output | StudentT_Survival | all | output_error | relative | 1.3E-12 | validated and frozen |
| Uniform_Cumulative.all.output | Uniform_Cumulative | all | output_error | relative | 5E-15 | validated and frozen |
| Uniform_Density.all.output | Uniform_Density | all | output_error | relative | 5E-15 | validated and frozen |
| Uniform_InverseCumulative.all.output | Uniform_InverseCumulative | all | output_error | relative | 5E-15 | validated and frozen |
| Uniform_Survival.all.output | Uniform_Survival | all | output_error | relative | 5E-15 | validated and frozen |
| Weibull_Cumulative.all.output | Weibull_Cumulative | all | output_error | relative | 5E-15 | validated and frozen |
| Weibull_Density.all.output | Weibull_Density | all | output_error | relative | 5E-15 | validated and frozen |
| Weibull_InverseCumulative.all.output | Weibull_InverseCumulative | all | output_error | relative | 5E-15 | validated and frozen |
| Weibull_Mean.all.output | Weibull_Mean | all | output_error | relative | 2E-14 | validated and frozen |
| Weibull_StdDev.all.output | Weibull_StdDev | all | output_error | relative | 5E-15 | validated and frozen |
| Weibull_Survival.all.output | Weibull_Survival | all | output_error | relative | 5E-15 | validated and frozen |
| Weibull_Variance.all.output | Weibull_Variance | all | output_error | relative | 2E-14 | validated and frozen |

**Numerical limitations** (documented, not accuracy contracts)

| Limitation | Affected | Domain | Observed effect | Status |
|---|---|---|---|---|
| IncompleteBeta.ExtremeShape | F_Cumulative;F_Survival;F_InverseCumulative | F degrees of freedom above 1E5 (measured accuracy envelope; the both-large orientation binds, crossing near df 1.5E5-3E5) | without the guard, errors grow to ~4E-7 by df~1E9; the public F surface now REJECTS df > 1E5 with CVErr(xlErrNum) | mitigated |
| SurvivalTailRel | Normal_Survival;Lognormal_Survival;NormalStandard_Survival | upper tail beyond the central region (standardized z above ~2.75-3.25) | relative error grows monotonically: ~1E-14 at z=3, ~1E-13 at z=4, ~5E-11 at z=5, ~5E-10 at z=6 | characterized |


<!-- END generated -->

**Regime-aware contracts.** Each real function may carry several contracts, one
per regime and measure. Balanced Beta/F keep their tight bounds; the strongly
unbalanced regime carries its own measured bound (the stable Lanczos log-gamma
difference removed the earlier catastrophic cancellation in `PROB_LogBeta`).
Inverse functions are judged both on quantile error and on the forward-probability
residual `|F(x_VBA) - p| / min(p, 1-p)`. `PROB_LogBeta` is judged on absolute
error, since downstream code computes `exp(-LogBeta)`. Every unbalanced/inverse
threshold was confirmed on an independent holdout (see `holdout/`) before being
frozen.

**Residual limitation.** The only accuracy limitation outside these contracts is
the incomplete-beta convergence range: for F with a shape parameter beyond
about 1E7, accuracy degrades (up to ~4E-7). This is a pre-existing convergence
limit, *not* the LogBeta normalization, and is recorded in
`numerical_limitations.csv` rather than as a contract it cannot meet.

## Metric note

`PROB_StirlingError` is checked on **absolute** error, not relative: its source
comment explains that relative error is the wrong metric there (it reaches
1.5E-13 near N = 501, where delta itself is 1.67E-04), because what propagates
into a log-probability is the absolute error. The harness honours that choice
per function via the `metric` column.

## Reference integrity

The continued-fraction / mpmath references are cross-checked against SciPy, an
INDEPENDENT implementation, wherever SciPy is reliable (incomplete-beta shape
parameter <= 1E7). This cross-check is a regenerable artifact,
`cross_check_scipy.py` -> `cross_check_scipy.md`, not a bare assertion: at the
current grids 242 points are independently confirmed by SciPy (worst relative
difference about 7E-11, i.e. SciPy's own double-precision limit), and 8 extreme
points lie beyond SciPy's reliable range and instead rest on the reference's
50-vs-80-digit self-consistency. SciPy corroborates the reference where it can
compute; it cannot certify the extreme-parameter points it cannot itself evaluate.
The survival references are computed on the upper tail directly (never `1 - CDF`),
so they stay accurate in the deep tail where a naive subtraction would collapse.
