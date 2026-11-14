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
| `generate_reference_values.py` | Builds 50-digit mpmath reference rows. Report-only by default; never writes the committed grid (#17). |
| `probability_accuracy_grid.csv` | The committed grid: inputs, references, and the Excel observations in `observed_vba`, which cannot be regenerated from Python. |
| `promote_grid_rows.py` | The sanctioned way to add, patch, re-reference or retire grid rows. Report-only unless `--write`; preserves every observation. |
| `migrate_references.py` | Audited, reference-only replacement driven by the reference audit. Report-only unless `--write`. |
| `_grid_safety.py`, `check_grid_keys.py` | Shared canonical-key and write-target safeguards, and the duplicate-key guard run on the committed grids. |
| `_frozen_sources.py` | Production sources at a pinned commit, for closed studies: their replay fixtures and preparation tooling read these, never the working tree (#338). |
| `_row_binding.py`, `check_study_grid_binding.py` | Bind a study to the grid rows it reads (`row_digest`), and fail when a study file hashes or compares a whole live grid without a declared reason (#302). |
| `M_STATS_PROBDIST_ACCURACYEXPORT.bas` | Excel macro that fills `observed_vba` by calling the library. Phase 2. |
| `compute_errors.py` | Joins observed vs reference, finds max-error locations, checks each claim, writes the summary. Phase 3. |
| `accuracy_summary.md` | The generated verdict table. |
| `_ibeta.py` | Continued-fraction incomplete beta / F reference helper, including the direct upper tails `ibeta_upper` and `f_sf`. Single-sourced: this is the only copy, and the study scripts import it from here. |
| `check_survival_references.py`, `_survival_oracle.py` | Fail-closed audit of every committed survival reference against a direct upper-tail recomputation at its binary64 argument (#159). See [Survival reference basis](#survival-reference-basis). |
| `survival_reference_debt.json` | The survival references known to miss the 25-digit floor that a binding prevents re-referencing yet, each with its blocker. |
| `rebase_survival_debt.py` | Re-bases the registered survival references from the source that owns each row, once their bindings can be rebuilt (#159, #81). |
| `_study_merge.py`, `study_reference_rebase.csv` | Reference-only merge of a regenerated study grid into the committed grid, keeping every observation, and the record of every reference it changed (old and new value). |
| `environment.txt` | Python and dependency versions, reference precision, date. |

### Study-folder conventions

Each study subfolder holds one standalone Excel export macro, named to match the
`Attribute VB_Name` it declares (for example `M_STATS_PROBDIST_DELTA_SEAM.bas`
declares `M_STATS_PROBDIST_DELTA_SEAM`). Keep one file per module: a second copy
under a different filename declares the same `VB_Name`, so importing both into one
workbook is a name collision, and two copies of the same source can drift apart.
The same rule is why `_ibeta.py` lives here rather than being copied into each
study folder.

## Running it

Follow the complete [fresh-export procedure](PROVENANCE.md#fresh-export-operating-procedure): certify the exact candidate, export grids and environment, finalize the canonical record, bind the selected manifests, regenerate summaries, and validate committed evidence. The current release blockers are listed in the [README assurance block](../README.md#assurance-at-a-glance), generated from `release_readiness.json`.

The abbreviated commands below run from `benchmark/`; they do not replace certification and binding.

```bash
cd benchmark
# Phase 1 — reference (Python). Report only: compares with the committed
# grid and writes nothing. --out writes the reference rows to a separate,
# non-authoritative file and refuses any file that carries observations.
python generate_reference_values.py
python generate_reference_values.py --out reference_rows.csv

# Changing the committed grid goes through promote_grid_rows.py (add rows,
# patch metadata, accept a counted and reasoned set of reference changes,
# retire rows) or migrate_references.py (audited references). Both are
# report-only unless --write. Full clean regeneration is #32, not a v1.0.0 path.

# Phase 2 — observed (Excel)
#   Import M_STATS_PROBDIST_ACCURACYEXPORT.bas into the workbook,
#   put probability_accuracy_grid.csv beside the workbook,
#   run Export_Accuracy_Observations. It fills the observed_vba column.
#   Unattended alternative: dispatch excel-vba-regression.yml with the
#   export_* inputs, then from the repository root run
#   python benchmark/apply_excel_export.py --artifact <dir> --grid main ...
#   which also finalizes certification and binds the exported grids
#   (docs/EXCEL_CERTIFICATION.md, "Automated fresh export").

# Before analysis: finalize certification and bind only the grids really
# exported, following PROVENANCE.md (its commands use the repository root).

# Phase 3 — analysis (Python)
python compute_errors.py                       # -> accuracy_summary.md
```

`compute_errors.py` is strict: every in-envelope row of an active contract must
be produced, parsed, and evaluated. A blank, `ERROR`, or unparseable observation
— for ordinary *and* tail-residual contracts alike — blocks that contract as
`PENDING` and fails the gate (non-zero exit), naming the offending rows. Only
envelope-reject rows that correctly return `#NUM!` are excluded from scoring. A
partial or corrupt export therefore fails loudly rather than producing a
misleading green summary.

> **Observed-value format.** VBA writes each observation as a two-part sum `hi;lo` (two 15-digit numbers), because VBA cannot emit more than ~15 significant digits in one literal. `compute_errors.py` sums the parts to recover the full-precision Double, so the harness measures accuracy below the 15-digit floor. A single number is also accepted for backward compatibility.

## Claims under test

The harness covers five modules: the special-function kernels
(`SPECIALFUNCS`), the test-statistic families (`TFAMILY`), the normal and
lognormal family (`NORMALFAMILY`), the continuous distributions
(`CONTINUOUS`), and the discrete distributions (`DISCRETE`). The generated
block below is the authoritative list: it names every contract and every
documented numerical limitation, and is derived from `accuracy_contracts.csv`
and `numerical_limitations.csv`, so it cannot drift from the evidence.

Where a contract comes from:

- The special-function, test-statistic and normal-family claims started as the
  `ACCURACY` comments in the source. Several functions now carry more than one
  contract, one per regime (for example F quantile `validated` at 2E-10, and
  Student t `envelope_domain` at 5E-15), so read the table, not the comments.
- The normal-family `~1E-15` source comments are interpreted as a hard bound of
  5E-15 for the harness.
- The continuous and discrete bounds were set from the measured worst-case
  error over the tested grid; the Beta source comment points back to this
  registry rather than stating a number. Exponential is parameterized by rate
  (Lambda), not scale.

<!-- BEGIN generated: accuracy_contracts.csv via render_contract_table.py. Do not hand-edit. -->

**177 contracts** across **23 function families**: Beta, Binomial, ChiSquare, DiscreteUniform, Exponential, F, Gamma, Geometric, Hypergeometric, LogChoose, LogGamma, LogGamma1p, LogGammaDelta, LogGammaHalfDiff, Lognormal, NegativeBinomial, Normal, NormalStandard, Poisson, StirlingError, StudentT, Uniform, Weibull. Kernel-level contracts: PROB_LogBeta.

The **guarantee** gates the release; the **measured target** is the registry threshold, reported beside it ([accuracy_tiers.py](accuracy_tiers.py)).

| Contract | Function | Regime | Measure | Metric | Guarantee | Measured target | Provenance |
|---|---|---|---|---|---|---|---|
| Beta_Cumulative.balanced.output_rel | Beta_Cumulative | balanced | output_error | relative | 2E-13 | 2E-14 | validated and frozen |
| Beta_Cumulative.extreme_shape.output_rel | Beta_Cumulative | extreme_shape | output_error | relative | 5E-10 | 5E-10 | validated and frozen |
| Beta_Cumulative.large_shape.output_rel | Beta_Cumulative | large_shape | output_error | relative | 1E-12 | 1E-12 | validated and frozen |
| Beta_Cumulative.tiny_unbalanced.output | Beta_Cumulative | tiny_unbalanced | output_error | relative | 5E-13 | 5E-13 | validated and frozen |
| Beta_Cumulative.unbalanced.output_rel | Beta_Cumulative | unbalanced | output_error | relative | 1E-10 | 1E-10 | validated and frozen |
| Beta_Density.balanced.output_rel | Beta_Density | balanced | output_error | relative | 1E-13 | 1E-14 | validated and frozen |
| Beta_Density.large_shape.output_rel | Beta_Density | large_shape | output_error | relative | 1E-5 | 1E-5 | validated and frozen |
| Beta_Density.tiny_unbalanced.output | Beta_Density | tiny_unbalanced | output_error | relative | 1E-13 | 5E-14 | validated and frozen |
| Beta_Density.unbalanced.output_rel | Beta_Density | unbalanced | output_error | relative | 4E-12 | 4E-12 | validated and frozen |
| Beta_InverseCumulative.balanced.quantile_rel | Beta_InverseCumulative | balanced | quantile_error | relative | 1E-12 | 5E-15 | validated and frozen |
| Beta_InverseCumulative.tiny_unbalanced_representable.tail_rel | Beta_InverseCumulative | tiny_unbalanced_representable | tail_probability_residual | relative | 1E-14 | 1E-14 | validated and frozen |
| Beta_InverseCumulative.unbalanced.quantile_rel | Beta_InverseCumulative | unbalanced | quantile_error | relative | 1E-10 | 1E-10 | validated and frozen |
| Beta_InverseCumulative.unbalanced.tail_rel | Beta_InverseCumulative | unbalanced | tail_probability_residual | relative | 1E-9 | 1E-9 | validated and frozen |
| Beta_Mean.all.output | Beta_Mean | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Beta_StdDev.all.output | Beta_StdDev | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Beta_Survival.balanced.output_rel | Beta_Survival | balanced | output_error | relative | 2E-13 | 5E-15 | validated and frozen |
| Beta_Survival.extreme_shape.output_rel | Beta_Survival | extreme_shape | output_error | relative | 5E-10 | 5E-10 | validated and frozen |
| Beta_Survival.large_shape.output_rel | Beta_Survival | large_shape | output_error | relative | 1E-12 | 1E-12 | validated and frozen |
| Beta_Survival.tiny_unbalanced.output | Beta_Survival | tiny_unbalanced | output_error | relative | 1E-13 | 5E-14 | validated and frozen |
| Beta_Survival.unbalanced.output_rel | Beta_Survival | unbalanced | output_error | relative | 2E-10 | 2E-10 | validated and frozen |
| Beta_Variance.all.output | Beta_Variance | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Binomial_Cumulative.all.output_rel | Binomial_Cumulative | all | output_error | relative | 5E-8 | 5E-8 | validated and frozen |
| Binomial_InverseCumulative.all.output_abs | Binomial_InverseCumulative | all | output_error | absolute | 1E-9 | 1E-9 | validated and frozen |
| Binomial_LogPMF.all.output_rel | Binomial_LogPMF | all | output_error | relative | 1E-13 | 1E-13 | validated and frozen |
| Binomial_Mean.all.output_rel | Binomial_Mean | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Binomial_PMF.all.output_rel | Binomial_PMF | all | output_error | relative | 1E-12 | 1E-12 | validated and frozen |
| Binomial_StdDev.all.output_rel | Binomial_StdDev | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Binomial_Survival.all.output_rel | Binomial_Survival | all | output_error | relative | 5E-8 | 5E-8 | validated and frozen |
| Binomial_Variance.all.output_rel | Binomial_Variance | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| ChiSquare_Cumulative.all.output | ChiSquare_Cumulative | all | output_error | relative | 2.6E-10 | 2.6E-10 | validated and frozen |
| ChiSquare_Cumulative.envelope_domain.output_rel | ChiSquare_Cumulative | envelope_domain | output_error | relative | 2E-12 | 2E-12 | validated and frozen |
| ChiSquare_Cumulative.positive_ratio_subnormal.output_rel | ChiSquare_Cumulative | positive_ratio_subnormal | output_error | relative | 5E-13 | 5E-13 | validated and frozen |
| ChiSquare_Density.all.output_rel | ChiSquare_Density | all | output_error | relative | 1E-13 | 1E-13 | validated and frozen |
| ChiSquare_Density.large_shape.output_rel | ChiSquare_Density | large_shape | output_error | relative | 1E-13 | 1E-13 | validated and frozen |
| ChiSquare_Density.positive_ratio_subnormal.output_rel | ChiSquare_Density | positive_ratio_subnormal | output_error | relative | 5E-13 | 5E-13 | validated and frozen |
| ChiSquare_InverseCumulative.all.output | ChiSquare_InverseCumulative | all | output_error | relative | 4.7E-12 | 4.7E-12 | validated and frozen |
| ChiSquare_InverseCumulative.scaled_quantile_from_underflowed_unit.quantile_rel | ChiSquare_InverseCumulative | scaled_quantile_from_underflowed_unit | quantile_error | relative | 1E-12 | 1E-14 | validated and frozen |
| ChiSquare_InverseCumulative.scaled_quantile_from_underflowed_unit.tail_rel | ChiSquare_InverseCumulative | scaled_quantile_from_underflowed_unit | tail_probability_residual | relative | 1E-14 | 1E-14 | validated and frozen |
| ChiSquare_Survival.all.output | ChiSquare_Survival | all | output_error | relative | 2.6E-10 | 2.6E-10 | validated and frozen |
| ChiSquare_Survival.envelope_domain.output_rel | ChiSquare_Survival | envelope_domain | output_error | relative | 2E-12 | 2E-12 | validated and frozen |
| ChiSquare_Survival.positive_ratio_hard_underflow.output_rel | ChiSquare_Survival | positive_ratio_hard_underflow | output_error | relative | 5E-13 | 5E-13 | validated and frozen |
| DiscreteUniform_Cumulative.all.output_rel | DiscreteUniform_Cumulative | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| DiscreteUniform_InverseCumulative.all.output_abs | DiscreteUniform_InverseCumulative | all | output_error | absolute | 1E-9 | 1E-9 | validated and frozen |
| DiscreteUniform_LogPMF.all.output_rel | DiscreteUniform_LogPMF | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| DiscreteUniform_Mean.all.output_rel | DiscreteUniform_Mean | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| DiscreteUniform_PMF.all.output_rel | DiscreteUniform_PMF | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| DiscreteUniform_StdDev.all.output_rel | DiscreteUniform_StdDev | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| DiscreteUniform_Survival.all.output_rel | DiscreteUniform_Survival | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| DiscreteUniform_Variance.all.output_rel | DiscreteUniform_Variance | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Exponential_Cumulative.all.output | Exponential_Cumulative | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Exponential_Density.all.output | Exponential_Density | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Exponential_InverseCumulative.all.output | Exponential_InverseCumulative | all | output_error | relative | 1E-12 | 5E-15 | validated and frozen |
| Exponential_Survival.all.output | Exponential_Survival | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| F_Cumulative.envelope_domain.output_rel | F_Cumulative | envelope_domain | output_error | relative | 2E-11 | 2E-11 | validated and frozen |
| F_Cumulative.tiny_unbalanced.output | F_Cumulative | tiny_unbalanced | output_error | relative | 5E-13 | 5E-13 | validated and frozen |
| F_Cumulative.validated.output_rel | F_Cumulative | validated | output_error | relative | 1.1E-10 | 1.1E-10 | validated and frozen |
| F_Density.all.output_rel | F_Density | all | output_error | relative | 1E-13 | 1E-13 | validated and frozen |
| F_Density.extreme_ratio.output_rel | F_Density | extreme_ratio | output_error | relative | 5E-12 | 5E-12 | validated and frozen |
| F_Density.large_shape.output_rel | F_Density | large_shape | output_error | relative | 5E-5 | 5E-5 | validated and frozen |
| F_Density.tiny_unbalanced.output | F_Density | tiny_unbalanced | output_error | relative | 1E-13 | 5E-14 | validated and frozen |
| F_InverseCumulative.tiny_unbalanced_representable.tail_rel | F_InverseCumulative | tiny_unbalanced_representable | tail_probability_residual | relative | 1E-13 | 1E-13 | validated and frozen |
| F_InverseCumulative.validated.quantile_rel | F_InverseCumulative | validated | quantile_error | relative | 2E-10 | 2E-10 | validated and frozen |
| F_InverseCumulative.validated.tail_rel | F_InverseCumulative | validated | tail_probability_residual | relative | 2E-10 | 2E-10 | validated and frozen |
| F_Survival.envelope_domain.output_rel | F_Survival | envelope_domain | output_error | relative | 2E-11 | 2E-11 | validated and frozen |
| F_Survival.tiny_unbalanced.output | F_Survival | tiny_unbalanced | output_error | relative | 1E-13 | 5E-14 | validated and frozen |
| F_Survival.validated.output_rel | F_Survival | validated | output_error | relative | 1.1E-10 | 1.1E-10 | validated and frozen |
| Gamma_Cumulative.all.output | Gamma_Cumulative | all | output_error | relative | 1E-13 | 3E-15 | validated and frozen |
| Gamma_Cumulative.extreme_shape.output_rel | Gamma_Cumulative | extreme_shape | output_error | relative | 2E-12 | 2E-12 | validated and frozen |
| Gamma_Cumulative.large_shape.output_rel | Gamma_Cumulative | large_shape | output_error | relative | 2E-13 | 2E-13 | validated and frozen |
| Gamma_Cumulative.positive_ratio_subnormal.output_rel | Gamma_Cumulative | positive_ratio_subnormal | output_error | relative | 5E-13 | 5E-13 | validated and frozen |
| Gamma_Density.all.output | Gamma_Density | all | output_error | relative | 1E-13 | 2E-14 | validated and frozen |
| Gamma_Density.large_shape.output_rel | Gamma_Density | large_shape | output_error | relative | 2E-13 | 2E-13 | validated and frozen |
| Gamma_Density.positive_ratio_subnormal.output_rel | Gamma_Density | positive_ratio_subnormal | output_error | relative | 5E-13 | 5E-13 | validated and frozen |
| Gamma_InverseCumulative.all.output | Gamma_InverseCumulative | all | output_error | relative | 1E-12 | 2E-14 | validated and frozen |
| Gamma_InverseCumulative.scaled_quantile_from_underflowed_unit.quantile_rel | Gamma_InverseCumulative | scaled_quantile_from_underflowed_unit | quantile_error | relative | 1E-12 | 1E-14 | validated and frozen |
| Gamma_InverseCumulative.scaled_quantile_from_underflowed_unit.tail_rel | Gamma_InverseCumulative | scaled_quantile_from_underflowed_unit | tail_probability_residual | relative | 1E-14 | 1E-14 | validated and frozen |
| Gamma_Mean.all.output | Gamma_Mean | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Gamma_StdDev.all.output | Gamma_StdDev | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Gamma_Survival.all.output | Gamma_Survival | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Gamma_Survival.extreme_shape.output_rel | Gamma_Survival | extreme_shape | output_error | relative | 5E-13 | 5E-13 | validated and frozen |
| Gamma_Survival.large_shape.output_rel | Gamma_Survival | large_shape | output_error | relative | 2E-13 | 2E-13 | validated and frozen |
| Gamma_Survival.positive_ratio_hard_underflow.output_rel | Gamma_Survival | positive_ratio_hard_underflow | output_error | relative | 5E-13 | 5E-13 | validated and frozen |
| Gamma_Variance.all.output | Gamma_Variance | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Geometric_Cumulative.all.output_rel | Geometric_Cumulative | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Geometric_InverseCumulative.all.output_abs | Geometric_InverseCumulative | all | output_error | absolute | 1E-9 | 1E-9 | validated and frozen |
| Geometric_LogPMF.all.output_rel | Geometric_LogPMF | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Geometric_Mean.all.output_rel | Geometric_Mean | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Geometric_PMF.all.output_rel | Geometric_PMF | all | output_error | relative | 1E-13 | 2E-15 | validated and frozen |
| Geometric_StdDev.all.output_rel | Geometric_StdDev | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Geometric_Survival.all.output_rel | Geometric_Survival | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Geometric_Variance.all.output_rel | Geometric_Variance | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Hypergeometric_Cumulative.all.output_rel | Hypergeometric_Cumulative | all | output_error | relative | 1E-13 | 5E-14 | validated and frozen |
| Hypergeometric_InverseCumulative.all.output_abs | Hypergeometric_InverseCumulative | all | output_error | absolute | 1E-9 | 1E-9 | validated and frozen |
| Hypergeometric_LogPMF.all.output_rel | Hypergeometric_LogPMF | all | output_error | relative | 2E-13 | 2E-13 | validated and frozen |
| Hypergeometric_Mean.all.output_rel | Hypergeometric_Mean | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Hypergeometric_PMF.all.output_rel | Hypergeometric_PMF | all | output_error | relative | 2E-13 | 2E-13 | validated and frozen |
| Hypergeometric_StdDev.all.output_rel | Hypergeometric_StdDev | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Hypergeometric_Survival.all.output_rel | Hypergeometric_Survival | all | output_error | relative | 1E-13 | 2E-14 | validated and frozen |
| Hypergeometric_Variance.all.output_rel | Hypergeometric_Variance | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| LogChoose.all.output | LogChoose | all | output_error | relative | 1E-13 | 3.2E-16 | validated and frozen |
| LogGamma.general.output_rel | LogGamma | general | output_error | relative | 5E-13 | 5E-13 | validated and frozen |
| LogGamma.near_zero.log_abs | LogGamma | near_zero | output_error | absolute | 5E-14 | 5E-14 | validated and frozen |
| LogGamma.small_positive.log_abs | LogGamma | small_positive | output_error | absolute | 5E-13 | 5E-13 | validated and frozen |
| LogGamma1p.series_seam.scaled_abs | LogGamma1p | series_seam | scaled_output_error | absolute |  |  | characterized |
| LogGamma1p.small.scaled_abs | LogGamma1p | small | scaled_output_error | absolute | 5E-16 | 5E-16 | validated and frozen |
| LogGammaDelta.all.output | LogGammaDelta | all | output_error | relative | 1E-13 | 5E-14 | validated and frozen |
| LogGammaHalfDiff.all.output | LogGammaHalfDiff | all | output_error | relative | 1E-13 | 2E-14 | validated and frozen |
| Lognormal_Cumulative.all.output | Lognormal_Cumulative | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Lognormal_Density.all.output | Lognormal_Density | all | output_error | relative | 1E-13 | 3E-14 | validated and frozen |
| Lognormal_InverseCumulative.all.output | Lognormal_InverseCumulative | all | output_error | relative | 1E-12 | 5E-15 | validated and frozen |
| Lognormal_InverseSurvival.all.output | Lognormal_InverseSurvival | all | output_error | relative | 1E-12 | 5E-15 | validated and frozen |
| Lognormal_InverseSurvival.deep_tail.output_rel | Lognormal_InverseSurvival | deep_tail | output_error | relative | 1E-12 | 1E-14 | validated and frozen |
| Lognormal_InverseSurvival.split_boundary.output_rel | Lognormal_InverseSurvival | split_boundary | output_error | relative | 1E-9 | 1E-9 | validated and frozen |
| Lognormal_Mean.all.output | Lognormal_Mean | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Lognormal_ParametersFromMeanStdDev.param_meanlog.output_rel | Lognormal_ParametersFromMeanStdDev | param_meanlog | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Lognormal_ParametersFromMeanStdDev.param_stddevlog.output_rel | Lognormal_ParametersFromMeanStdDev | param_stddevlog | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Lognormal_StdDev.all.output | Lognormal_StdDev | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Lognormal_Survival.all.output | Lognormal_Survival | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Lognormal_Variance.all.output | Lognormal_Variance | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| NegativeBinomial_Cumulative.all.output_rel | NegativeBinomial_Cumulative | all | output_error | relative | 2E-11 | 2E-11 | validated and frozen |
| NegativeBinomial_InverseCumulative.all.output_abs | NegativeBinomial_InverseCumulative | all | output_error | absolute | 1E-9 | 1E-9 | validated and frozen |
| NegativeBinomial_LogPMF.all.output_rel | NegativeBinomial_LogPMF | all | output_error | relative | 2E-13 | 2E-13 | validated and frozen |
| NegativeBinomial_Mean.all.output_rel | NegativeBinomial_Mean | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| NegativeBinomial_PMF.all.output_rel | NegativeBinomial_PMF | all | output_error | relative | 2E-12 | 2E-12 | validated and frozen |
| NegativeBinomial_StdDev.all.output_rel | NegativeBinomial_StdDev | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| NegativeBinomial_Survival.all.output_rel | NegativeBinomial_Survival | all | output_error | relative | 5E-11 | 5E-11 | validated and frozen |
| NegativeBinomial_Variance.all.output_rel | NegativeBinomial_Variance | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| NormalStandard_Cumulative.all.output | NormalStandard_Cumulative | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| NormalStandard_Density.all.output | NormalStandard_Density | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| NormalStandard_IntervalProbability.all.output | NormalStandard_IntervalProbability | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| NormalStandard_InverseCumulative.all.output | NormalStandard_InverseCumulative | all | output_error | relative | 1E-12 | 5E-15 | validated and frozen |
| NormalStandard_InverseCumulative.deep_tail.output_rel | NormalStandard_InverseCumulative | deep_tail | output_error | relative | 1E-12 | 5E-16 | validated and frozen |
| NormalStandard_InverseCumulative.split_boundary.output_rel | NormalStandard_InverseCumulative | split_boundary | output_error | relative | 1E-10 | 1E-10 | validated and frozen |
| NormalStandard_InverseCumulativeFast.all.output | NormalStandard_InverseCumulativeFast | all | output_error | relative | 5E-9 | 5E-9 | validated and frozen |
| NormalStandard_InverseSurvival.all.output | NormalStandard_InverseSurvival | all | output_error | relative | 1E-12 | 5E-15 | validated and frozen |
| NormalStandard_InverseSurvival.deep_tail.output_rel | NormalStandard_InverseSurvival | deep_tail | output_error | relative | 1E-12 | 5E-16 | validated and frozen |
| NormalStandard_InverseSurvival.split_boundary.output_rel | NormalStandard_InverseSurvival | split_boundary | output_error | relative | 1E-10 | 1E-10 | validated and frozen |
| NormalStandard_Survival.all.output | NormalStandard_Survival | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Normal_Cumulative.all.output | Normal_Cumulative | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Normal_Density.all.output | Normal_Density | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Normal_IntervalProbability.all.output_rel | Normal_IntervalProbability | all | output_error | relative | 1E-13 | 1E-14 | validated and frozen |
| Normal_InverseCumulative.all.output | Normal_InverseCumulative | all | output_error | relative | 1E-12 | 5E-15 | validated and frozen |
| Normal_InverseSurvival.all.output | Normal_InverseSurvival | all | output_error | relative | 1E-12 | 5E-15 | validated and frozen |
| Normal_InverseSurvival.deep_tail.output_rel | Normal_InverseSurvival | deep_tail | output_error | relative | 1E-12 | 5E-16 | validated and frozen |
| Normal_InverseSurvival.split_boundary.output_rel | Normal_InverseSurvival | split_boundary | output_error | relative | 2E-10 | 2E-10 | validated and frozen |
| Normal_Survival.all.output | Normal_Survival | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Normal_ZScore.all.output | Normal_ZScore | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| PROB_LogBeta.all.log_abs | PROB_LogBeta | all | log_absolute_error | absolute | 2E-13 | 2E-13 | validated and frozen |
| PROB_LogBeta.tiny_unbalanced.output | PROB_LogBeta | tiny_unbalanced | log_absolute_error | absolute | 5E-14 | 5E-14 | validated and frozen |
| Poisson_Cumulative.all.output_rel | Poisson_Cumulative | all | output_error | relative | 2E-13 | 2E-13 | validated and frozen |
| Poisson_InverseCumulative.all.output_abs | Poisson_InverseCumulative | all | output_error | absolute | 1E-9 | 1E-9 | validated and frozen |
| Poisson_LogPMF.all.output_rel | Poisson_LogPMF | all | output_error | relative | 1E-13 | 2E-15 | validated and frozen |
| Poisson_Mean.all.output_rel | Poisson_Mean | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Poisson_PMF.all.output_rel | Poisson_PMF | all | output_error | relative | 1E-13 | 1E-14 | validated and frozen |
| Poisson_StdDev.all.output_rel | Poisson_StdDev | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Poisson_Survival.all.output_rel | Poisson_Survival | all | output_error | relative | 2E-13 | 2E-13 | validated and frozen |
| Poisson_Variance.all.output_rel | Poisson_Variance | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| StirlingError.all.output | StirlingError | all | output_error | absolute | 1E-13 | 1E-13 | validated and frozen |
| StudentT_Cumulative.all.output | StudentT_Cumulative | all | output_error | relative | 1.3E-12 | 1.3E-12 | validated and frozen |
| StudentT_Cumulative.envelope_domain.output_rel | StudentT_Cumulative | envelope_domain | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| StudentT_Density.all.output | StudentT_Density | all | output_error | relative | 1E-13 | 2E-14 | validated and frozen |
| StudentT_Density.large_df.output_rel | StudentT_Density | large_df | output_error | relative | 5E-12 | 5E-12 | validated and frozen |
| StudentT_InverseCumulative.all.output | StudentT_InverseCumulative | all | output_error | relative | 3.0E-12 | 3.0E-12 | validated and frozen |
| StudentT_InverseCumulative.envelope_domain.quantile_rel | StudentT_InverseCumulative | envelope_domain | quantile_error | relative | 1E-12 | 5E-15 | validated and frozen |
| StudentT_Survival.all.output | StudentT_Survival | all | output_error | relative | 1.3E-12 | 1.3E-12 | validated and frozen |
| StudentT_Survival.envelope_domain.output_rel | StudentT_Survival | envelope_domain | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Uniform_Cumulative.all.output | Uniform_Cumulative | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Uniform_Density.all.output | Uniform_Density | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Uniform_InverseCumulative.all.output | Uniform_InverseCumulative | all | output_error | relative | 1E-12 | 5E-15 | validated and frozen |
| Uniform_Survival.all.output | Uniform_Survival | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Weibull_Cumulative.all.output | Weibull_Cumulative | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Weibull_Density.all.output | Weibull_Density | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Weibull_InverseCumulative.all.output | Weibull_InverseCumulative | all | output_error | relative | 1E-12 | 5E-15 | validated and frozen |
| Weibull_Mean.all.output | Weibull_Mean | all | output_error | relative | 1E-13 | 2E-14 | validated and frozen |
| Weibull_StdDev.all.output | Weibull_StdDev | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Weibull_Survival.all.output | Weibull_Survival | all | output_error | relative | 1E-13 | 5E-15 | validated and frozen |
| Weibull_Variance.all.output | Weibull_Variance | all | output_error | relative | 1E-13 | 2E-14 | validated and frozen |

**Numerical limitations** (documented, not accuracy contracts)

| Limitation | Affected | Domain | Observed effect | Status |
|---|---|---|---|---|
| IncompleteBeta.ExtremeShape | F_Cumulative;F_Survival;F_InverseCumulative | F degrees of freedom above 1E10 (measured accuracy envelope; raised from 1E5 after CR-P1-02, benchmark/envelope_probe; the both-large orientation binds, crossing near df 1.5E5-3E5) | the original 1E5 boundary was measuring the CR-P1-02 prefactor cancellation rather than an intrinsic limit; re-measured in benchmark/envelope_probe the kernel is accurate to 5.1E-10 relative at df 1E10, and the public F surface now REJECTS df above PROB_F_MAX_DF with CVErr(xlErrNum) | mitigated |
| SurvivalTailRel | Normal_Survival;Lognormal_Survival;NormalStandard_Survival | upper tail beyond the central region (standardized z above ~2.75-3.25) | relative error grows monotonically: ~1E-14 at z=3, ~1E-13 at z=4, ~5E-11 at z=5, ~5E-10 at z=6 | resolved |
| BetaInverse.InteriorQuantileRepresentability | Beta_InverseCumulative | both shapes sufficiently small (roughly a, b below 1E-3), interior probabilities | No representable interior quantile or meaningful forward round-trip exists; e.g. at a=1E-12, b=9.9E-14 the interior CDF is flat at ~0.0901 across the whole representable interval and the p=0.5 quantile needs 1-x ~ 1E-1E13. | characterized |
| FInverse.InteriorQuantileRepresentability | F_InverseCumulative | both degrees of freedom sufficiently small, interior probabilities | The interior CDF plateaus and the mathematical quantile leaves the representable Double range for probabilities away from the plateau. | characterized |
| IncompleteBeta.StudentTLargeDF | StudentT_Cumulative;StudentT_Survival;StudentT_InverseCumulative | Student t degrees of freedom above 1E8 (measured envelope; raised from 1E6 after CR-P1-02, benchmark/envelope_probe) | measured 1.2E-09 relative at df 1E8 and degrading roughly tenfold per decade beyond; the public t surface REJECTS df > 1E8 with CVErr(xlErrNum) | mitigated |
| IncompleteBeta.StudentTLargeDFFarTail | StudentT_Cumulative;StudentT_Survival;StudentT_InverseCumulative | Student t with df from 1E5 to 1E8 and |t| above 8, outside the large-df expansion's region (#34) | about 4.7E-10 relative at t = 8 + 1 ulp, df = 1E8, where the tail is about 6E-16 (modeled, benchmark/student_t_large_df_study/production_predictions.json seams); 8.7E-13 at df = 1E5; the Excel regression holds t = 8 + 1 ulp, df = 1E8 within 1E-9 | characterized |
| StudentT.LargeDFSeamMonotonicity | StudentT_Cumulative;StudentT_Survival | the Doubles on either side of the large-df seams: df = 1E5 (|t| <= 8) and |t| = 8 (df 1E5..1E8) (#210) | survival and cumulative are not monotone across the seams: up to 3.16E-12 relative across df = 1E5 (t = 2) and 8.1E-11 across |t| = 8 (df = 1E7), in Excel; inside the expansion's region they are monotone to within 5.4E-16. Complement (survival + cumulative - 1) is within 0.5 * 2^-53 and survival(t) equals cumulative(-t) bit for bit at all 52 seam pairs | characterized |
| IncompleteGamma.ChiSquareLargeDF | ChiSquare_Cumulative;ChiSquare_Survival;ChiSquare_InverseCumulative | Chi-square degrees of freedom above 1E8 (measured envelope; raised from 1E6 after CR-P1-02, benchmark/envelope_probe) | measured 3.1E-11 relative at df 1E8; the P2-02 divergence at df 1E16 was the CR-P1-02 prefactor cancellation, not an intrinsic limit; the public chi-square surface REJECTS df > 1E8 with CVErr(xlErrNum) | mitigated |
| LogGamma1p.SubnormalResultRepresentability | PROB_TryLogGamma1p | X below 3.854839696505424E-308, that is PROB_MIN_NORMAL / EulerGamma | scaled absolute error grows toward the half-ulp bound 0.5 * minimum_subnormal / X: measured 1.92E-16 at X = 1E-308, 1.29E-14 at 1E-310, 1.41E-04 at 1E-320 and 4.18E-01 at the smallest positive subnormal | characterized |


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

**Limitations outside the contracts.** Numerical boundaries that no contract
claims, such as the degrees-of-freedom envelopes the public F, Student t and
Chi-square functions reject with `#NUM!`, are recorded in
`numerical_limitations.csv` and rendered in the *Numerical limitations* table
above. That file, not this prose, is the list.

## Metric note

`PROB_StirlingError` is checked on **absolute** error, not relative: its source
comment explains that relative error is the wrong metric there (it reaches
1.5E-13 near N = 501, where delta itself is 1.67E-04), because what propagates
into a log-probability is the absolute error. The harness honours that choice
per function via the `metric` column.

## Reference integrity

The references in three grids are cross-checked against SciPy, an INDEPENDENT
implementation. This cross-check is a regenerable artifact,
`cross_check_scipy.py` -> `cross_check_scipy.md`, not a bare assertion. The
committed record covers `beta_f_unbalanced/beta_f_unbalanced_grid.csv`,
`beta_f_unbalanced/beta_f_inverse_grid.csv` and `holdout/holdout_grid.csv` as
they stood at commit `ce9a274` (2026-11-05), after the #159 re-reference of
`beta_f_unbalanced`. It lists the
SHA-256 of each grid and the SciPy, NumPy and mpmath versions it ran with.

**The record is dated.** Its two `beta_f_unbalanced` digests still match the
committed grids; its holdout digest (`0de20929…`, 559 rows) does not. Since
then the #81 wave re-based 15 holdout survival references and #34 added 100
Student-t holdout rows, for which the script has no SciPy mapping. The record
covers neither. Regenerating it is #154. The counts below describe the record
as it stands.

Each of the 675 reference rows in the record is in exactly one class:

- **655 confirmed by SciPy.** The function has an exact SciPy equivalent and
  the point lies where SciPy is reliable (incomplete-beta shape parameter
  <= 1E7 for Beta and F). They cover Beta, F, `PROB_LogBeta`, the chi-square
  and F densities, the normal and lognormal inverses, and every Binomial,
  Poisson, Geometric, Negative Binomial, Hypergeometric and Discrete Uniform
  function in the grids. The largest relative differences are
  `Binomial_LogPMF` 2.3E-10, `Poisson_PMF` 1.6E-10, `PROB_LogBeta` 7.1E-11,
  `Poisson_LogPMF` 1.9E-11 and `Hypergeometric_LogPMF` 1.5E-11; every other
  function agrees within 2.0E-12. Agreement corroborates a reference only to
  the difference shown, and the record does not say which side carries it.
- **8 beyond SciPy's reliable range**: `F_Cumulative` and `F_Survival` at
  df2 = 1E8 and 1E10, where the incomplete-beta shape parameter is above 1E7.
  They are **not independently confirmed** by SciPy. The four `F_Survival`
  rows are recomputed by a direct route in `check_survival_references.py`
  (see [Survival reference basis](#survival-reference-basis)); no committed
  tool checks the four `F_Cumulative` references.
- **12 not mapped**: `Lognormal_ParametersFromMeanStdDev` (8 rows) and
  `Normal_IntervalProbability` (4 rows) have no exact SciPy equivalent, so
  they were never compared and are **not independently confirmed** here.

SciPy corroborates the reference where it can compute; it cannot certify the
extreme-parameter points it cannot itself evaluate, or functions the script
does not map. The record covers no other grid: the main grid
`probability_accuracy_grid.csv` and the study grids are outside it. When SciPy is installed, `test_cross_check_scipy.py` checks every
mapping against hand-computed and mpmath values, so a wrong parameterisation
fails. CI has no SciPy and runs only its classification fixture, which shows
that an unmapped function is reported as not mapped, never as beyond range.
The survival references are computed on the upper tail directly (never `1 - CDF`),
so they stay accurate in the deep tail where a naive subtraction would collapse.

### Survival reference basis

Every committed survival reference must equal the survival function at the
binary64 Double its argument tokens parse to, which is what the exporter passes,
to at least 25 significant digits (relative difference at most `1E-24`, the
`SER_FLOOR` of `audit_references.py`). `check_survival_references.py` enforces
this on every survival row of the main grid, the independent holdout and every
study grid that carries one (#159):

- It recomputes each value by a direct upper-tail route at the exact Double
  (`_survival_oracle.py`): `I_{1-x}(b, a)` for Beta and F, the upper incomplete
  gamma for Gamma and Chi-square, `erfc` for the normal family, closed forms and
  exact sums elsewhere, with quadrature where both shapes are huge. It never uses
  `mpmath.betainc(a, b, x, 1)`, which is itself a complement.
- Each value is computed at two working precisions that rise with the shape
  magnitude, and must agree with itself to 32 digits before it can judge a
  reference.
- A CSV under `benchmark/` that carries survival rows but is not listed in the
  checker fails the check, so a new grid cannot escape it.
- Rows that miss the floor but cannot be re-referenced yet are registered in
  `survival_reference_debt.json`, one exact cell per entry, with the blocker
  that prevents the change. A registered cell that is fixed or removed makes its
  entry stale, and the check fails until the entry is deleted.

The registry is empty since the #81 wave (2026-11-06). It held 73 references:
49 in the main grid, 15 in the holdout, and 9 that the #72 normal-tail record
reads (7 in `holdout_older`, 2 in `survival_boundary`).
`rebase_survival_debt.py` re-based all 73 from the source that owns each row:
- the `beta_f_unbalanced` and `cdf_large_shape` study grids for their
  main-grid copies;
- `generate_reference_values.py`, `holdout/generate_holdout.py` and the two
  study generators for the rest.

Each new value had to meet the floor against the oracle. Only the reference
cell changed, and the old and new values are in `study_reference_rebase.csv`.
The main grid and the holdout were then re-exported from Excel to rebind them.

The check runs every row, without sampling, in about 100 CPU seconds, in
parallel by default. It is wired into `test_evidence_tools.py` with its negative
fixtures, `test_survival_references.py`; one of them stores a reference at the
decimal token instead of the binary64 argument and requires the check to fail.

Study generators re-reference their committed grids with
`--merge-references` (report) and `--merge-references --write` (patch). The
merge changes reference cells only, refuses when the generator and the grid
disagree on any other column or on the row set, and records every change in
`study_reference_rebase.csv`.
