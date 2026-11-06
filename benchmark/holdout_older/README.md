# Older-families holdout

A retroactive, off-compliance-grid holdout for every family whose contracts
predate the independent-holdout discipline (F-02). The points were never used to
calibrate a threshold.

| File | Role |
|---|---|
| `gen_holdout_older.py` | Builds the grid's references with mpmath at 40 digits; `--merge-references` re-references the committed grid. |
| `holdout_older_grid.csv` | The grid; `observed_vba` filled by the macro. |
| `M_STATS_PROBDIST_HOLDOUTOLDER.bas` | Export macro. |
| `analyze_holdout_older.py` | Writes `holdout_older_summary.md`: the per-contract holdout worst against its threshold, every failing row, and every row no contract scores, with its disposition. Exits non-zero while a contract fails. |
| `holdout_older_summary.md` | The recorded summary. |
| `check_holdout_older_summary.py` | Fails when the committed summary differs from a re-run on the committed grid and registry, or when a row is neither scored nor covered by a `characterization_only` contract. Runs in `../test_evidence_tools.py`. |
| `test_analyze_holdout_older.py` | Fixture proof that a stale, hand-edited or missing summary fails the check, and so does a row with no disposition. |

## Amendment, 2026-11-05 (#159)

`gen_holdout_older.py` took every reference at the decimal argument token
(`"1.35"` as an exact decimal) rather than at the Double the exporter passes,
and formed the Student t, Chi-square, F, Gamma and Beta survivals as `1 - CDF`.
26 of the 27 survival references kept only 15 to 19 of their 30 digits. The generator now
evaluates every reference at the binary64 argument and computes each survival
directly. It also stops emitting the two deep-tail survival rows removed from
the grid in `ecf3e45`, so it reproduces the committed row set.

The grid was re-referenced with
`python gen_holdout_older.py --merge-references --write`, which changes
reference cells only: 74 references moved, 19 of them survival, by at most
2.6E-15 relative, and no observation changed. The old values are in
`../study_reference_rebase.csv`. The merge holds back the 34 rows of the
functions the #72 normal-tail record reads, 7 of them survival; they are re-based
with the #81 wave, when the #72 kernel evidence is redone, and
`../survival_reference_debt.json` registers the 7 survival references until then.

The corrected references reveal a contract breach, tracked in #163. Against
the new reference, the measured error of `Beta_Survival(0.82, 2.7, 3.4)` is
5.45E-15, above the 5E-15 threshold of `Beta_Survival.balanced.output_rel`;
against the old one it read 4.34E-15. `holdout_older_summary.md` is not
regenerated here: #164 owns it, and it is already stale against the current
contract registry.

## Amendment, 2026-11-05 (#164)

`holdout_older_summary.md` is regenerated with `python analyze_holdout_older.py`
on the committed grid and registry. It records 59 PASS and 3 FAIL across 62
contracts. No grid cell, reference, observation, contract or threshold changed,
so the grid's SHA-256, which `../normal_tail_study/results.json` binds, is the
same.

### How the summary went stale

The summary was written in `ecf3e45` (2026-07-21). Re-running the analyzer
on that commit's grid and registry reproduces it byte for byte. Nothing ran the
analyzer after that, so three later changes did not reach it:

| Commit | Date | Change | Effect on a re-run |
|---|---|---|---|
| `43d2415` | 2026-08-06 | Retightened `Gamma_Cumulative.all.output` 2E-14 -> 3E-15 and `Gamma_Survival.all.output` 5E-14 -> 5E-15 from main-grid measurements after `416f877` | Both fail |
| `1fac483` | 2026-08-13 | Retired `LogGamma.all.output` | 63 -> 62 contracts; the 3 `LogGamma` rows are no longer scored |
| `3160ae9` | 2026-11-05 | Re-based 74 references to the binary64 argument (#159) | `Beta_Survival.balanced.output_rel` fails (#163) |

The observations themselves have not changed since `9ee6d7c` (2026-07-21). They
predate three source changes that move the Gamma rows: `17f9930` (2026-07-29)
moved the incomplete-gamma prefactor from `Exp(-X + A*Log(X) - PROB_LogGamma(A))`
to the Loader log-density, `416f877` (2026-08-06) stopped the ascending series at
machine epsilon, and `8c1506f` (2026-08-27, #23) replaced the off-grid
`PROB_StirlingError` that the Loader prefactor reads. The main grid was
re-exported for `416f877`; this grid was not.

### Failing rows

| Row | Regime and branch | Error vs 3E-15 or 5E-15 |
|---|---|---:|
| `Gamma_Cumulative(3.4, 2.7, 1.3)` | `all`; X = 2.615 < A + 1, series | 1.45E-14 |
| `Gamma_Cumulative(0.6, 2.7, 1.3)` | `all`; X = 0.462, series | 1.44E-14 |
| `Gamma_Survival(3.4, 2.7, 1.3)` | `all`; series, then `1 - P` | 1.85E-14 |
| `Gamma_Survival(9.0, 4.3, 0.8)` | `all`; X = 11.25 >= A + 1, continued fraction | 2.16E-14 |
| `Beta_Survival(0.82, 2.7, 3.4)` | `balanced`; min/max shape 0.79, A + B = 6.1 | 5.45E-15 |

I recomputed each reference with 60-digit mpmath at the binary64 arguments,
where X is the exact quotient x / scale, by two independent routes: for Gamma,
the lower series and quadrature of the density over the tail being measured;
for Beta, quadrature of the upper tail and `mpmath.betainc` over [x, 1]. The
two routes agree to better than 1E-57 and the
committed references to better than 1E-30, so the errors above do not come from
the references.

**Gamma.** The rows sit in the `all` regime, whose domain is "full tested
range". Nothing excludes shape 2.7 or 4.3 there. The error is the 2026-07-21
kernel's. A binary64 replay of the `9ee6d7c` source, with a Lanczos
`PROB_LogGamma` prefactor and a 3E-14 series stop, reproduces the six Gamma CDF
and survival observations to within 2.6E-16 relative, except 2.3E-15 on the
continued-fraction survival row. The cause is the absolute error of that `PROB_LogGamma`
in the prefactor: 1.37E-14 at 2.7 and 1.78E-14 at 4.3. With a correctly rounded
prefactor, the same 3E-14 stop errs by at most 5.5E-16. The same replay of today's source gives 4.7E-16 or
less on every Gamma row. That replay is a prediction, not a measurement: the
current kernel has not been exported at these shapes. The main grid's
`all`-regime Gamma rows, re-exported after `416f877`, measure at most 3.6E-16,
but their shapes (1.5, 2, 3) are integers or half-integers. Those take `PROB_StirlingError`'s
exact table, not the off-grid path that 2.7 and 4.3 take. The post-fix evidence
for that path is the open work of #23. `Gamma_Density(3.4, 2.7, 1.3)` carries the
same 1.43E-14 prefactor error and passes only because its contract is 2E-14.

**Beta.** `Beta_Survival(0.82, 2.7, 3.4)` belongs to `Beta_Survival.balanced`,
whose domain is min/max shape >= 0.1. #163 tracks it. For A + B < 1000 the current
source takes the same arithmetic as `9ee6d7c`, so unlike the Gamma rows this
observation still represents today's kernel. The literal factor
`A*Log(X) + B*Log(Y) - PROB_LogBeta(A, B)` and the beta continued fraction are
unchanged, and so is the Lanczos `PROB_LogGamma` for Z >= 0.5. In a binary64
replay, the Lanczos `PROB_LogBeta(3.4, 2.7)` has an absolute error of 4.4E-15.
With a correctly rounded `LogBeta`, the same continued fraction gives 2.6E-16.
The main grid's balanced `Beta_Survival` rows measure up to 4.31E-15 against the
same 5E-15.

### Staleness check

`check_holdout_older_summary.py` renders the summary with the analyzer's own
`render()` and fails, with a diff, when the committed file differs. It runs from
`../test_evidence_tools.py` in both the full and the `--checks-only` modes, so
the Accuracy Gate and `../refresh_evidence.py --check` run it. A bare
`refresh_evidence.py` regenerates the summary. The check reads only the CSVs and
needs no mpmath. It checks currency, not verdicts: the summary records the three
FAIL verdicts and still passes. A registry or grid change that is committed
without a re-run fails. `test_analyze_holdout_older.py` is the negative fixture.

## Amendment, 2026-11-05 (#171)

The three `LogGamma` rows move from regime `all` to regime `general`, the label
the main grid gives the same domain. I changed only those three `regime` cells:
no argument, reference or observation changed. `gen_holdout_older.py` emits the
same label, so `--merge-references` still matches every row and changes no
reference. The summary is regenerated: 60 PASS and 3 FAIL across 63 contracts,
and no row is left unscored.

### Why the rows were unscored

`1fac483` (2026-08-13) retired `LogGamma.all.output`, the global relative claim.
Three regime contracts replaced it, and the main grid labels its `LogGamma` rows
with their regimes:

| Contract | Metric | Threshold | Domain |
|---|---|---|---|
| `LogGamma.small_positive.log_abs` | absolute | 5E-13 | 0 < Z <= 0.25 |
| `LogGamma.near_zero.log_abs` | absolute | 5E-14 | 0.9 <= Z <= 1.1 and 1.9 <= Z <= 2.1 |
| `LogGamma.general.output_rel` | relative | 5E-13 | Z >= 0.5 outside those two neighbourhoods |

The reflection interval 0.25 < Z < 0.5 has no contract. This grid kept the
label `all`, which no contract has any more, so the analyzer dropped the rows.
Since #164 the summary listed them under "Rows no contract scores".

### Evidence for each row

I recomputed each reference at 70 digits at the binary64 argument by two
independent routes: `mpmath.loggamma`, and the recurrence to Z + 60 followed by
the Stirling series with exact Bernoulli numbers. The two routes agree to
better than 1.2E-58 relative, and the committed references agree with them to
better than 4.3E-31.

| Row | Binary64 argument | `PROB_LogGamma` branch | Contract domain | Relative error | Absolute error |
|---|---|---|---|---:|---:|
| `LogGamma(3.7)` | 3.700000000000000177636 | Lanczos, Z >= 0.5 | `general` | 1.30E-14 | 1.86E-14 |
| `LogGamma(12.3)` | 12.30000000000000071054 | Lanczos, Z >= 0.5 | `general` | 2.24E-15 | 4.08E-14 |
| `LogGamma(0.65)` | 0.6500000000000000222045 | Lanczos, Z >= 0.5 | `general` | 6.46E-15 | 2.10E-15 |

Against the 5E-13 of `LogGamma.general.output_rel`, the worst row has a margin
of 38.5x. None of the three arguments is in the main grid or in either grid of
`../loggamma_regimes_study`, so they did not set that threshold.

### The observations still represent today's `PROB_LogGamma`

The observations were exported on 2026-07-21 (`9ee6d7c`). Unlike the Gamma rows,
they still describe the current code:

- Since `9ee6d7c`, the only commit that changed `PROB_LogGamma` is `f6c7d9f`
  (2026-08-12). It added the small-positive branch for Z <= 0.25 and rewrote the
  accuracy comment. The reflection branch and the Lanczos branch, which all
  three rows take, are unchanged.
- The constants that branch reads have the same values: the nine
  `PROB_LANCZOS_P*` coefficients, `PROB_LANCZOS_G` and `PROB_HALF_LOG_TWO_PI`.
  `ae37be9` only joined `PROB_HALF_LOG_TWO_PI`'s split declaration onto one
  line.
- `17f9930` and `8c1506f` (#23) changed callers of `PROB_LogGamma`, the
  incomplete-gamma prefactor and `PROB_StirlingError`, not `PROB_LogGamma`
  itself. The exporter calls `PROB_LogGamma` directly for these rows.
- A binary64 replay of the Lanczos branch reproduces all three observations
  bit for bit. The replay evaluates each statement with 64-bit-significand
  intermediates and rounds each `Log` result and each assignment to binary64.
  The same replay reproduces all 27 main-grid observations of the Lanczos
  branch bit for bit, so it models the exporter's arithmetic and not only
  these three points.

### The disposition check

`check_holdout_older_summary.py` now also fails when the summary is current and
a row is neither scored by a contract with a threshold nor covered by a contract
with status `characterization_only`. That status is the registry's existing
marker for evidence that is measured but not claimed; the main-grid coverage
gate uses it too. The summary lists each unscored row with its disposition.
`test_analyze_holdout_older.py` adds two negative fixtures: a row with no
contract, the state `1fac483` left these rows in, and a row whose only contract
has no threshold and is not `characterization_only`.

The grid's SHA-256 changed, so I regenerated `../normal_tail_study/results.json`
with `validate_kernel.py --regenerate`. Only the bound hash, the timestamp, the
runtime and a new `regenerations` entry changed. Every recorded value is the
same, and `--verify` reports VERIFIED. The 7 survival references that
`../survival_reference_debt.json` holds for #159 are untouched.
