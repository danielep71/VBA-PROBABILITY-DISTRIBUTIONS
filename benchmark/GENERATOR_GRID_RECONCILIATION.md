# The reference generator and the main grid agree (#349)

`generate_reference_values.py` and `probability_accuracy_grid.csv` had drifted
apart. At `074b800`, before this work, 111 of the generator's 1,519 rows were
not in the grid as the generator wrote them: 27 would have been appended and
84 patched, 45 of those with a different reference.

Each group below was settled by deciding which side was right. Where a
reference was in question, both values were compared with a 90-digit
evaluation at the exact binary64 arguments the exporter passes to Excel.
The row-level table is `generator_grid_dispositions.csv`: 125 entries for the
111 rows, because 14 rows needed two corrections.

`check_generator_grid.py` now runs in CI and fails when a generator row is
absent from the grid or held differently.

## Corrected in the generator

The committed rows were right; no grid row changed for these.

| Rows | What differed | Cause | Correction |
| ---: | --- | --- | --- |
| 5 | `StudentT_InverseCumulative` at `p = 0.5`, regime `all` | The grid scores the median under `exact_median` since #22 | The generator emits `exact_median` for `p = 0.5` |
| 6 | Inverse normal and lognormal at `q = 1E-12`, regime `deep_tail` | The `deep_tail` contracts hold `q <= 1E-13`; the grid has these rows under `split_boundary` | The generator emits `split_boundary` above 1E-13 |
| 4 | `DiscreteUniform_InverseCumulative`, probability one ulp away | `(j + 0.37) / n` was formed at 50 digits and then rounded; the committed Double is the binary64 quotient | The design point is formed in binary64 |
| 12 | F `tiny_unbalanced`, `x` one ulp away | `df2 / df1` and its multiples by 10 were formed at 50 digits | The design points are formed in binary64 |
| 14 | Beta and F inverse `tiny_unbalanced_representable`, reference | The generator held quantiles solved at the decimal arguments, up to 8E-17 off in probability; the committed references are the roots at the Doubles, within 2E-25 | The constants are the committed values |

A design point only selects where the library is tested. Forming it as the
committed rows were keeps the test points where Excel observed them, instead
of adding neighbours one ulp away that would need a new observation and leave
the old rows without a generator.

## Corrected in the grid

The generator was right. The grid was changed through `promote_grid_rows.py`,
with the 45 reference changes acknowledged by count and reason. No observation
changed: an observation is the value Excel returned, and it does not depend on
the reference.

| Rows | What differed | Cause | Largest relative move |
| ---: | --- | --- | ---: |
| 39 | Claim text of six contracts | The committed claim predates a change of the contract threshold; the registry is the source of the claim | none: metadata |
| 24 | Chi-square and Student-t `envelope_domain` forward references | Digits of an earlier oracle, beyond double precision | 1.5E-22 |
| 7 | F `envelope_domain` forward references | Four were the value at the decimal `0.999` or `1.001`, which they match to 2E-22, not at the Double; three differ beyond double precision | 2.8E-14 |
| 6 | `split_boundary` references at `q = 1E-12` | Each was a binary64 result printed to 25 digits: it equals a Double to 3E-25 | 2.6E-16 |
| 8 | F `tiny_unbalanced` density and cumulative references | Each was a binary64 result printed to 25 digits: it equals a Double to 4E-26 | 2.1E-15 |

The six contracts whose claim text was stale: `Beta_Density.balanced`
(5E-15, now 1E-14), `F_InverseCumulative.validated` (5.9E-13, now 2E-10),
`Gamma_Cumulative.all` (2E-14, now 3E-15), `Gamma_Survival.all` (2E-14, now
5E-15), `NormalStandard_Survival.all` (2E-14, now 5E-15) and
`StirlingError.all` (3E-17 absolute, now 1E-13). The gate reads thresholds
from the registry, so no verdict was computed from the stale text.

### Effect on the measured errors

Worst relative error of the affected rows, by contract, against the old and
the new references. Every row stays inside its contract by a factor of nine
or more.

| Contract | Before | After | Threshold |
| --- | ---: | ---: | ---: |
| `ChiSquare_Cumulative.envelope_domain` | 3.27E-15 | 3.27E-15 | 2E-12 |
| `ChiSquare_Survival.envelope_domain` | 2.04E-15 | 2.04E-15 | 2E-12 |
| `F_Cumulative.envelope_domain` | 2.14E-12 | 2.16E-12 | 2E-11 |
| `F_Survival.envelope_domain` | 1.08E-13 | 1.08E-13 | 2E-11 |
| `F_Cumulative.tiny_unbalanced` | 4.65E-14 | 4.65E-14 | 5E-13 |
| `F_Density.tiny_unbalanced` | 1.95E-15 | 3.68E-15 | 5E-14 |
| `StudentT_Cumulative.envelope_domain` | 1.02E-16 | 1.02E-16 | 5E-15 |
| `StudentT_Survival.envelope_domain` | 9.26E-17 | 9.26E-17 | 5E-15 |
| `NormalStandard_InverseSurvival.split_boundary` | 5.21E-17 | 2.22E-17 | 1E-10 |
| `NormalStandard_InverseCumulative.split_boundary` | 5.21E-17 | 2.22E-17 | 1E-10 |
| `Normal_InverseSurvival.split_boundary` | 4.37E-17 | 2.18E-16 | 2E-10 |
| `Lognormal_InverseSurvival.split_boundary` | 3.89E-16 | 2.22E-16 | 1E-9 |

### No other reference is a Double in disguise

Fourteen of the replaced references were binary64 results written out to 25
digits. Every reference of the main grid, the holdout and the study grids was
then tested for the same signature: 20 or more digits that equal a Double to
1E-23. In the main grid 19 rows have it, and each is a value that is a Double
by nature: a probability equal to its own parameter, a symmetric point at 1/2,
or a result within 1E-23 of one. Twelve of them are generator rows and agree
with the 50-digit evaluation.

## What the check does not cover

696 grid rows are not emitted by the generator: 356 belong to study evidence
sets and 340 are main-grid rows promoted from study folders.
`grid_reconciliation.csv` lists them. They are kept; their references are
owned by the studies that produced them.

The count in the issue, 106 rows, was taken at `de28cf0`. The five Student-t
medians were added to the difference by #22, which moved them to
`exact_median`.
