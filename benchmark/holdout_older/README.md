# Older-families holdout

A retroactive, off-compliance-grid holdout for every family whose contracts
predate the independent-holdout discipline (F-02). The points were never used to
calibrate a threshold.

| File | Role |
|---|---|
| `gen_holdout_older.py` | Builds the grid's references with mpmath at 40 digits; `--merge-references` re-references the committed grid. |
| `holdout_older_grid.csv` | The grid; `observed_vba` filled by the macro. |
| `M_STATS_PROBDIST_HOLDOUTOLDER.bas` | Export macro. |
| `analyze_holdout_older.py` | Writes `holdout_older_summary.md`, the per-contract holdout worst against its threshold. |
| `holdout_older_summary.md` | The recorded summary. |

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
