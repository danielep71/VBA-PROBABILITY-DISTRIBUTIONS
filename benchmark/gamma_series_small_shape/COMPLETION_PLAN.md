# Remaining #23 evidence

Prepared on 2026-11-07 against release/1.0.0 at f157e67.

## Cutoff rerun

Run the four unchanged prs_gamma_probe, prs_chisquare_probe,
prs_gamma_holdout and prs_chisquare_holdout exporters. No production source,
point, claim, or threshold changes. Preserve the pre-fix records in Git.
Run the existing analyzer in fitting mode on the two probe files and with
--claims on the two holdouts. The historical holdouts have already been
observed; this is the explicitly required post-#23 replay, not a new untouched
confirmation set. Never tune a cutoff to a holdout crossover. A rejected claim
requires a new design and independent confirmation before production wiring.

## Permanent main-grid coverage

Promote exactly the two original counterexample rows and all three scope
probes, selected by their frozen roles rather than their measured errors.
Use promote_representatives.py with benchmark/promote_grid_rows.py.
References are regenerated at the binary64 arguments; observations are blank.
No holdout row is promoted or relabelled as fitting evidence.
All five rows use the existing all-regime CDF contracts of 3E-15.
Export the main grid and global holdout from the candidate and commit grids,
bindings, certification and generated summaries together. The 27-row coverage
debt must remain unchanged. No threshold is relaxed.

## Closure

Require the cutoff comparison, five freshly exported main-grid rows, current
Excel regression and strict checks. Keep #26 open: the #231 wave measured
4.53E-14 residual survival error, above its unchanged 5E-15 contract.
Keep the original falsified B4 prediction and its explanation visible.
