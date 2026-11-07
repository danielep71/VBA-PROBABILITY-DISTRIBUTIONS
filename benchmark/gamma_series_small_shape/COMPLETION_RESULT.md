# #23 remaining evidence completed

The initial wave and its falsified B4 prediction remain in `WAVE_RESULT.md`.
This follow-up completes the two outstanding acceptance items without changing
VBA source or relaxing any contract.

## Frozen positive-ratio crossover replay

Run [428](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/42421734701)
at `f157e67` exported all four unchanged arms with 955/955 regression PASS.
The fitting crossovers remain Gamma density 48 bits, Gamma CDF 40 bits,
Chi-square density 48 bits and Chi-square CDF 48 bits. All four frozen holdout
claims pass. `../positive_ratio_subnormal/POST23_RESULT.md` records the report,
CSV hashes, exposure history and repeatable checker. There is no survival
cutoff claim; production wiring remains #13.

## Five permanent main-grid rows

`promote_representatives.py` regenerates references only for the two original
counterexamples and all three scope probes, selected by their frozen roles.
No holdout row or study observation is copied. The promotion tool appended
five blank-observation rows; the existing study fixture now requires their
permanent presence and unchanged references, claims and evidence set.

Run [429](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/42422488108)
at `2d7a775c34874ce516e26cbe1c09e069d1637cfe` passed 955/955 assertions in
Excel 16.0 build 20430, 64-bit, and freshly exported the complete main grid
and global holdout in the same job. Applying that artifact binds both grids
truthfully and updates the retained certification record atomically.

| Function | Arguments | Relative error in main-grid encoding |
|---|---|---:|
| Gamma_Cumulative | 0.5, 0.6, 1 | 1.75E-16 |
| Gamma_Cumulative | 1, 5.3, 1 | 6.23E-16 |
| Gamma_Cumulative | 10.8, 10.3, 1 | 2.60E-16 |
| Gamma_Cumulative | 0.5, 0.0001, 1 | 1.18E-15 |
| ChiSquare_Cumulative | 1, 0.0002 | 1.18E-15 |

All five meet the frozen study bound of 3E-15 relative. Gamma all is
registered at 3E-15, while Chi-square all is broader (2.6E-10). The study
fixture independently scores all five actual main observations at 3E-15,
with exact, degraded 1E-11 and unavailable-observation controls; the row
claim alone is not relied on for enforcement. No registry threshold changes. These errors use the
main gate's hi/lo observation encoding. The 2088 previous main-grid observations
and all 659 global holdout rows are unchanged. Main-grid total is 2093, of
which 1737 belong to the main evidence set; coverage debt remains 27.

Export zip SHA256:
`c7277fa5f421cb47b21fe7c3bfec88c28f157e861bc08e4f94aa8eae99156f02`.

## Scope of closure

#23's prefactor repair, fitting/holdout evidence, permanent coverage and cutoff
revalidation are complete. #13 still owns production wiring; #24 owns the
lower-series complement defect; #26 still owns the measured 4.53E-14
continued-fraction survival breach. B4 remains falsified and explained, never
rewritten into a passing prediction. Release v1.0.0 is not certified by this
issue's completion.
