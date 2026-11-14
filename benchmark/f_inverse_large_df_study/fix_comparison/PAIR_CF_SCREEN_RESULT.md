# Pair-CF screen result — 2026-11-12

Protocol frozen at b9cc7166e4440f9d79be297df50a4a089512cb25 before execution.
Decision: REJECT this exact candidate; no fitting or holdout evaluation.

The candidate uses two binary64 components, including the exact difference
DD(1)-DD(x) for the reflected CF argument. No arbitrary precision is used
in its numerical path. References are separate and pass the 50/80-dps check.

| Kind | p | df1 | df2 | Relative error | Pass |
|---|---:|---:|---:|---:|---|
| affected | 0.5 | 1000000.0 | 3.0 | 2.9150e-15 | True |
| affected | 0.5 | 100000000.0 | 10.0 | 3.9118e-15 | False |
| affected | 0.9 | 2.5 | 1000000.0 | 5.7518e-15 | False |
| affected | 0.95 | 2.5 | 1000000.0 | 4.0570e-15 | False |
| affected | 0.99 | 2.5 | 1000000.0 | 7.4763e-16 | True |
| affected | 0.9 | 5.0 | 100000.0 | 1.3429e-15 | True |
| affected | 0.9 | 100.0 | 10000000000.0 | 1.8228e-16 | True |
| affected | 0.99 | 100.0 | 10000000000.0 | 3.2502e-16 | True |
| control | 0.9 | 1000000.0 | 3.0 | 1.8318e-15 | True |
| control | 0.99 | 1000000.0 | 3.0 | 1.1480e-17 | True |
| control | 0.5 | 2.5 | 1000000.0 | 3.4447e-15 | True |
| control | 0.9 | 100000000.0 | 10.0 | 1.0452e-15 | True |
| control | 0.5 | 100.0 | 10000000000.0 | 2.1657e-15 | True |

Five of eight affected rows and all five controls pass. The worst error is
5.75e-15 against the unchanged 2^-48 limit (about 3.55e-15). This is a major
reduction from the 1e-9..1e-11 errors, but not adoption evidence.

The candidate deliberately retains the baseline prefactor, inverse stop and
final transform; the remaining floor is not isolated by this result. A future
diagnostic should quantify their individual effects before proposing another
candidate. Do not silently tighten the CF stop, change the Newton rule, or
relax the threshold and reuse this protocol as if it predicted that change.

Four arithmetic/isolation tests pass; all three new study test files are now
registered in the per-PR evidence runner. The earlier CI failure was missing
registration, not a numerical result. The frozen older evidence is unchanged.
No Excel run or VBA modification was made for this screen.

Reproduce with:

```sh
python benchmark/f_inverse_large_df_study/fix_comparison/pair_cf_screen.py
python benchmark/f_inverse_large_df_study/fix_comparison/test_pair_cf_screen.py
```
