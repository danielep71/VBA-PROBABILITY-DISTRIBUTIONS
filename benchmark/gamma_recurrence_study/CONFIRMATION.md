# Recurrence reuse: completed screening and confirmation plan

## Completed run

Run [43043166533, attempt 2](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/43043166533/attempts/2), source e2951b6a2c6e4714f86399f0f21a6c6361a62839, passed completeness and bit-identity validation: 216 samples and 1008 evaluations. The retained JSON is results/run_43043166533_attempt2.json. Artifact 13868695234 ZIP SHA-256: 01ba65cd303d2d58980330280de16a41d3b5ad366df8beaa3c1e3d2fc082538c (verified before extraction).

At shape 10000, x=10100, upper-only speedups were 1.094x for P and 1.097x for Q (approximately 8.6% and 8.8% less elapsed time). Lower-only and combined did not meet the 5% reduction target for both P and Q at their affected primary points. Timing controls also varied, so this is screening evidence, not a production acceptance.

## Confirmation specified before the next run

Repeat the complete unchanged four-mode, three-round experiment. Keep all points, mode orders, warmups, call counts, generated numerical sources and validation criteria unchanged. Evaluate upper-only against baseline on the preregistered shape=10000, x=10100 P and Q cases: require identical bits everywhere and at least 5% lower median time on both primary cases in this independent complete run. Report all cases including regressions and controls. Do not pool runs to rescue a failed confirmation, select lower/combined from the confirmation, or splice partial runs.

If confirmed, upper-only remains a candidate for implementation and full numerical regression; this experiment alone does not authorize production promotion or closure of #307. The untouched routing holdout remains untouched.
