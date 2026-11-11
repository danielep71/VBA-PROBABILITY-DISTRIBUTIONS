# Recurrence reuse: confirmation below target (#307)

Do not promote upper-only recurrence reuse to production. The independent confirmation failed the frozen requirement of at least 5% lower median time on BOTH P and Q at shape=10000, x=10100. Do not pool runs or switch candidates using confirmation results.

Both complete runs passed all 216 samples, bit identity and 1008 evaluations. ZIP digests were verified before extraction and summarize.py passed locally.

| Run | Upper P speedup | Upper Q speedup | Decision |
|---|---:|---:|---|
| 43043166533 attempt 2 | 1.094x | 1.097x | Screening target met |
| 43048952341 | 1.125x | 1.034x | Confirmation target NOT met |

Confirmation reductions are approximately 11.1% for P and 3.3% for Q. A green workflow certifies measurement completeness and bit identity, not performance acceptance. Timing controls and other cases also vary. Raw results retain every case.

Screening source: e2951b6a2c6e4714f86399f0f21a6c6361a62839; artifact 13868695234 SHA-256: 01ba65cd303d2d58980330280de16a41d3b5ad366df8beaa3c1e3d2fc082538c.

Confirmation source: 7f5fbeb488113e7c62be8f4371afdcc641e287a4; artifact 13868318641 SHA-256: 5ef0b0453a77c6529ca123c202fe64df3b7347da5ab47ee1cf1987f5a49e8967.

The [confirmation plan](CONFIRMATION.md) was committed before that run. Both complete reports are retained in results/.

Earlier runs 43030484382 (193/216, RPC failure), 43036251207 (0/216, COM activation failure), and 43043166533 attempt 1 (COM activation failure) remain historical failed attempts. Their samples were not pooled with complete runs. Earlier retained reports and HOST_RECOVERY.md document the host failures; the later complete runs supersede that blocker. Root cause remains undetermined.

Production source, routing, thresholds and accuracy manifests are unchanged. The untouched routing holdout was not read or evaluated. #307 remains open for its remaining work; this recurrence optimization is not accepted.
