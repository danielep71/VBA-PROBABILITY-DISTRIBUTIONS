# Prospective pair-CF screen — 2026-11-12

Freeze this document in Git before evaluating the candidate. This is a new,
limited rejection screen informed by the revealed residual diagnosis at
4763425; neither earlier protocol nor its rejection results is amended.

## Candidate P1

Use F-A orientation. Retain the baseline inverse seed, density, bracket,
termination, Loader prefactor and public F transformation. Replace only the
forward CF with the existing DD operation graph from
`gamma_cf_study/coverage/precision_probe.py`. Its operations use two binary64
components, without arbitrary precision in the candidate numerical path.

Form the CF argument directly from x on the direct route; on the reflected
route form DD(1)-DD(x), never DD(float(1-x)). Port the existing modified-Lentz
coefficient order, with shape sums and recurrence in DD. Keep the production
iteration cap and 3e-14 convergence criterion; use the same positive FPMIN
clamp for near-zero denominators. Round the returned CF once, then retain
the baseline binary64 product/complement. Reject nonfinite intermediates.
Study domain only: retained shapes, positive finite intermediates bounded
by 1e100; no public-domain or subnormal arithmetic claim.

## Evaluation and predictions

Evaluate only the 8 affected and 5 clean retained identities, already revealed
by both earlier studies. No new fitting/invariance/holdout observation is read.
References: recompute at 50 and 80 dps, requiring 40 relative digits agreement.
Record exact inputs, pair, result, error, iteration count and runtime.

Prediction: preserving the CF complement and adding recurrence precision will
bring all affected errors below 2^-48, retain every control below 2^-48 and
within four times its baseline error. The baseline is the source-bound F-0
record in baseline_result.json, not the rejected F-A result.

This screen passes only if every point satisfies those bounds, references are
stable and every numerical path converges without an unsupported-range error.
A single failure rejects this exact candidate; record it without retuning.
Do not run later sets after rejection. A pass is permission to design the next
prospective phase, not to adopt or change VBA: broader fitting, branch seams,
Beta/Student-t invariance, independent holdout and real Excel remain mandatory.
No performance adoption claim is made by a retained-row screen.
