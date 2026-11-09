# Extended scaled Gamma: exact-source Excel result

2026-11-08; #24/#26; PR #242 remains open and draft.

**All 1,772 added private assertions PASS.** Full Excel suite: **9,962 total /
9,921 PASS / 41 FAIL**. All 41 previous public failure details match exactly.
The public source has not changed; this is an extended private-candidate
checkpoint, not public integration or release acceptance.

## Exact identity

- Candidate: `924ff800077a3001bd1bcc246ad3a370f1c78fba`.
- [Excel run 42515382823](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/42515382823).
- Job: 128837153196; artifact: 13631475342.
- ZIP SHA-256:
  `9f6facd9a70831c3b2e4230664da513fd352156e0260f853c4fef86a22b9fc2d`.
- Excel 16.0 build 20430, 64-bit. Import, execution-backed compilation and
  cleanup PASS. Full regression remains FAIL because of the 41 public failures.
- No exports. Original certification/log bytes are retained losslessly in
  `excel_extended_gamma_artifact.json` with their original SHA-256 digests.

The evidence fixture checks all seven imported source blobs against the exact
candidate, original log bytes, policy count, no-export flags and every previous
failure detail. Old Excel records are not rebound to this candidate.

## What the host verifies

| Added checks | Assertions | Result |
| --- | ---: | --- |
| 587 distinct pairs: success, P bits and Q bits | 1,761 | PASS |
| Invalid inputs and cleared outputs | 8 | PASS |
| Rejected bounded-switch control and repaired P/Q shape seam | 3 | PASS |
| Total | 1,772 | PASS |

The rejected simple-switch counterexample at x=predecessor(2), a=1 to
successor(1), is reproduced using the actual public fallback. The extended
private pair removes the wrong-direction step. The shape-two subnormal midpoint
case also matches the independent rationally justified result, P=0 and Q=1.

The port spans finite positive a<=128 and x<=1024, including subnormals and
the new positive lower series. Python model fidelity is verified on prior
revealed cases, all 220 newly revealed validation coordinates, routing probes,
domain edges and shape-two rational cases. It is finite coverage, not a theorem
for the entire domain. See [the model review and proof](EXTENDED_GAMMA_REVIEW.md).

## Local verification and limits

All 18 strict static checks pass on the port. The full 73-command numerical
tooling suite passes in one complete run. Eight generation, source, model,
disjointness, rational and artifact test methods pass after retaining evidence.
The first local static attempt identified the new test as untracked; staging
the already-listed file corrected the inventory, and the complete rerun passed.
No guard or expectation was removed to bypass that check.

The v1 220-coordinate screen is retained unchanged, including its limits. Its
additional halfway counterexample remains reproducible. V2's change has an
exact rational argument for the shape-two subnormal region; the old 220 points
are revealed replay evidence for v2, not a new untouched holdout.

Remaining gates: public routing and its new outer boundaries, larger-shape
handling, accumulated-error control, inverse/ratio composition, fresh validation
of the public candidate, current exports and final CI/review. #24/#26 and both
draft PRs remain open; the global regression result remains red.
