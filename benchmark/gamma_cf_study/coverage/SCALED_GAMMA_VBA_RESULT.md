# Private scaled Gamma VBA: exact-source Excel result

2026-11-08; #24/#26; draft PR #242 remains open.

The bounded private Gamma port now passes **all 475 added assertions** in real
Excel. The full suite has **8,190 assertions / 8,149 PASS / 41 FAIL**. All 41
public failure details match the previous scaled-DD run exactly. This is a
successful private model-fidelity checkpoint; the public regression is red.

## Source and artifact identity

- Candidate: `f0e8ef0e70f77e161f193be969b8ec3625580e10`.
- [Excel run 42512900145](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/42512900145);
  job 128829612986; artifact 13629714744.
- Artifact ZIP SHA-256:
  `b83fe11337dd961a73ec2b4c240625dc87c09ec008f111b9df061bb87efd91d4`.
- Excel 16.0 build 20430, 64-bit; import, execution-backed compilation and
  cleanup PASS. Regression FAIL because the retained 41 public failures remain.
- No grid export. No source or manifest rebinding. Original certification and
  log bytes are preserved in `excel_scaled_gamma_artifact.json`.
- The evidence fixture verifies all seven source blobs against the tested
  commit, the original log digest, assertion totals and complete failure
  details. The current imported sources still match that commit exactly.

## Verified scope

| Private checks | Assertions | Result |
| --- | ---: | --- |
| 129 revealed input pairs: success, P bits, Q bits | 387 | PASS |
| Invalid inputs and cleared outputs | 8 | PASS |
| Eight three-point routing seams: both tails, both intervals | 32 | PASS |
| Independent exact-dyadic normalization boundary cases | 48 | PASS |
| Total added checks | 475 | PASS |

The domain of the private implementation is finite positive binary64 a<=1 and
x<=32, including subnormals. The finite development set contains 69 in-domain
B6 points, 36 original points and 24 routing neighbors, with intentional
repetitions. It does not exhaust that domain. P/Q references in this host test
are Python-model bit patterns. Separate earlier mpmath comparisons support the
revealed development screen; they are not fresh independent acceptance data.

## Diagnosis retained

Four prior runs on `da14d50`, `1a3854b`, `974c661` and `6754fe7` each produced
8,142 / 8,003 PASS / 139 FAIL: 98 private failures plus the same 41 public
failures. Their verbatim artifacts remain in the four adjacent attempt JSONs.
The final bit diagnostic showed a noncanonical pair with a low term just
above half an ULP; repeating the host sum did not change its high term.

Private SGD canonicalization repairs this using an exact adjacent binary64
step and residual correction. It handles ties-to-even, signs and binade
transitions without relaxing the residual bound. The 48 independent Fraction
references exercise these boundaries. The frozen SD, SP and PV2 code and
public source are unchanged. The observations do not establish the processor
cause or a general accumulated-error theorem.

## Limits and remaining work

No public Gamma route was changed. The 41 public failures, caller/inverse
integration, the wider domain, accumulated-error control and fixed-candidate
fresh independent validation remain outstanding. No release acceptance,
merge, PR closure or issue closure is implied.

Local verification: all 72 numerical-tooling commands PASS in one complete
run on the corrected port; the six generation/source/evidence fixture methods
PASS after retaining the artifact. Strict static checks pass 18/18. These
local checks do not turn the public accuracy gate or global Excel result green.
