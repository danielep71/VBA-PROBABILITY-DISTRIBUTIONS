# B6 boundary and public-caller result: blocked

Excel 16.0 build 20430, 64-bit, executed the same 1881 assertions twice:

| Kernel | Head / synthetic merge | Run | PASS | FAIL |
| --- | --- | --- | --- | --- |
| Pre-B6 | d5d0bb9 / 1149059 | 42467939162 | 1706 | 175 |
| B6 | bb8c0d6 / ba15070 | 42467893663 | 1856 | 25 |

The two head trees differ only in `src/M_STATS_PROBDIST_SPECIALFUNCS.bas`.
The control uses the exact release f2ff901 blob 0d615b1906abd9c2394e0f01c4db9d34aa297368.
Tests, policy (1881), inputs and thresholds are identical. PR #243 is diagnostic
only and must not merge. Neither run exported observations or rebound manifests.

`python benchmark/gamma_cf_study/coverage/compare_excel.py` regenerates
`paired_excel_result.json` from retained, timestamped job-log excerpts. All 175
and 25 failure lines are retained. There are 150 resolved failure labels, no new
failure labels and no changed diagnostic values among the 25 common failures.
The comparison uses printed Excel diagnostics, not a claim of bitwise equality.

## Remaining failures

| Surface / coordinates | Assertion count | Measured disposition |
| --- | --- | --- |
| P and lower-region Q immediately below X=A+1, A=1e-100, 1e-6, 1e-4, .01 | 17 | Present before B6; lower-series/complement route untouched. Q amplification belongs with #24; P accuracy must also be retained. |
| Gamma Q, A=100, X=202 and 808, scales .5/1/2 | 6 | Present before B6; A>1 route untouched. Root cause not isolated by this paired run. |
| Gamma inverse A=.01, p=.1 and .5 | 2 | Present before B6. Root cause not isolated by this paired run. |

These remain failing assertions with their original thresholds. They are not
waivers or a new known-limitation contract. Maximum displayed normalized errors
are 5.21245878151806E86 in tiny-shape Q, 7.38298311375729E-14 at A=100, and
1.39521727504643E-12 in the inverse cases. The huge Q figure is a relative,
normalized error; it is not an out-of-range raw probability.

All 24 fresh public holdout pairs (indices 65–88, nine assertions each) pass
with B6. All 40 underflow probes, 18 Poisson forward checks, 18 Poisson inverse
inequalities, eight shape-seam monotonicity checks and seven invalid/status
checks pass. No claim is made about untested inputs. The holdout is now revealed
and cannot be reused as untouched evidence for a retuned candidate.

The initial run 42467568896 had 28 failures. Three were a fixture defect: the
Poisson inverse lower inequality evaluated CDF(N-1) at N=0 despite the public
API's count input policy. Commit bb8c0d6 handles the mathematical lower support
boundary directly, retaining one assertion per case and total 1881. Both paired
runs use that corrected suite.

## Outstanding acceptance

- Resolve or separately govern the measured pre-existing defects without
  relaxing thresholds or silently dropping tests.
- Complete explicit iteration-cap/non-convergence and broader X-seam proofs.
- Promote representative rows to the permanent main grid and obtain fresh
  exact-source exports once the complete regression passes.
- Obtain current CI and review before integration. #26 remains open; this
  expanded suite is not production acceptance.
