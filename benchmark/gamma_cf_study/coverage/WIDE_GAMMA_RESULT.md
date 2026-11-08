# Wide Gamma model and private Excel port result

2026-11-08. #24/#26 remain open; #241/#242 remain open and draft.

## Completed model screen

Candidate and 198-coordinate protocol were frozen at
`69b849ac7873c40db1ca8f2962a011b8996d5557`. The 14 shapes are disjoint from
the explicitly audited 81-shape inventory. The original attempt stopped on
an oracle convergence error. The first completion retained four oracle
failures; neither attempt is a numerical pass. Separate direct-density
quadrature completions were frozen before evaluation at `830eaa7` and
`9cbb53b`; no candidate, coordinate or acceptance threshold changed.

The completed record `wide_complete_v2_screen.json` has **198/198 point
checks passing and zero monotonicity failures**. Maximum relative errors for
normal results are P 1.05958e-16 and Q 8.55024e-17. Subnormal/zero results use
exact reference bits. Both direct tails have stable rounded references at
160 and 320 digits; rows identify the actual oracle method used at each
precision. These points are now revealed, not an untouched future holdout.

The revealed outer-boundary replay has **93/93 point checks passing and zero
monotonicity failures**. It corrects the rejected extended-v2/B6 switch's
13 accuracy failures, one new and two retained wrong-direction steps.
Clarification of the frozen design narrative: at a=128 and its successor,
x=1024, the B6 comparison is *flat*, hence nondecreasing; it does not strictly
increase. Both B6 Q values are `0x1.a294dd60d47f8p-917`. The rejected switch
falls from `0x1.a294dd60d4810p-917` to that B6 value. Wide Q instead rises to
`0x1.a294dd60d49c4p-917`. Original records remain unchanged.

## Actual Excel result

- Port head: `b8efaf3c7810c48fadbfece38950ded190f79a56`.
- Tested PR merge: `d5b2a054676f643ddd1ae42099487cdeb50251c7`.
- Both trees: `dba5ae64a9f5854e5c20d1761ade562b4ad87eaf`.
- [Run 42521579897](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/42521579897), job 128856020382.
- Artifact 13632789471; ZIP SHA-256 `aeb7dd577c2abcff43804fccdc7bc269c012670dea1fda06b287dbff89c389d5`.
- Excel 16.0 build 20430, Windows X64, Office 64-bit.
- Actual regression: **10848 total / 10807 PASS / 41 FAIL**.
- **All 886 new assertions PASS**: 290 bitwise pairs, seven invalid inputs,
  four rejected/repaired routing controls, five natural-cap/recovery checks.
- All 41 earlier public failure details are identical to run 42515382823.
- Import, execution-backed compilation and cleanup PASS. Regression FAIL.
  No benchmark grid was exported.

The original artifact members are retained losslessly in
`excel_wide_gamma_artifact.json`. The fixture checks member/log hashes,
seven imported-source hashes and the 10848 assertion policy. It also retains
the raw synthetic merge commit object and verifies its Git object hash and
parent/tree identities. The identical reachable port tree provides the source
blobs for replay, without depending on an obsolete `refs/pull/242/merge` ref
or relabeling the artifact's actual tested SHA.

## Validation and remaining work

The port tree passes **18/18 strict local static checks and all 74 numerical
and tooling commands in one full run**. The new fixture separately regenerates
coefficients, verifies source freezes and shape disjointness, replays all
587 previous interior pairs, the 198-point screen and 93 boundary points,
and proves that removing only the WG_ block restores the prior test module.
Historical evidence is not rebound to the new source.

This is private-model/VBA fidelity, not public acceptance or a global error
bound. The production modules remain unchanged. The stricter geometric
remainder screen's resource boundary and actual Excel performance need to
be considered during public integration. The upper fraction's stopping rule
is still heuristic; the analytic Stirling remainder alone does not certify
all rounded arithmetic.

Next: integrate the completed numerical path into public Gamma P/Q, preserve
cap/diagnostic contracts, validate Gamma/Chi-square callers and inverses,
resolve all 41 public failures, freeze/pass fresh public validation and obtain
current exact-source exports plus final CI/review. No merge or closure is
justified by this private-port result.
