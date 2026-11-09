# Scaled two-Double VBA result

2026-11-08; #24/#26 and draft PR #242.

Exact candidate: `86a365f5f7caa57ba0deff166549dfd357a244d5`.
[Manual Excel run 42494438289](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/42494438289),
job 128773565914; Excel 16.0 build 20430, 64-bit.

| Check | Result |
|---|---|
| Import, execution-backed compile, cleanup | PASS |
| Complete harness | 7,715 assertions / 7,674 PASS / 41 FAIL |
| New scaled-DD arithmetic | All 611 PASS |
| Prior suite | 7,104 assertions / 7,063 PASS / 41 FAIL, unchanged |
| Retained public failure details | All 41 match the preceding artifact exactly |
| Grid exports | None |

The workflow remains red because the public failures remain open. A successful
private arithmetic port is not a successful Gamma integration, independent
probability validation, or release certification. Public kernels are unchanged.

Artifact 13622408263 has ZIP SHA-256
`1950940ee8bb62330a02ab9fad59ef143e93dff21783ce62e708428b2d203144`.
The original log SHA-256 is
`7a083d7c22a5ecb1e81fc16664f7feb60af41ef4081cb4ed00778c1daeee1067`.
`excel_scaled_dd_artifact.json` retains the original member bytes as base64,
including BOM and CRLF; decode before verifying the recorded hashes.
`test_scaled_dd_vba_port.py` checks the seven source blobs against this exact
candidate, the log binding, harness counts, stages, absent exports and unchanged
failure details. The earlier certification records are not rebound.

The arithmetic reference and its empirical error screen are described in
[SCALED_DD_REVIEW.md](SCALED_DD_REVIEW.md); port scope and generation are in
[SCALED_DD_VBA_PROTOCOL.md](SCALED_DD_VBA_PROTOCOL.md). Remaining work is complete
Gamma P/Q integration, accumulated error control and routing, followed by a
fixed candidate, fresh disjoint validation and exact-source Excel acceptance.
