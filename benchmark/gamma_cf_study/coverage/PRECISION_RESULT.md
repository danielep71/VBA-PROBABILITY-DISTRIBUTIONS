# Two-Double lower-tail feasibility result

**REJECT overall. No production candidate or merge acceptance.**

Frozen design `0b7c4c9c0087ca4834e15624c3077adde118b53a` was published before
any of its 220 points were evaluated. All 220 points have stable 90/120-digit
references and independent quadrature agreement within 1E-80. All pointwise
P/Q errors pass the frozen 3E-15 / 5E-15 thresholds. Maximum relative errors
are 1.569E-16 (P) and 1.698E-15 (Q). The original A=.5 seam is repaired in the
model. This is a Python two-Double prototype on a bounded domain, not VBA.

The overall screen fails one adjacent monotonicity assertion:

- A=.39 (`0x1.8f5c28f5c28f6p-2`), Q/SF, rows 148 -> 149.
- X: `0x1.63d70a3d70a41p+0` -> `0x1.63d70a3d70a42p+0`.
- Q: `0x1.1897def153c27p-4` -> `0x1.1897def153c28p-4` (+1 ulp).
- The computed branch boundary is `0x1.63d70a3d70a3ep+0`; both points
  are above it and execute the unchanged B6 CF. The baseline B6 model gives
  the identical values. This is a newly revealed defect in that model path,
  not a new regression caused by the two-Double lower formula.

All evaluated lower-path pairs and crossings pass this screen, but the frozen
whole-function monotonicity rule is not waived. The 220 points are now revealed
and cannot be reused as independent evidence for a revised candidate.

`screen_precision.py` retains every point, its references, diagnostics and input
hashes in `precision_screen.json`, and exits 1 on rejection. The original
269-point exploratory fitting result is not counted as independent evidence.

## Primitive fixtures

`python benchmark/gamma_cf_study/coverage/test_precision_probe.py` currently
reports **2 test methods passing, 1 failing**. The failing subcase requires the
exact identity log(1)=0; the bounded atanh series and rounded ln(2) leave
+3.0814879110195774E-33. This is a known prototype defect, not silently exempted.
Nonzero logarithm, log1p and expm1 cases meet the fixture's 2^-96 relative bound;
product/division residual and invalid-range fixtures pass. No change was made
to the frozen prototype after discovering this failure. Its log(1) defect must
be corrected in a separately identified revision before production use.

## Real-Excel follow-up

Two permanent assertions for the revealed A=.39 CDF/SF adjacent pair are added
by `install_precision_counterexample.py`, with hi;lo diagnostics. The suite
count becomes 1947. These assertions validate the existing B6 VBA path; they
do not port the two-Double prototype or change a numerical source file.
[Excel run 42470937344](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/42470937344)
ran merge checkout `e5a65792b49c8f8e3af0e2159b96990ebcd26a5d` for head `5fbbbe6`,
Excel 16.0 build 20430, 64-bit: **1947 total / 1906 PASS / 41 FAIL**.
The CDF assertion passes; the SF assertion fails, and its two hi;lo outputs
reconstruct exactly the model's `...3c27` -> `...3c28` binary64 values above.
The previous 40 failures remain. Complete failure excerpt:
`precision_excel_42470937344.txt`. No export occurred.

## Post-hoc upper-path decomposition

`upper_rounding_diagnosis.py/json` retains the two-point source-bound experiment.
Both Lentz runs stop at depth 62; both backward values round to
`0x1.19f0d60282c4ep-1`. The prefactor incorrectly rises from
`0x1.fd8d8c7d60b32p-4` to `0x1.fd8d8c7d60b33p-4`.
Replacing only the prefactor with its high-precision value makes the pair
monotone; replacing only the fraction also removes the rise in this pair.
This isolates rounding interactions for this counterexample, not a universal
root-cause or fix claim. The exact finite depth-62 fraction is still distinct
from the converged oracle; it is not silently certified as exact Q.

Next work must address upper-CF monotonicity as well as the lower-path precision
strategy, correct the exact-zero primitive, and cover the public domain outside
this deliberately restricted prototype. No contract, seam or cap is relaxed.
#24 and #26 stay open; #241 and #242 remain draft and blocked.

CI follow-up: the v1 fixture now asserts the known positive residual +2^-108
as a negative regression case, while v2 separately requires exact log(1)=0.
Both fixture modules are registered in benchmark/test_evidence_tools.py.
The earlier written minus sign was incorrect; the measured magnitude and
rejection were unchanged. The original frozen v1 numerical code is untouched.
