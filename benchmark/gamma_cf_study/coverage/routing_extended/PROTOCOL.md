# Extended public-routing preflight

2026-11-08; #24/#26; PRs #241/#242 remain open and draft.

Freeze the proposed dispatch before evaluation: use extended Gamma v2 for
finite positive a<=128, x<=1024, otherwise retain the source-derived B6 path.
This is development boundary coverage, not a new independent holdout.

At each outer boundary use predecessor, boundary and successor binary64
coordinates. Test both axes and both tails. Require the existing 3e-15 P and
5e-15 Q normal-output screens; require reference bits for subnormal/zero
outputs. Direct mpmath lower and upper references at 160/320 digits must round
identically. A wrong-direction step is a failure without a tolerance allowance;
compare against B6 to distinguish newly introduced and retained failures.

Retain every row and any rejected design unchanged. Failed old-path values
remain failures: a passing interior candidate cannot certify a failing public
composition. Model observations must not be described as Excel results. Public
source, old artifacts and thresholds remain unchanged in this preflight.
