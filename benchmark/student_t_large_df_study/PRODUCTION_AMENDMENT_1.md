# Production amendment 1 — outside-df control tolerance (#34)

`PRODUCTION_PREREGISTRATION.md` stays as frozen. This amendment supersedes
exactly one number in its §7, and nothing else.

## What failed

The first Excel run of the implementation (run 42245283526, commit
`8a64b3d`) passed 949 of 950 assertions. The one failure was the control just
below the df seam:

    t sf(2, 1E+05 - 1 ulp) on the continued fraction
    got 2.27514817286814E-02, expected 2.27514817287532E-02
    rel err 3.16E-12, tol 2E-12

§7 set that control at 2E-12 from §4's seam figure, 1.25E-12 at `t = 2`. That
figure came from the Python model of the continued fraction
(`validate_expansion.cf_double_tail`), not from Excel. Excel's continued
fraction is less accurate there. The committed `tchi_large_df` baseline,
exported from the base commit (`5273b0c`) before any route existed, already
shows **3.16E-12** at `t = 2`, `df = 1E5`, the same as this run.

The control checks code the change does not touch. The large-df route itself
passed: every inside-region assertion held at 2E-15, including both seam
points.

## What was not seen

The regression failed, so the grid-export step was skipped. No PT1–PT4 row and
no holdout row has been observed. The rows of §6 are still unexported.

## The change

| §7 item | Was | Now |
| --- | --- | --- |
| Outside control below `df = 1E5`, `\|t\| ≤ 8` | 2E-12 relative | **5E-12** relative |

5E-12 is the committed Excel characterization, 3.16E-12, with room for the
last bits. It is not a route tolerance, and it only shows that the continued
fraction still serves the point, as §7 intends.

Unchanged:
- the route, its arithmetic and its coefficients (§1–§3);
- the model measurement (§4);
- the predictions PT1–PT4 and their bounds (§5);
- the holdout (§6);
- the contracts (§7);
- the inside-region tolerance of 2E-15 and the 1E-9 control above `|t| = 8`;
- the failure conditions (§8).

§4's A3 df-seam figure, 1.25E-12, is the model's figure. Excel's is 3.16E-12.
The design's conclusions do not depend on it: below `df = 1E5` the continued
fraction is retained either way.
