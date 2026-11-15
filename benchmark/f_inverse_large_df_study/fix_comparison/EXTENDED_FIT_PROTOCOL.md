# P2 extended fitting gate — prospective protocol, 2026-11-12

Freeze this protocol and its runner before executing any row. P2 is unchanged
from `run_tight_cf.py`: P1 plus the 2^-100 DD continued-fraction stop. Earlier
screens remain immutable. Their success is not adoption evidence.

Evaluate all 24 original fitting identities: p in {0.05,0.25,0.75,0.999}
crossed with (df1,df2) in {(1e5,2),(1e7,4),(3,1e7),(10,1e9),
(1e3,1e3),(1e8,1e8)}. Run F-0 and P2, recording exact binary64 inputs,
complements, quantile/pair outputs, Newton and CF traces and timings.

Use the existing 230-step bisection oracle at 50 and 80 decimal digits,
requiring 40-relative-digit agreement for F and both pair members. Its CF
must explicitly report convergence: this runner supplies the same CF formula,
1e-45 stop and original 19999-iteration cap, but raises on exhaustion rather
than returning a partial value. This checking does not change candidate code.
Use that checked 80-dps forward for the F-quantile tail residual.

Apply the original fitting rules: each P2 relative F error must be <=2 times
F-0, and must remain below 2^-48 wherever F-0 is below it. For balanced-huge
(df1=df2=1e8), require <=2 times baseline also for each pair-member relative
error, tail residual and Newton count. Zero baseline error permits only zero
candidate error. Median Newton count across the 24 rows may rise by at most
2. Require finite outputs, successful original solver caps and stable oracles.
Continue recording all rows after a failure; never tune P2 during this run.

Any failing row rejects this candidate at this gate; unsupported baseline or
oracle exhaustion is an unresolved blocker, not a numerical pass. Preserve
both candidate failures and evidence blockers separately. No holdout, seam,
shared-path invariance or Excel runs are authorized by a failure here. A full
pass authorizes preregistration of seams and Beta/Student-t invariance only;
these and independent holdout and exact-source Excel remain adoption gates.
No VBA edits, release merge, threshold relaxation or claim of issue closure.
