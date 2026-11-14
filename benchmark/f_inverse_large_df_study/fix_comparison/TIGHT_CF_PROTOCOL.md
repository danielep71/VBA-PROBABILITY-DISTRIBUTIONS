# Prospective tight-CF screen — 2026-11-12

Freeze before evaluating the eight new points below. The P1 failure and the
subsequent tight-stop diagnostic on its 13 revealed rows informed this design.
Those rows are fitting evidence, never an independent success. Earlier frozen
protocols and rejections remain unchanged.

Candidate P2 is precisely P1 with the CF relative delta-stop changed from
3e-14 to 2^-100 in the study's two-Double arithmetic. No other algorithm,
iteration cap, prefactor, seed, Newton rule, pair unwind or F transformation
changes. No high-precision arithmetic in the candidate. Same study-only
range checks as P1. Compare against P1 and production-model F-0.

New screen: p in {0.2,0.8} crossed with (df1,df2) in
{(2e4,6),(6,2e4),(40,2e7),(2e7,40)}. These eight identities have not been
computed by this study. They are distinct from the 13 retained identities and
the original frozen fit/holdout Cartesian products. Record exact binary64
inputs and complements. Existing independent holdouts remain unopened.

Predict all eight P2 quantiles below 2^-48 relative error. Require 50/80-dps
reference stability to 40 relative digits, finite outputs, convergence within
the original caps, and a direct comparison with F-0/P1. On all 13 revealed
rows require P2 below 2^-48 and clean controls within four times F-0 error.
No performance threshold: record runtime and iteration counts without an
adoption/performance claim. One failed point rejects this exact screen;
no tuning after observing it.

A pass is limited to this screen. It does not authorize VBA or holdout access.
Before adoption: broader fit, balanced-huge comparator, geometry/CF seams,
Beta and Student-t invariance including small shapes, genuinely independent
holdout and exact-source Excel validation are still required. Any subsequent
candidate change requires a new protocol and new independent validation.
