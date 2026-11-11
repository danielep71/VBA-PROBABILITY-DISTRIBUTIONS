# #307: transition diagnosis and two development candidates

This work uses only revealed phase-A fit shapes. It neither evaluates the
117-point holdout nor changes the original rejected routing rule, production
VBA, contracts, exported measurements or regression expected values.

## Diagnosis of B1

Independent direct P and Q references at exact binary64 inputs agree at
80/120 decimal digits to better than 1e-65 relative. P errors, measured in
units of its output ulp, are:

| Shape | Series side | Complemented CF side |
| --- | ---: | ---: |
| 1e-8 | +0.974728 | -0.025272 |
| 1e-4 | +1.062178 | +0.062105 |
| 0.1 | +2.411257 | +0.347030 |
| 10 | +1.304077 | -1.432331 |

The drop is a difference between independently rounded component answers.
Pointwise contract compliance does not imply monotonic composition. Shape 10
also has a material opposite-signed error on the CF side; the diagnosis does
not attribute every error exclusively to the lower series.

## B2: double-double transition band — REJECT

Use the old public double-double route where abs(X-(A+1)) <= sqrt(A), otherwise
B1, retaining A2=10000. This is an exploratory development choice, not a
preregistered acceptance result. Across 26 revealed shapes and three triples
per shape (lower band edge, original seam, upper band edge), four new CDF
failures occur at the lower edge: A=1e-8, 1e-5, 0.001 and 0.1. The baseline
is monotone at these pairs. The band moves the problem rather than solves it.

## B3: directly computed smaller tail only — REJECT

Return the fast route and its complement only when its directly computed tail
is in [0,0.5]; otherwise use the old DD route. Retain A2=10000.
The median decision is located mechanically by integer-bit bisection in the
lower region, at revealed fit shapes only. No bracket is synthesized where
the finite-positive domain contains no crossing.

The initial 48 predecessor/boundary/successor triples pass. This is retained,
not promoted to acceptance. Expanding each same boundary to 16 representable
values on either side (1,584 evaluations) reveals 13 decreasing-P/increasing-Q
pairs across seven shapes: 0.25, 0.5, 0.75, 1, 1.5, 5 and 50.

One pair switches DD to series. Twelve fail without a routing switch: six
within the series and six within the continued fraction. The old DD model is
monotone at all 13 pairs. Routing adjustments alone cannot repair these
observed within-routine errors wherever those routines remain exposed.

## Reproduction and evidence boundary

- `band_development.py` writes `band_development.json`: B1 independent-reference
  diagnosis and B2 exact inputs, output bits and baseline values; exits 1.
- `smaller_tail_development.py` writes `smaller_tail_development.json`: both B3
  narrow/wide outcomes, exact failure pairs, baseline outputs, deterministic
  hashes of all evaluated rows and input/script hashes; exits 1.
- `test_b64_fit.py` replays both reports, verifies the retained rejection and
  checks baseline monotonicity at all B3 counterexamples. A passing fixture
  means the rejection reproduces, not that either candidate is accepted.

These are source-derived Python/VBA-model observations, not new Excel
measurements of the proposed public adapters. Phase-A host fidelity does not
certify the newly examined neighboring inputs.

## Next decision

A2 remains 10000 and the holdout remains untouched. Develop and measure a
separately identified improvement to the fast kernel's arithmetic, including
series accumulation/final multiplication and CF/prefactor composition, on
these revealed development pairs. Preserve both rejected candidates. A new
fit/fidelity wave is needed before freezing another public routing candidate;
changing thresholds, patching individual ulps or claiming these checks as an
untouched holdout is not an acceptable repair. #307 remains open.
