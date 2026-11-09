# Two-Double lower P/Q feasibility preregistration

The prototype passed all 269 previously revealed A<=1 points in exploratory
work. Those points are fitting evidence only. No VBA candidate exists.

Freeze `precision_probe.py` and the 220 disjoint binary64 coordinates in
`precision_protocol.json` before evaluating any of them. The boundary and upper
B6 branch stay unchanged. Lower logP, series, log1p and expm1 use bounded
two-Double arithmetic; existing rounded logGamma1p is retained. No high-precision
arithmetic or FMA is used by the prototype numerical path.

Require all pointwise P/Q errors <=3E-15 /5E-15, all adjacent pairs monotone,
finite in-range probabilities, oracle stability at 90/120 digits and independent
quadrature agreement <=1E-80. Missing or unconverged values reject; all failures
are retained. Preserve exact original-seam diagnosis and all source hashes.

The exploratory domain is explicitly restricted in JSON. A successful screen
is feasibility only, not full-domain acceptance, Excel validation, export
provenance or permission to merge. A failure consumes the set; do not retune
against it and still call it independent.
