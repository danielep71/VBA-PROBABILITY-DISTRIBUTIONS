# Bounded public-routing preflight

2026-11-08; #24/#26; PR #242 stays open and draft.

Before modifying public routing, test the proposed dispatch: use the verified
private scaled Gamma pair for 0<a<=1, 0<x<=32, otherwise retain the current B6
path. This is a development preflight, not an independent holdout. Its purpose
is to reject a locally accurate implementation whose new switching surfaces
introduce discontinuities when composed with the existing public path.

Freeze this protocol, model, coordinates and source hashes before evaluation.
No source or coordinate retuning is allowed under this identity. Both models
and the full original point set are retained if this preflight rejects.

The tests straddle x=32 at fixed a and a=1 at fixed x with predecessor, exact
boundary and successor binary64 inputs. Compare both proposed and B6 outputs.
At increasing x, require nondecreasing P and nonincreasing Q; at increasing a,
require nonincreasing P and nondecreasing Q. No tolerance permits a jump in the
wrong direction. Identify existing failures separately from new failures.

Direct lower/upper mpmath references at 160 and 320 decimal digits must round
to identical binary64 values. The unchanged P/Q normal-output screens are
3e-15/5e-15 relative. At subnormal/zero references, require the correctly
rounded binary64 bit pattern; do not apply a relative-error denominator floor.
Oracle failure or invalid probability rejects rather than silently passing.

Retain every result and failure, including model failures. Source-derived B6
predictions are not Excel observations. Any counterexample used to block public
activation must be verified in Excel with the same proposed dispatch in the
private test module. This leaves the existing public source and 41 baseline
failures unchanged. After a successful preflight, a separately frozen public
candidate still needs new independent validation, full caller/inverse and
extreme-domain coverage, current accuracy evidence and review.
