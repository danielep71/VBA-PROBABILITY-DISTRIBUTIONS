# Completing the direct upper reference

The first completion evaluated all 198 points, but retained four oracle errors
at indices 158, 171, 184 and 197 (x=2a). There were no recorded candidate
accuracy or monotonicity violations on the other 194 points. Its REJECT record
remains `wide_complete_screen.json`; it is not a completed numerical pass.

Before rerunning, freeze `screen_wide_complete_v2.py`. It uses the same candidate,
198 coordinates, thresholds, precision levels and comparison rules. If mpmath's
direct upper incomplete Gamma raises a convergence error, integrate the
normalized upper density directly after t=u/rate, rate=1-(a-1)/x. Factor out the
density at x to keep the quadrature integrand of order one even for a tiny tail.
This is a direct upper-tail integral, not subtraction from the lower result.
The fallback is restricted to a>128 and x>a; other oracle failures remain errors.
The direct lower quadrature is unchanged except for recognizing the alternative
ValueError convergence exception emitted by mpmath. Each row records its method.

The candidate is unchanged. This is completion of the same frozen screen, not
another independent holdout. Any remaining ordinary exception is recorded.
