# Single-round reference conversion: preserve the initial public screen

The first frozen public screen retained one failure: a=95.0625, x=1024,
Q was reported as 0x0.f0c9cf0026f39p-1022 against a reference converted by
mpmath to 0x0.f0c9cf0026f3ap-1022. The 320-digit reference is approximately
4235992831782713.33053521835056825067 units of 2^-1074, so its nearest
binary64 is the former value. Exact rational conversion of the represented
scaled pair also gives the former value. The pair differs from the direct
reference by about 7.08e-30 relative, safely away from the half-unit boundary.

Keep the original REJECT record, source and cohort unchanged. Before completing
the screen freeze `reference_binary64.py` and `screen_public_fresh_complete.py`.
The completion converts the exact mpmath dyadic using integer quotient/remainder
and ties-to-even at the final binary64 quantum, without an intermediate rounded
normal significand. Very small references are classified before constructing
large powers of two. Both 160/320-digit references must still round identically.
The candidate, 206 coordinates, 24 inverse points and acceptance criteria do not
change. This is completion of the same screen, not a new independent cohort.

Independently test the converter against exact Fraction rounding, halfway and
normal/subnormal transitions, both signs, and extreme underflow. Retain a direct
scaled-density quadrature cross-check of the reported counterexample. This
oracle correction does not authorize widening the exact-subnormal criterion.
