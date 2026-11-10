"""Separate revision: exact leading product in the shape-two subnormal region.

For a=2 and 0<x<=2^-512, e^-x*x^2/2 < P(2,x) < x^2/2.
The omitted correction is much smaller than the dyadic product's distance
to a rounding midpoint. At an exact midpoint, an odd square is 1 mod 8,
so nearest-even selects the lower neighbor, as does the true probability.
Independent rational enclosure fixtures verify the tested rounding decisions.
"""
import math
import extended_gamma as v1
from scaled_product import product, rounded


def public(a,x):
    if a==2. and math.isfinite(x) and 0<x<=2.**-512:
        h,l,e=product(x,x)
        return rounded(h,l,e-1),1.
    return v1.public(a,x)
