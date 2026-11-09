"""The frozen #24 holdout inputs (HOLDOUT_DESIGN.md). Do not edit after evaluation."""
import math

# Kernel shapes A: one per decade from 1E-7 to 1E-2, plus 0.465, chosen before
# any evaluation and checked new by check_disjoint.py.
SHAPES = (1.7e-7, 4.5e-6, 6.5e-5, 8.5e-4, 3.7e-3, 4.3e-2, 0.465)
# Standardized arguments below the branch seam X = A + 1.
INTERIOR = (0.05, 0.25, 0.6)


def seam(shape):
    """The binary64 predecessor of A + 1, A + 1 itself, and its successor."""
    s = shape + 1.0
    return (math.nextafter(s, 0.0), s, math.nextafter(s, math.inf))


def points(shape):
    return INTERIOR + seam(shape)
