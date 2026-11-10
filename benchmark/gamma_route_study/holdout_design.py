"""The frozen #307 routing holdout (HOLDOUT_DESIGN.md). Do not edit after evaluation.

Fixed before the binary64 route map was observed, so no routing rule can be
fitted to these points.
"""
import math

# Kernel shapes, new to the repository (check_disjoint.py), spread over the
# map's domain.
SHAPES = (3.7e-6, 7.7e-5, 2.9e-3, 0.0613, 0.383, 0.871, 1.33, 2.71, 6.17,
          13.7, 37.9, 171.3, 743.0, 4111.0, 27191.0, 314159.0, 2718281.0)
# Shapes below 1: fixed arguments; from 1: X = A + k sqrt(A).
SMALL_X = (2.3e-5, 0.0371, 0.29, 1.13, 3.7, 17.3)
OFFSETS = (-7.3, -2.2, -0.6, 0.0, 0.7, 2.3, 7.9, 13.1)


def points(shape):
    if shape < 1.0:
        return tuple(sorted(SMALL_X))
    r = math.sqrt(shape)
    return tuple(sorted(shape + k * r for k in OFFSETS if shape + k * r > 0))
