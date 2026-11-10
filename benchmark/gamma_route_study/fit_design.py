"""The #307 binary64 route fit set (FIT_DESIGN.md). Do not edit after export.

Fitting points only: the binary64 route map's points, more shapes for the
cost, and edge cases of the route's domain. None is a holdout point
(check_disjoint.py), so the routing rule fitted here can still be tested on
holdout_design.py.
"""
import math

from design import OFFSETS, SHAPES as MAP_SHAPES, points as map_points, seam

# Shapes between the map's decades, for the cost and accuracy of large shapes.
COST_SHAPES = (2000.0, 3000.0, 5000.0, 20000.0, 30000.0, 50000.0, 200000.0, 500000.0)
# Edge cases: X / A subnormal (the route declines), the factor floor at A > 1,
# the small-shape factor's X >= 708 floor, a tiny shape, a tiny normal X / A.
EDGES = ((1.5, 1e-310), (2.0, 800.0), (0.5, 800.0), (1e-8, 0.5), (3.0, 1e-305), (0.75, 700.0))


def cost_points(shape):
    r = math.sqrt(shape)
    xs = {shape + k * r for k in OFFSETS if shape + k * r > 0}
    xs |= set(seam(shape))
    return tuple(sorted(xs))


def rows():
    """(point_id, shape, x) in a fixed order."""
    out = []
    for i, a in enumerate(MAP_SHAPES):
        out += [(f"m{i:02d}_x{j:02d}", a, x) for j, x in enumerate(map_points(a))]
    for i, a in enumerate(COST_SHAPES):
        out += [(f"c{i:02d}_x{j:02d}", a, x) for j, x in enumerate(cost_points(a))]
    out += [(f"e{i:02d}", a, x) for i, (a, x) in enumerate(EDGES)]
    return out
