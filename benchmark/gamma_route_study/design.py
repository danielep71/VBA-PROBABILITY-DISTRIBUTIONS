"""The frozen #307 binary64 route map (MAP_DESIGN.md). Do not edit after export."""
import math

# Kernel shapes A: two to four per decade from 1E-6 to 1E7, denser near 1,
# where the pre-#242 failures clustered.
SHAPES = (1e-6, 1e-5, 1e-4, 1e-3, 0.01, 0.05, 0.1, 0.25, 0.5, 0.75, 1.0, 1.5,
          2.0, 3.0, 5.0, 10.0, 20.0, 50.0, 100.0, 300.0, 1000.0, 1e4, 1e5, 1e6, 1e7)
# Shapes below 1: fixed arguments spanning the lower region and both tails.
SMALL_X = (1e-6, 1e-3, 0.05, 0.5, 2.0, 5.0, 10.0, 30.0)
# Shapes from 1: X = A + k * sqrt(A), from the far lower tail to the far upper tail.
OFFSETS = (-10.0, -6.0, -3.0, -1.0, 0.0, 1.0, 3.0, 6.0, 10.0, 20.0)


def seam(shape):
    """The binary64 predecessor of A + 1, A + 1 itself, and its successor."""
    s = shape + 1.0
    return (math.nextafter(s, 0.0), s, math.nextafter(s, math.inf))


def points(shape):
    if shape < 1.0:
        xs = set(SMALL_X)
    else:
        r = math.sqrt(shape)
        xs = {shape + k * r for k in OFFSETS if shape + k * r > 0}
    xs |= set(seam(shape))
    return tuple(sorted(xs))
