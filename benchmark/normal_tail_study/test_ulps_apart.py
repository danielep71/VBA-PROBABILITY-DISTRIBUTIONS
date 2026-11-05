"""
Fixtures for the ULP distance of the #72 normal-tail validator (#148).

validate_kernel.ulps_apart records, for every §7 row with a committed Excel
observation, how many binary64 steps separate the model of today's code from
that observation (model_vs_committed_ulps). Until #148 it mapped -0.0 to a
huge negative integer, so a pair that crosses zero, such as (-0.0, 5e-324),
measured about 2^63 instead of 1. No recorded pair was affected.

  1. Zero: +0 and -0 are one point; each is one step from the smallest
     subnormal of either sign, and the two smallest subnormals are two apart.
  2. Ordinary steps: the subnormal/normal boundary, binade boundaries (1, 2
     and the smallest normal power of two), and the largest finite Double to
     infinity are one step, with either sign.
  3. Order: the distance is symmetric, zero only between equal values, and
     additive along the order, across the sign change as elsewhere.
  4. NaN has no position in the order: ulps_apart raises ValueError when
     either argument is NaN, so a NaN model value or observation stops the
     run instead of recording a distance.
  5. Agreement: on every pair of a set of edge values, and on seeded random
     pairs, ulps_apart equals model_fidelity.ulp_distance of the Student-t
     study (#77). The validator keeps its own copy, so that results.json does
     not depend on a file it does not bind.

It never evaluates the holdout H and writes nothing.

Run: python3 test_ulps_apart.py   (exit 0 = pass, nonzero = fail)
"""
import itertools
import math
import os
import random
import struct
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "student_t_large_df_study"))
import validate_kernel as vk                                     # noqa: E402
import model_fidelity as mf                                      # noqa: E402

TRUE_MIN = 5e-324                       # smallest positive subnormal
MAX_SUBNORMAL = struct.unpack("<d", struct.pack("<Q", 0x000FFFFFFFFFFFFF))[0]
MIN_NORMAL = 2.2250738585072014e-308    # 2^-1022
MAX_FINITE = 1.7976931348623157e308
INF = math.inf
EDGES = (0.0, TRUE_MIN, 2 * TRUE_MIN, MAX_SUBNORMAL, MIN_NORMAL, math.nextafter(MIN_NORMAL, INF),
         math.nextafter(1.0, 0.0), 1.0, math.nextafter(1.0, INF), math.nextafter(2.0, 0.0), 2.0,
         1e-300, 0.5, 3.0, 37.5, 1e300, MAX_FINITE, INF)
SIGNED = tuple(sorted({v for x in EDGES for v in (x, -x)})) + (-0.0,)


class Zero(unittest.TestCase):
    def test_signed_zeros_are_one_point(self):
        self.assertEqual(vk.ulps_apart(-0.0, 0.0), 0)
        self.assertEqual(vk.ulps_apart(0.0, -0.0), 0)
        self.assertEqual(vk.ulps_apart(-0.0, -0.0), 0)

    def test_smallest_subnormals_are_one_step_from_zero(self):
        self.assertEqual(vk.ulps_apart(-0.0, TRUE_MIN), 1)
        self.assertEqual(vk.ulps_apart(0.0, TRUE_MIN), 1)
        self.assertEqual(vk.ulps_apart(-0.0, -TRUE_MIN), 1)
        self.assertEqual(vk.ulps_apart(0.0, -TRUE_MIN), 1)
        self.assertEqual(vk.ulps_apart(-TRUE_MIN, TRUE_MIN), 2)
        self.assertEqual(vk.ulps_apart(-2 * TRUE_MIN, TRUE_MIN), 3)

    def test_crossing_zero_counts_steps_on_both_sides(self):
        self.assertEqual(vk.ulps_apart(-MIN_NORMAL, MIN_NORMAL), 2 * 0x0010000000000000)
        self.assertEqual(vk.ulps_apart(-1.0, 1.0), 2 * 0x3FF0000000000000)
        self.assertEqual(vk.ulps_apart(-INF, INF), 2 * 0x7FF0000000000000)


class Steps(unittest.TestCase):
    def assertOneStep(self, lo, hi):
        self.assertEqual(math.nextafter(lo, INF), hi, (lo, hi))
        self.assertEqual(vk.ulps_apart(lo, hi), 1, (lo, hi))
        self.assertEqual(vk.ulps_apart(-hi, -lo), 1, (-hi, -lo))

    def test_subnormal_normal_boundary(self):
        self.assertOneStep(MAX_SUBNORMAL, MIN_NORMAL)
        self.assertEqual(vk.ulps_apart(TRUE_MIN, MAX_SUBNORMAL), 0x000FFFFFFFFFFFFE)

    def test_binade_boundaries(self):
        self.assertOneStep(math.nextafter(1.0, 0.0), 1.0)
        self.assertOneStep(1.0, math.nextafter(1.0, INF))
        self.assertOneStep(math.nextafter(2.0, 0.0), 2.0)
        self.assertOneStep(MIN_NORMAL, math.nextafter(MIN_NORMAL, INF))
        # A whole binade holds 2^52 Doubles, whatever its exponent.
        self.assertEqual(vk.ulps_apart(1.0, 2.0), 1 << 52)
        self.assertEqual(vk.ulps_apart(2.0 ** -1000, 2.0 ** -999), 1 << 52)

    def test_largest_finite_to_infinity(self):
        self.assertOneStep(MAX_FINITE, INF)


class Order(unittest.TestCase):
    def test_symmetric_and_zero_only_between_equal_values(self):
        for a, b in itertools.product(SIGNED, repeat=2):
            d = vk.ulps_apart(a, b)
            self.assertEqual(d, vk.ulps_apart(b, a), (a, b))
            self.assertEqual(d == 0, a == b, (a, b))

    def test_additive_along_the_order(self):
        values = sorted(SIGNED)
        for a, b, c in itertools.combinations(values, 3):
            if a < b < c:
                self.assertEqual(vk.ulps_apart(a, c), vk.ulps_apart(a, b) + vk.ulps_apart(b, c), (a, b, c))

    def test_ordered_image_is_monotone(self):
        values = sorted(SIGNED)
        images = [vk.ordered(v) for v in values]
        self.assertEqual(images, sorted(images))
        self.assertEqual(vk.ordered(-0.0), vk.ordered(0.0))
        self.assertEqual(vk.ordered(0.0), 0)


class NaN(unittest.TestCase):
    def test_nan_raises(self):
        quiet = math.nan
        signalling = struct.unpack("<d", struct.pack("<Q", 0x7FF0000000000001))[0]
        negative = struct.unpack("<d", struct.pack("<Q", 0xFFF8000000000000))[0]
        for nan in (quiet, signalling, negative):
            self.assertTrue(math.isnan(nan))
            for other in (0.0, -0.0, 1.0, INF, nan):
                with self.assertRaises(ValueError):
                    vk.ulps_apart(nan, other)
                with self.assertRaises(ValueError):
                    vk.ulps_apart(other, nan)


class Agreement(unittest.TestCase):
    def test_edge_pairs_agree_with_model_fidelity(self):
        for a, b in itertools.product(SIGNED, repeat=2):
            self.assertEqual(vk.ulps_apart(a, b), mf.ulp_distance(a, b), (a, b))

    def test_random_pairs_agree_with_model_fidelity(self):
        rng = random.Random(148)
        checked = 0
        while checked < 20000:
            a, b = (struct.unpack("<d", struct.pack("<Q", rng.getrandbits(64)))[0] for _ in range(2))
            if math.isnan(a) or math.isnan(b):
                continue
            self.assertEqual(vk.ulps_apart(a, b), mf.ulp_distance(a, b), (a, b))
            # Near neighbours, including pairs that straddle zero.
            near = math.nextafter(a, -a if a else INF)
            self.assertEqual(vk.ulps_apart(a, near), mf.ulp_distance(a, near), (a, near))
            checked += 1


if __name__ == "__main__":
    unittest.main(verbosity=1)
