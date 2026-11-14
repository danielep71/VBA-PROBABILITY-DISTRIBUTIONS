"""
Fixtures for the #314 routing holdout tools (holdout_route_build.py,
holdout_route_analyze.py, fast_route.py).

They use shapes of the revealed fit set only, never a holdout shape, and they
do not read the observed columns of the committed grid: a holdout that has not
been exported yet must not be scored by a test.

Run: python3 test_holdout_route.py   (exit 0 = pass)
"""
import copy
import csv
import math
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import fast_route as F  # noqa: E402
import holdout_design as H  # noqa: E402
import holdout_route_analyze as A  # noqa: E402
import holdout_route_build as B  # noqa: E402
from generate_binary64_map import references  # noqa: E402

SHAPES = (0.5, 10.0)        # revealed fit shapes


def synthetic():
    """Rows as the exporter would fill them if production equalled the model."""
    rows = []
    for a in SHAPES:
        xs = sorted({0.02, 0.4, a, a + 0.9} | set(B.seam_triple(a)) | {a + 5.0, a + 40.0})
        for i, x in enumerate(xs):
            p, q = F.routed(a, x)
            dp, dq = F._dd(a, x)
            ref_p, ref_q = references(a, x)
            rows.append({"point_id": f"s{a}_{i}", "kind": "holdout", "a_bits": B.hexbits(a),
                         "x_bits": B.hexbits(x), "ref_p": ref_p, "ref_q": ref_q,
                         "gamma_cdf": B.hexbits(p), "gamma_sf": B.hexbits(q),
                         "chi_cdf": B.hexbits(p), "chi_sf": B.hexbits(q),
                         "dd_p": B.hexbits(dp), "dd_q": B.hexbits(dq)})
    return rows


def prepared(rows):
    out = copy.deepcopy(rows)
    for r in out:
        r["a"], r["x"] = A.from_bits(r["a_bits"]), A.from_bits(r["x_bits"])
        r["refs"] = {"P": A.mpf(r["ref_p"]) if r["ref_p"] else None,
                     "Q": A.mpf(r["ref_q"]) if r["ref_q"] else None}
        r["exact"] = {"P": r["ref_p"] in ("0", "1"), "Q": r["ref_q"] in ("0", "1")}
        r["v"] = {k: A.value(r[k]) for k in A.PUBLIC + ("dd_p", "dd_q")}
    return out


class Rule(unittest.TestCase):
    def test_constants_agree(self):
        self.assertEqual(float(A.FLOOR), F.FLOOR)
        self.assertEqual((F.COMPLEMENT_MAX, F.UPPER_MAX), (0.9, 0.5))

    def test_shapes_above_the_ceiling_are_delegated(self):
        # Addendum 2: near X = A the binary64 deviance is wrong by more than
        # the guarantee from about shape 1E18.
        self.assertEqual(F.MAX_SHAPE, 1e15)
        self.assertNotEqual(F.paths(1e15, 1e15 + 1e8), ("dd", "dd"))
        for a in (1.0000000000000002e15, 1e18, 1e20):
            for x in (a - 2 * a ** 0.5, a + 2 * a ** 0.5):
                self.assertEqual(F.paths(a, x), ("dd", "dd"))

    def test_production_has_the_same_ceiling(self):
        with open(os.path.join(HERE, "..", "..", "src", "M_STATS_PROBDIST_SPECIALFUNCS.bas"),
                  encoding="latin-1") as handle:
            source = handle.read()
        self.assertIn("PROB_GAMMA_FAST_MAX_SHAPE As Double = 1E+15", source)
        self.assertEqual(source.count("If A > 0# And A <= PROB_GAMMA_FAST_MAX_SHAPE And X > 0# Then"), 2)

    def test_sources_follow_the_rule(self):
        self.assertEqual(F.decide("series", 0.3), ("series", "complement"))
        self.assertEqual(F.decide("series", 0.9), ("series", "complement"))
        self.assertEqual(F.decide("series", 0.95), ("series", "dd"))
        self.assertEqual(F.decide("series", 1e-300), ("dd", "complement"))
        self.assertEqual(F.decide("cf", 0.2), ("complement", "cf"))
        self.assertEqual(F.decide("cf", 1e-300), ("complement", "dd"))
        self.assertEqual(F.decide("cf", 0.6), ("dd", "dd"))
        self.assertEqual(F.decide("series", None), ("dd", "dd"))

    def test_routed_values_complement_exactly(self):
        p, q = F.routed(10.0, 4.0)
        self.assertEqual(F.paths(10.0, 4.0), ("series", "complement"))
        self.assertEqual(q, 1.0 - p)
        p, q = F.routed(10.0, 25.0)
        self.assertEqual(F.paths(10.0, 25.0), ("complement", "cf"))
        self.assertEqual(p, 1.0 - q)


class Build(unittest.TestCase):
    def test_bits_round_trip(self):
        for x in (0.0, 5e-324, 1.0, math.nextafter(1.0, 2.0), 1.7e308):
            self.assertEqual(A.from_bits(B.hexbits(x)), x)
            self.assertEqual(B.from_bits(B.bits(x)), x)

    def test_transitions_bracket_a_change_of_state(self):
        found = B.transitions(10.0, ())
        self.assertTrue(found)
        for last, first, after, old, new in found:
            self.assertEqual(math.nextafter(last, math.inf), first)
            self.assertEqual(math.nextafter(first, math.inf), after)
            self.assertEqual(F.paths(10.0, last), old)
            self.assertEqual(F.paths(10.0, first), new)
            self.assertNotEqual(old, new)
        self.assertIn(11.0, [first for _, first, _, _, _ in found])    # the X = A + 1 seam

    def test_the_committed_grid_is_the_design(self):
        with open(B.GRID, newline="") as handle:
            rows = list(csv.DictReader(handle))
        held = [r for r in rows if r["kind"] == "holdout"]
        self.assertEqual(len(held), sum(len(H.points(a)) for a in H.SHAPES))
        self.assertEqual(len(held), 117)
        self.assertEqual({r["a_bits"] for r in held}, {B.hexbits(a) for a in H.SHAPES})
        self.assertEqual(len({(r["a_bits"], r["x_bits"]) for r in rows}), len(rows))
        self.assertTrue(all(r["ref_p"] and r["ref_q"] for r in rows))
        self.assertEqual({r["kind"] for r in rows}, {"holdout", "zero", "transition", "shape_probe"})


class Verdict(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.rows = synthetic()

    def run_on(self, rows):
        return A.analyze(prepared(rows), model=F.routed)

    def test_the_model_passes_every_rule(self):
        fails, report = self.run_on(self.rows)
        self.assertEqual(fails, [])
        self.assertEqual(report["counts"]["equal to model"], report["counts"]["compared with model"])
        self.assertLess(max(report["worst"].values()), 1e-14)

    def change(self, index, **cells):
        rows = copy.deepcopy(self.rows)
        rows[index].update(cells)
        return rows

    def test_r1_an_inaccurate_value_is_rejected(self):
        p = A.from_bits(self.rows[1]["gamma_cdf"])
        off = B.hexbits(p * (1 + 5e-13))
        fails, _ = self.run_on(self.change(1, gamma_cdf=off, chi_cdf=off))
        self.assertTrue(any(f.startswith("R1") for f in fails), fails)

    def test_r2_the_two_families_must_agree(self):
        p = A.from_bits(self.rows[1]["gamma_cdf"])
        fails, _ = self.run_on(self.change(1, chi_cdf=B.hexbits(math.nextafter(p, 2.0))))
        self.assertTrue(any(f.startswith("R2") for f in fails), fails)

    def test_r3_a_reversal_beyond_the_allowances_is_rejected(self):
        rows = copy.deepcopy(self.rows)
        for key in ("gamma_cdf", "chi_cdf"):
            rows[1][key], rows[2][key] = rows[2][key], rows[1][key]
        fails, _ = self.run_on(rows)
        self.assertTrue(any(f.startswith("R3") for f in fails), fails)

    def test_r3_one_ulp_backwards_is_within_the_allowances(self):
        rows = copy.deepcopy(self.rows)
        seam = [i for i, r in enumerate(rows) if r["x_bits"] == B.hexbits(11.0)][0]
        p = A.from_bits(rows[seam - 1]["gamma_cdf"])
        low = B.hexbits(math.nextafter(p, 0.0))
        rows[seam].update(gamma_cdf=low, chi_cdf=low)
        fails, _ = self.run_on(rows)
        self.assertFalse(any(f.startswith("R3") for f in fails), fails)

    def test_r4_the_tails_must_sum_to_one(self):
        q = A.from_bits(self.rows[1]["gamma_sf"])
        off = B.hexbits(q * (1 - 1e-12))
        fails, _ = self.run_on(self.change(1, gamma_sf=off, chi_sf=off))
        self.assertTrue(any(f.startswith("R4") for f in fails), fails)

    def test_r5_a_tail_below_the_floor_must_be_delegated(self):
        rows = copy.deepcopy(self.rows)
        tiny = B.hexbits(1e-300)
        rows[0].update(ref_q="1e-300", gamma_sf=tiny, chi_sf=tiny, dd_q=B.hexbits(2e-300))
        fails, _ = self.run_on(rows)
        self.assertTrue(any(f.startswith("R5") for f in fails), fails)
        rows[0].update(dd_q=tiny)
        fails, _ = self.run_on(rows)
        self.assertFalse(any(f.startswith("R5") for f in fails), fails)

    def test_r6_a_status_is_not_a_number(self):
        for status in ("CVERR", "ERROR", ""):
            fails, _ = self.run_on(self.change(1, gamma_cdf=status))
            self.assertTrue(any(f.startswith("R6") for f in fails), status)

    def test_exact_references_need_the_exact_double(self):
        rows = copy.deepcopy(self.rows)
        one, zero = B.hexbits(1.0), B.hexbits(0.0)
        last = len(rows) - 1                    # the largest X of the last shape
        rows[last].update(ref_p="1", ref_q="0", gamma_cdf=one, chi_cdf=one, gamma_sf=zero,
                          chi_sf=zero)
        self.assertEqual(self.run_on(rows)[0], [])
        below = B.hexbits(math.nextafter(1.0, 0.0))
        rows[last].update(gamma_cdf=below, chi_cdf=below)
        fails, _ = self.run_on(rows)
        self.assertTrue(any(f.startswith("R1") for f in fails), fails)


if __name__ == "__main__":
    unittest.main()
