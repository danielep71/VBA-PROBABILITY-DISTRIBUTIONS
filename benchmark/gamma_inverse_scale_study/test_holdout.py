"""
#14: fixtures for the model, the holdout construction and each acceptance
rule. They use another seed and other seam shapes than the holdout, so no
holdout point is evaluated here.

Run: python3 test_holdout.py
"""
import copy
import os
import sys
import unittest

from mpmath import mp, mpf

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import holdout_analyze as A  # noqa: E402
import holdout_build as B  # noqa: E402
import inv_scale_model as M  # noqa: E402

DEV_SEED = 1400
DEV_SEAM_SHAPES = (0.05, 0.3, 3.0)


def fill(row):
    """The row as an exact implementation of the rule would export it: the
    model in log space, the correctly rounded reference elsewhere."""
    p, a, s = (B.from_bits(row[k]) for k in ("p_bits", "a_bits", "s_bits"))
    out = dict(row)
    if M.uses_log(p, a):
        value = M.candidate(p, a, s)
    elif row["ref_x"] == "NUM":
        value = None
    else:
        value = float(mpf(row["ref_x"]))
    out["gamma_inv"] = "CVERR" if value is None else B.bits(value)
    out["chi_inv"] = out["gamma_inv"] if s == 2.0 else "NA"
    out["gamma_cdf"] = B.bits(p) if value is not None and value >= M.MIN_NORMAL else "NA"
    out["log_unit"] = B.bits(M.log_unit(p, a))
    return out


class Fixture(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        mp.dps = 60
        cls.rows = [fill(r) for r in B.build(seed=DEV_SEED, seam_shapes=DEV_SEAM_SHAPES)]

    def verdict(self, rows=None):
        return A.analyze(self.rows if rows is None else rows)[0]

    def edit(self, kind, **cells):
        rows = copy.deepcopy(self.rows)
        row = next(r for r in rows if kind(r))
        row.update(cells)
        return rows, row


class Model(Fixture):
    def test_the_counterexample_is_representable(self):
        value = M.candidate(1e-4, 0.01, 1e100)
        self.assertTrue(M.uses_log(1e-4, 0.01))
        self.assertAlmostEqual(value / 5.660738147062017e-301, 1.0, delta=1e-14)

    def test_log_space_matches_the_exact_quantile(self):
        worst = mpf(0)
        for r in self.rows:
            p, a, s = (B.from_bits(r[k]) for k in ("p_bits", "a_bits", "s_bits"))
            if r["ref_x"] in ("NUM", "ZERO") or not M.uses_log(p, a):
                continue
            ref = mpf(r["ref_x"])
            if ref < A.MIN_NORMAL:
                continue
            worst = max(worst, abs(mpf(M.candidate(p, a, s)) - ref) / ref)
        self.assertLess(worst, mpf("2e-15"))

    def test_a_final_underflow_is_the_exact_zero(self):
        self.assertEqual(M.candidate(1e-4, 0.01, 1e-300), 0.0)

    def test_arguments_outside_the_pair_logarithm_are_rescaled_exactly(self):
        # Addendum 1: a subnormal probability, and a scale from 2^978.
        for p, a, s in ((5e-320, 1.0, 1e300), (1e-4, 0.01, 1e100), (1e-100, 0.25, 8e307),
                        (3e-310, 0.9, 1.7e308)):
            self.assertTrue(M.uses_log(p, a))
            value, ref = M.candidate(p, a, s), M.exact(p, a, s)
            self.assertLess(abs(mpf(value) - ref) / ref, mpf("2e-15"), (p, a, s))


class Construction(Fixture):
    def test_the_committed_grid_is_the_design(self):
        self.assertEqual(B.main(["--check"]), 0)

    def test_sizes(self):
        rows = B.build()
        kinds = [r["kind"] for r in rows]
        self.assertEqual(len(rows), 455)
        self.assertEqual([kinds.count(k) for k in ("underflowed", "log_region", "seam",
                                                   "chisquare", "limits")],
                         [120, 120, 130, 60, 25])
        self.assertEqual(len({r["point_id"] for r in rows}), 455)

    def test_the_seam_brackets_the_crossover(self):
        for a in DEV_SEAM_SHAPES:
            probes = B.seam(a)
            self.assertEqual(probes, sorted(probes))
            self.assertEqual([M.uses_log(p, a) for p in probes], [True] * 7 + [False] * 6)

    def test_underflowed_rows_are_what_they_claim(self):
        for r in self.rows:
            if r["kind"] != "underflowed":
                continue
            p, a = B.from_bits(r["p_bits"]), B.from_bits(r["a_bits"])
            self.assertLess(M.log_unit(p, a), B.LOG_MIN_NORMAL)
            self.assertGreaterEqual(mpf(r["ref_x"]), A.MIN_NORMAL)

    def test_the_holdout_is_not_the_fixture(self):
        holdout = {tuple(r[k] for k in ("p_bits", "a_bits", "s_bits")) for r in B.build()}
        fixture = {tuple(r[k] for k in ("p_bits", "a_bits", "s_bits")) for r in self.rows}
        # The fixed rows are shared: the counterexample and the shape-1 limits.
        self.assertLessEqual(len(holdout & fixture), 6)


class Rules(Fixture):
    def test_the_model_passes_every_rule(self):
        self.assertEqual(self.verdict(), [])

    def test_a_blank_grid_is_rejected(self):
        failures = self.verdict(B.build(seed=DEV_SEED, seam_shapes=DEV_SEAM_SHAPES))
        self.assertEqual(len(failures), len(self.rows))

    def test_g1_an_inaccurate_quantile_is_rejected(self):
        rows, row = self.edit(lambda r: r["kind"] == "underflowed")
        row["gamma_inv"] = B.bits(B.from_bits(row["gamma_inv"]) * (1 + 3e-12))
        self.assertTrue(any(f.startswith("G1") for f in self.verdict(rows)))

    def test_g1_a_zero_where_a_number_is_due_is_rejected(self):
        # The defect of #14 itself.
        rows, row = self.edit(lambda r: r["kind"] == "underflowed")
        row.update(gamma_inv=A.ZERO_BITS, gamma_cdf="NA")
        self.assertTrue(any(f.startswith("G1") for f in self.verdict(rows)))

    def test_g2_limits(self):
        rows, row = self.edit(lambda r: r["ref_x"] == "ZERO")
        row["gamma_inv"] = B.bits(5e-324)
        self.assertTrue(any(f.startswith("G2") for f in self.verdict(rows)))
        rows, row = self.edit(lambda r: r["ref_x"] == "NUM")
        row["gamma_inv"] = B.bits(1.7e308)
        self.assertTrue(any(f.startswith("G2") for f in self.verdict(rows)))

    def test_g3_a_reversal_beyond_the_allowances_is_rejected(self):
        rows = copy.deepcopy(self.rows)
        cluster = [r for r in rows if r["kind"] == "seam"][:13]
        cluster[5]["gamma_inv"] = B.bits(B.from_bits(cluster[5]["gamma_inv"]) * (1 - 5e-12))
        self.assertTrue(any(f.startswith("G3") for f in self.verdict(rows)))

    def test_g3_one_ulp_backwards_is_within_the_allowances(self):
        rows = copy.deepcopy(self.rows)
        cluster = [r for r in rows if r["kind"] == "seam"][:13]
        cluster[5]["gamma_inv"] = B.bits(B.step(B.from_bits(cluster[4]["gamma_inv"]), -1))
        self.assertFalse(any(f.startswith("G3") for f in self.verdict(rows)))

    def test_g4_the_two_families_must_agree(self):
        rows, row = self.edit(lambda r: r["kind"] == "chisquare")
        row["chi_inv"] = B.bits(B.step(B.from_bits(row["gamma_inv"]), 1))
        self.assertTrue(any(f.startswith("G4") for f in self.verdict(rows)))

    def test_g5_the_round_trip_must_return_the_probability(self):
        rows, row = self.edit(lambda r: r["kind"] == "log_region")
        row["gamma_cdf"] = B.bits(B.from_bits(row["p_bits"]) * (1 + 1e-10))
        self.assertTrue(any(f.startswith("G5") for f in self.verdict(rows)))

    def test_g6_a_fault_is_not_a_number(self):
        for cell, text in (("gamma_inv", "ERROR"), ("log_unit", "FAIL"), ("gamma_cdf", "ERROR")):
            rows, row = self.edit(lambda r: r["kind"] == "log_region")
            row[cell] = text
            self.assertTrue(any(f.startswith("G6") for f in self.verdict(rows)), cell)

    def test_g7_a_cluster_on_one_side_only_is_rejected(self):
        rows = copy.deepcopy(self.rows)
        for r in [r for r in rows if r["kind"] == "seam"][:13]:
            r["log_unit"] = B.bits(-41.0)
        self.assertTrue(any(f.startswith("G7") for f in self.verdict(rows)))


if __name__ == "__main__":
    unittest.main()
