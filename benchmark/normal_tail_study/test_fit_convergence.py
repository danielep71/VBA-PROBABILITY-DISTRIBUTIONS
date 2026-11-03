"""
Fixtures for the fit convergence decision of the #72 normal-tail study (#79,
PREREGISTRATION.md Amendment 3).

  1. The committed record: fit_convergence.json is CONVERGED for the
     committed kernel_coefficients.json, under the Amendment 3 criterion, with
     the full 120-iteration history; its status follows from that history and
     the §2 selection; the committed 40-digit coefficients, recomputed here,
     alternate at 16 points and are levelled within 1%; results.json records
     the same check as PASS.
  2. The decision itself, on synthetic error curves: a levelled 16-point
     alternation is CONVERGED; too few alternations, a ratio above 1%, a
     non-finite error or a nonpositive denominator is not; an unlevelled
     selected iterate is CAP_EXHAUSTED when it is the last one and
     NOT_CONVERGED otherwise; an empty history is NOT_CONVERGED.
  3. Deliberate fits, small and fast, through the real fit(): two iterations
     of a (3, 3) fit stop at the cap unlevelled (CAP_EXHAUSTED); forty
     iterations of the same fit pass their smallest node maximum at iterate 38
     unlevelled (NOT_CONVERGED); a (1, 1) fit converges.
  4. derive_kernel.py refuses before recommending: a capped, unconverged or
     missing-status fit writes no coefficient file, leaves an existing one
     byte-identical, and exits 1; a converged fit that would change the
     recorded coefficients is refused and writes nothing.
  5. The validator refuses the same: a missing record, a missing status,
     NOT_CONVERGED, CAP_EXHAUSTED, a record for other coefficients or under
     another criterion, a status the history does not support, a selection
     that breaks §2, and selected figures that the committed coefficients do
     not reproduce all fail the check, and so the study.

It never evaluates the holdout H and never writes outside a temporary
directory.

Run: python3 test_fit_convergence.py   (exit 0 = pass, nonzero = fail)
"""
import copy
import hashlib
import json
import os
import shutil
import sys
import tempfile
import unittest

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import derive_kernel as dk                                       # noqa: E402
import kernel_model as km                                        # noqa: E402
import validate_kernel as vk                                     # noqa: E402

POINTS = dk.DEGREE_P + dk.DEGREE_D + 2


def sha(path):
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


def evidence(errors, points=POINTS, denominator=1.0):
    xs = [float(i) for i in range(len(errors))]
    return dk.measure(xs, errors, [denominator] * len(errors), points)


def equioscillating(n, size=1e-17, wobble=0.0):
    """n alternating extrema of |e| = size * (1 + k * wobble / n), with a small point between each."""
    out = []
    for k in range(n):
        s = 1 if k % 2 == 0 else -1
        out += [s * size * (1 + k * wobble / n), s * size / 3]
    return out


def small_fit(degree, nodes, iterations):
    with mp.workdps(30):
        return dk.fit(degree_p=degree, degree_d=degree, nodes=nodes, iterations=iterations)


class CommittedRecord(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.record = vk.load_convergence()
        cls.raw = km.COEFFICIENTS["raw"]
        cls.coefficients_sha = sha(km.COEFFICIENTS_PATH)
        with open(vk.OUT, encoding="utf-8") as f:
            cls.results = json.load(f)

    def test_record_is_converged_for_the_committed_coefficients(self):
        r = self.record
        self.assertEqual(dk.recommendable(r), [])
        self.assertEqual(r["status"], dk.CONVERGED)
        self.assertEqual(r["kernel_coefficients_sha256"], self.coefficients_sha)
        self.assertEqual(r["criterion"]["levelling_tolerance"], 0.01)
        self.assertEqual(r["criterion"]["alternation_points_required"], 16)
        self.assertEqual(r["criterion"]["iteration_cap"], 120)
        self.assertEqual(len(r["history"]), 120)
        self.assertEqual(r["selected"], r["history"][r["selected_iteration"] - 1])
        self.assertEqual(dk.decide(r["history"], r["selected_iteration"], POINTS), dk.CONVERGED)
        self.assertEqual(self.raw["fit"]["node_max_relative_error"], mp.nstr(mp.mpf(r["selected"]["node_max"]), 6))

    def test_observed_evidence_matches_amendment_3(self):
        r = self.record
        maxima = [h["node_max"] for h in r["history"]]
        self.assertTrue(all(b < a for a, b in zip(maxima, maxima[1:])), "node maximum not falling throughout")
        self.assertEqual({h["alternation_points"] for h in r["history"]}, {16})
        self.assertEqual(r["selected_iteration"], 120)
        self.assertEqual(r["first_levelled_iteration"], 53)
        self.assertLess(r["selected"]["levelling"], 3.34e-3)
        self.assertGreater(r["selected"]["levelling"], 3.33e-3)

    def test_validator_recomputes_the_committed_coefficients(self):
        check = vk.convergence_check(self.raw, self.record, self.coefficients_sha)
        self.assertTrue(check["pass"], check["problems"])
        m = check["recomputed_from_committed_coefficients"]
        self.assertEqual(m["alternation_points"], 16)
        self.assertLessEqual(m["levelling"], dk.LEVELLING_TOLERANCE)
        self.assertEqual(self.results["acceptance"]["fit_convergence"], json.loads(json.dumps(check)))


class Decision(unittest.TestCase):
    def test_levelled_alternation_converges(self):
        m = evidence(equioscillating(16, wobble=0.005))
        self.assertEqual(m["alternation_points"], 16)
        self.assertTrue(dk.levelled(m, POINTS))
        self.assertEqual(dk.decide([m, m], 1, POINTS), dk.CONVERGED)

    def test_too_few_alternations(self):
        m = evidence(equioscillating(15))
        self.assertIsNone(m["levelled_min"])
        self.assertFalse(dk.levelled(m, POINTS))

    def test_unlevelled_is_capped_at_the_last_iterate_and_unconverged_before(self):
        m = evidence(equioscillating(16, wobble=0.05))
        self.assertGreater(m["levelling"], 0.01)
        self.assertEqual(dk.decide([m, m, m], 3, POINTS), dk.CAP_EXHAUSTED)
        self.assertEqual(dk.decide([m, m, m], 2, POINTS), dk.NOT_CONVERGED)

    def test_lower_bound_is_the_largest_alternating_level(self):
        errors = equioscillating(16)
        errors[0] = 1e-15                    # one large extremum; the bound stays at the levelled size
        m = evidence(errors)
        self.assertEqual(m["levelled_min"], 1e-17)
        self.assertAlmostEqual(m["levelling"], 99.0)
        self.assertFalse(dk.levelled(m, POINTS))

    def test_non_finite_and_nonpositive(self):
        errors = equioscillating(16)
        errors[3] = float("nan")
        self.assertFalse(dk.levelled(evidence(errors), POINTS))
        self.assertFalse(dk.levelled(evidence(equioscillating(16), denominator=-1.0), POINTS))
        self.assertFalse(dk.levelled(evidence(equioscillating(16), denominator=0.0), POINTS))

    def test_empty_or_unselected(self):
        m = evidence(equioscillating(16))
        self.assertEqual(dk.decide([], 1, POINTS), dk.NOT_CONVERGED)
        self.assertEqual(dk.decide([m], None, POINTS), dk.NOT_CONVERGED)
        self.assertEqual(dk.decide([m], 2, POINTS), dk.NOT_CONVERGED)

    def test_missing_status_is_not_recommendable(self):
        self.assertTrue(dk.recommendable(None))
        self.assertTrue(dk.recommendable({}))
        self.assertTrue(dk.recommendable({"status": "MAYBE"}))
        self.assertTrue(dk.recommendable({"status": dk.CAP_EXHAUSTED}))
        self.assertTrue(dk.recommendable({"status": dk.NOT_CONVERGED}))
        self.assertEqual(dk.recommendable({"status": dk.CONVERGED}), [])


class DeliberateFits(unittest.TestCase):
    def test_cap_exhausted(self):
        _, _, _, conv = small_fit(3, 60, 2)
        self.assertEqual(conv["status"], dk.CAP_EXHAUSTED)
        self.assertEqual(conv["selected_iteration"], 2)
        self.assertGreater(conv["selected"]["levelling"], 0.01)

    def test_not_converged(self):
        _, _, _, conv = small_fit(3, 60, 40)
        self.assertEqual(conv["status"], dk.NOT_CONVERGED)
        self.assertEqual(conv["selected_iteration"], 38)
        self.assertIsNone(conv["first_levelled_iteration"])
        later = conv["history"][38]["node_max"]
        self.assertGreater(later, conv["selected"]["node_max"], "the node maximum should rise after iterate 38")

    def test_small_fit_converges(self):
        _, _, _, conv = small_fit(1, 30, 60)
        self.assertEqual(conv["status"], dk.CONVERGED)
        self.assertEqual(conv["selected"]["alternation_points"], 4)


class DerivationRefuses(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.out = os.path.join(self.tmp, "kernel_coefficients.json")
        self.conv = os.path.join(self.tmp, "fit_convergence.json")
        self.logs = []

    def tearDown(self):
        shutil.rmtree(self.tmp)

    def run_main(self, fitter, argv=()):
        return dk.main(list(argv), out=self.out, convergence_out=self.conv, fitter=fitter, log=self.logs.append)

    def committed(self):
        shutil.copy(km.COEFFICIENTS_PATH, self.out)
        with open(self.out, "rb") as f:
            return f.read()

    def test_capped_fit_writes_no_candidate(self):
        result = small_fit(3, 60, 2)
        self.assertEqual(self.run_main(lambda: result), 1)
        self.assertFalse(os.path.exists(self.out))
        with open(self.conv, encoding="utf-8") as f:
            recorded = json.load(f)
        self.assertEqual(recorded["status"], dk.CAP_EXHAUSTED)
        self.assertIsNone(recorded["kernel_coefficients_sha256"])
        self.assertIn("no candidate recommended", "\n".join(self.logs))

    def test_unconverged_fit_leaves_the_recorded_candidate(self):
        before = self.committed()
        result = small_fit(3, 60, 40)
        self.assertEqual(self.run_main(lambda: result), 1)
        with open(self.out, "rb") as f:
            self.assertEqual(f.read(), before)

    def test_missing_status_is_refused(self):
        before = self.committed()
        node_max, p, d, conv = small_fit(1, 30, 60)
        for broken in ({k: v for k, v in conv.items() if k != "status"}, None):
            self.assertEqual(self.run_main(lambda b=broken: (node_max, p, d, b)), 1)
            with open(self.out, "rb") as f:
                self.assertEqual(f.read(), before)
        self.assertIn("WITHOUT STATUS", "\n".join(self.logs))

    def test_converged_fit_that_changes_the_candidate_is_refused(self):
        before = self.committed()
        result = small_fit(1, 30, 60)
        self.assertEqual(result[3]["status"], dk.CONVERGED)
        self.assertEqual(self.run_main(lambda: result), 1)
        with open(self.out, "rb") as f:
            self.assertEqual(f.read(), before)
        self.assertFalse(os.path.exists(self.conv))
        self.assertIn("differ from the recorded", "\n".join(self.logs))

    def test_converged_new_fit_is_written(self):
        result = small_fit(1, 30, 60)
        self.assertEqual(self.run_main(lambda: result), 0)
        with open(self.conv, encoding="utf-8") as f:
            self.assertEqual(json.load(f)["kernel_coefficients_sha256"], sha(self.out))
        self.assertEqual(self.run_main(lambda: result, ["--verify"]), 0)


class ValidatorRefuses(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.record = vk.load_convergence()
        cls.raw = km.COEFFICIENTS["raw"]
        cls.coefficients_sha = sha(km.COEFFICIENTS_PATH)

    def fails(self, record, expect, coefficients_sha=None):
        check = vk.convergence_check(self.raw, record, coefficients_sha or self.coefficients_sha)
        self.assertFalse(check["pass"])
        self.assertTrue(any(expect in p for p in check["problems"]), check["problems"])

    def mutated(self, change):
        r = copy.deepcopy(self.record)
        change(r)
        return r

    def test_missing_record(self):
        self.fails(None, "no convergence record")
        self.assertIsNone(vk.load_convergence(os.path.join(HERE, "no_such_record.json")))

    def test_missing_status(self):
        self.fails(self.mutated(lambda r: r.pop("status")), "has no status")

    def test_not_converged_and_capped(self):
        self.fails(self.mutated(lambda r: r.update(status=dk.NOT_CONVERGED)), "NOT_CONVERGED")
        self.fails(self.mutated(lambda r: r.update(status=dk.CAP_EXHAUSTED)), "CAP_EXHAUSTED")

    def test_other_coefficients(self):
        self.fails(self.record, "belongs to coefficients", coefficients_sha="0" * 64)

    def test_other_criterion(self):
        self.fails(self.mutated(lambda r: r["criterion"].update(levelling_tolerance=0.1)), "another criterion")

    def test_status_not_supported_by_history(self):
        def unlevel(r):
            for h in r["history"]:
                h["levelling"] = 0.02
            r["selected"] = r["history"][r["selected_iteration"] - 1]
        self.fails(self.mutated(unlevel), "does not follow from the recorded evidence")

    def test_selection_breaks_section_2(self):
        def earlier(r):
            r["selected_iteration"] = 60
            r["selected"] = r["history"][59]
        self.fails(self.mutated(earlier), "not the first with the smallest node maximum")

    def test_selected_figures_not_reproduced(self):
        def inflate(r):
            r["selected"]["node_max"] *= 1.001
            r["history"][r["selected_iteration"] - 1] = r["selected"]
        self.fails(self.mutated(inflate), "does not match the committed coefficients")


if __name__ == "__main__":
    unittest.main(verbosity=1)
