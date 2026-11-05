"""
Fixtures for model_fidelity.py (#77): true binary64 ULP distance and the
source binding of finding F1's model.

  1. ULP distance counts representable steps: +0 and -0 coincide; the
     smallest subnormal is one step from zero and two from its negative; the
     subnormal/normal boundary, binade boundaries and infinity are single
     steps; stepping k times with nextafter gives distance k; NaN is refused.
  2. Negative fixtures for the metric: a three-step mismatch fails the
     two-ULP fidelity limit, including at the committed observation
     NormalStandard_Survival(1.96), where the retired relative-epsilon scale
     |a - b| / (|a| 2^-52) reported 1.875 for the same mismatch and passed it.
  3. The committed binding is current and binds exactly the declared
     procedures and constants.
  4. Mutation fixtures, on a copy of the source and model: a Hart
     coefficient, the 7.07 split, the continued-fraction depth, a route
     predicate, the density, a comment-only edit and a model edit each make
     the binding stale, naming what changed; a missing or duplicated
     procedure and a missing record fail closed; CRLF and LF checkouts bind
     identically.
  5. Re-recording: --record refuses while the model does not mirror the
     source and leaves the record untouched; after a consistent change of
     source and model it records the changed identities in the history, and
     the binding is current again.
  6. The F1 fixture itself (test_binary64_route_model.py) fails on a copy
     whose VBA continued-fraction depth was changed, reporting the binding.

It writes only in temporary directories.

Run: python3 test_model_fidelity.py   (exit 0 = pass, nonzero = fail)
"""
import json
import math
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import model_fidelity as F                                       # noqa: E402

MIN_SUBNORMAL = 5e-324
MIN_NORMAL = 2.2250738585072014e-308
MAX_SUBNORMAL = math.nextafter(MIN_NORMAL, 0.0)
MAX_FINITE = 1.7976931348623157e308
SURVIVAL_196 = 0.024997895148220456          # committed Excel NormalStandard_Survival(1.96)


def steps(x, k):
    """x moved k representable steps (k < 0 moves down)."""
    for _ in range(abs(k)):
        x = math.nextafter(x, math.inf if k > 0 else -math.inf)
    return x


def relative_epsilon_units(observed, modelled):
    """The retired metric of test_binary64_route_model.py, kept here only to show what it admitted."""
    return abs(observed - modelled) / (abs(observed) * 2.0 ** -52)


class UlpDistance(unittest.TestCase):
    def test_zero_and_sign(self):
        self.assertEqual(F.ulp_distance(0.0, -0.0), 0)
        self.assertEqual(F.ordered(-0.0), F.ordered(0.0))
        self.assertEqual(F.ulp_distance(0.0, MIN_SUBNORMAL), 1)
        self.assertEqual(F.ulp_distance(-0.0, MIN_SUBNORMAL), 1)
        self.assertEqual(F.ulp_distance(-MIN_SUBNORMAL, MIN_SUBNORMAL), 2)
        self.assertEqual(F.ulp_distance(-1.0, 1.0), 2 * F.ordered(1.0))

    def test_subnormal_and_binade_boundaries(self):
        self.assertEqual(F.ulp_distance(MAX_SUBNORMAL, MIN_NORMAL), 1)
        self.assertEqual(F.ulp_distance(steps(MIN_NORMAL, -3), steps(MIN_NORMAL, 2)), 5)
        below = math.nextafter(1.0, 0.0)
        above = math.nextafter(1.0, 2.0)
        self.assertEqual(F.ulp_distance(below, 1.0), 1)
        self.assertEqual(F.ulp_distance(1.0, above), 1)
        self.assertEqual(F.ulp_distance(below, above), 2)
        self.assertEqual(above - 1.0, 2 * (1.0 - below), "the spacing doubles across the boundary")
        self.assertEqual(F.ulp_distance(-below, -above), 2)
        self.assertEqual(F.ulp_distance(MAX_FINITE, math.inf), 1)

    def test_stepping_counts(self):
        for x in (MIN_SUBNORMAL, MAX_SUBNORMAL, MIN_NORMAL, 0.5, 1.0, 2.0, SURVIVAL_196, 0.975, 1e300,
                  -3.5, -MIN_NORMAL, 0.0, -0.0):
            for k in (-5, -2, -1, 1, 2, 5):
                y = steps(x, k)
                self.assertEqual(F.ulp_distance(x, y), abs(k), (x, k))
                self.assertEqual(F.ulp_distance(y, x), abs(k), (x, k))

    def test_agrees_with_spacing_inside_a_binade(self):
        for x in (SURVIVAL_196, 0.15865525393145705, 0.8413447460685429, 1e-310, 3.0):
            for k in (1, 2, 3):
                self.assertEqual(abs(steps(x, k) - x) / math.ulp(x), F.ulp_distance(x, steps(x, k)))

    def test_nan_is_refused(self):
        with self.assertRaises(ValueError):
            F.ulp_distance(math.nan, 1.0)
        with self.assertRaises(ValueError):
            F.within_ulps(1.0, math.nan)


class MetricNegativeFixtures(unittest.TestCase):
    def test_three_steps_fail_the_two_ulp_limit(self):
        self.assertEqual(F.FIDELITY_LIMIT_ULPS, 2)
        for x in (SURVIVAL_196, math.nextafter(2.0, 0.0), 1.0, MIN_NORMAL, MIN_SUBNORMAL * 7, -0.975):
            for direction in (1, -1):
                self.assertTrue(F.within_ulps(x, steps(x, 2 * direction)), (x, direction))
                self.assertFalse(F.within_ulps(x, steps(x, 3 * direction)), (x, direction))

    def test_retired_scale_admitted_three_steps(self):
        three = steps(SURVIVAL_196, -3)
        self.assertAlmostEqual(relative_epsilon_units(SURVIVAL_196, three), 1.875, places=2)
        self.assertLessEqual(relative_epsilon_units(SURVIVAL_196, three), 2.0)
        self.assertFalse(F.within_ulps(SURVIVAL_196, three))
        top = math.nextafter(2.0, 0.0)
        self.assertLessEqual(relative_epsilon_units(top, steps(top, -3)), 1.51)
        self.assertFalse(F.within_ulps(top, steps(top, -3)))

    def test_committed_196_observation(self):
        modelled = 0.02499789514822045
        self.assertAlmostEqual(relative_epsilon_units(SURVIVAL_196, modelled), 1.25, places=2)
        self.assertEqual(F.ulp_distance(SURVIVAL_196, modelled), 2)


class CommittedBinding(unittest.TestCase):
    def test_current(self):
        record = F.load_record()
        self.assertEqual(F.stale(record), [])
        self.assertEqual(list(record["procedures"]), list(F.PROCEDURES))
        self.assertEqual(list(record["constants"]), [name for name, _ in F.CONSTANTS])
        self.assertEqual((record["source"], record["model"]), (F.SOURCE, F.MODEL))
        self.assertTrue(record["history"])
        self.assertEqual(F.extraction_problems(), [])


class Copy(unittest.TestCase):
    """A temporary root with the bound source, the study folder and the grids F1 reads."""

    def setUp(self):
        self.root = tempfile.mkdtemp()
        self.study = os.path.join(self.root, "benchmark", "student_t_large_df_study")
        shutil.copytree(HERE, self.study, ignore=shutil.ignore_patterns("__pycache__"))
        for name in (F.SOURCE, "benchmark/probability_accuracy_grid.csv", "benchmark/holdout/holdout_grid.csv"):
            target = os.path.join(self.root, name)
            os.makedirs(os.path.dirname(target), exist_ok=True)
            shutil.copyfile(os.path.join(F.ROOT, name), target)
        self.record_path = os.path.join(self.study, "source_binding.json")
        self.assertEqual(self.stale(), [])

    def tearDown(self):
        shutil.rmtree(self.root)

    def stale(self):
        return F.stale(F.load_record(self.record_path), self.root)

    def replace(self, name, old, new):
        path = os.path.join(self.root, name)
        with open(path, "rb") as f:
            data = f.read()
        self.assertEqual(data.count(old.encode()), 1, f"mutation anchor {old!r} in {name}")
        with open(path, "wb") as f:
            f.write(data.replace(old.encode(), new.encode()))


class Mutations(Copy):
    def assert_names(self, problems, *fragments):
        for fragment in fragments:
            self.assertTrue(any(fragment in p for p in problems), (fragment, problems))

    def test_hart_coefficient(self):
        self.replace(F.SOURCE, "SumA = SumA * Z + 220.206867912376", "SumA = SumA * Z + 220.206867912377")
        problems = self.stale()
        self.assert_names(problems, "procedure PROB_NormalUpperTailPositive", "Hart SumA coefficients")
        self.assertEqual(len(problems), 2, problems)

    def test_hart_coefficient_above_observed_range(self):
        self.replace(F.SOURCE, "8.83883476483184E-02 * Z", "8.83883476483185E-02 * Z")
        self.assert_names(self.stale(), "procedure PROB_NormalUpperTailPositive", "Hart SumB coefficients")

    def test_split(self):
        self.replace(F.SOURCE, "PROB_CDF_SPLIT As Double = 7.07106781186547", "PROB_CDF_SPLIT As Double = 7.5")
        problems = self.stale()
        self.assert_names(problems, "constant PROB_CDF_SPLIT:", "constant PROB_CDF_SPLIT = 7.5 in VBA")
        self.assertEqual(len(problems), 2, problems)

    def test_continued_fraction_depth(self):
        self.replace(F.SOURCE, "PROB_NORMAL_TAIL_CF_TERMS As Long = 16", "PROB_NORMAL_TAIL_CF_TERMS As Long = 24")
        problems = self.stale()
        self.assert_names(problems, "constant PROB_NORMAL_TAIL_CF_TERMS:",
                          "constant PROB_NORMAL_TAIL_CF_TERMS = 24.0 in VBA")
        self.assertEqual(len(problems), 2, problems)

    def test_route_predicate(self):
        self.replace(F.SOURCE, "If Z < PROB_CDF_SPLIT Then", "If Z <= PROB_CDF_SPLIT Then")
        self.assert_names(self.stale(), "procedure PROB_NormalUpperTailPositive",
                          "no longer contains the statement 'If Z < PROB_CDF_SPLIT Then'")
        self.replace(F.SOURCE, "If Z <= PROB_CDF_SPLIT Then", "If Z < PROB_CDF_SPLIT Then")
        self.replace(F.SOURCE, "PROB_NormalSurvival = 1# - PROB_NormalUpperTailPositive(-Z)",
                     "PROB_NormalSurvival = 1# - PROB_NormalUpperTailPositive(Z)")
        self.assert_names(self.stale(), "procedure PROB_NormalSurvival", "no longer contains the statement")

    def test_density(self):
        self.replace(F.SOURCE, "PROB_NormalPDF = PROB_INV_SQRT_TWO_PI * Exp(-0.5 * Z * Z)",
                     "PROB_NormalPDF = Exp(-0.5 * Z * Z) / PROB_SQRT_TWO_PI")
        self.assert_names(self.stale(), "procedure PROB_NormalPDF", "PROB_NormalPDF no longer contains")

    def test_comment_only_edit_fails_closed(self):
        self.replace(F.SOURCE, "'Use the Hart rational approximation below the tail split",
                     "'Use the Hart rational approximation below the split")
        self.assertEqual(len(self.stale()), 1)
        self.assert_names(self.stale(), "procedure PROB_NormalUpperTailPositive")

    def test_model_edit(self):
        self.replace(F.MODEL, "NORMAL_TAIL_CF_TERMS = 16", "NORMAL_TAIL_CF_TERMS = 17")
        self.assert_names(self.stale(), F.MODEL + ": recorded", "the model's NORMAL_TAIL_CF_TERMS = 17.0")
        self.replace(F.MODEL, "NORMAL_TAIL_CF_TERMS = 17", "NORMAL_TAIL_CF_TERMS = 16")
        self.replace(F.MODEL, "denominator = z + float(k) / denominator", "denominator = z + k / denominator")
        self.assert_names(self.stale(), "model vba_upper_tail no longer contains 'z + float(k) / denominator'")

    def test_missing_or_duplicated_procedure(self):
        self.replace(F.SOURCE, "Private Function PROB_NormalCDF(", "Private Function PROB_NormalCDF2(")
        self.assert_names(self.stale(), "cannot extract", "procedure PROB_NormalCDF: found 0")
        self.replace(F.SOURCE, "Private Function PROB_NormalCDF2(", "Private Function PROB_NormalCDF(")
        path = os.path.join(self.root, F.SOURCE)
        with open(path, "ab") as f:
            f.write(b"Private Function PROB_NormalSurvival(ByVal Z As Double) As Double\r\nEnd Function\r\n")
        self.assert_names(self.stale(), "procedure PROB_NormalSurvival: found 2")

    def test_missing_or_altered_record(self):
        os.unlink(self.record_path)
        self.assert_names(self.stale(), "source_binding.json is missing")
        self.assertNotEqual(F.stale({}, self.root), [])
        record = F.load_record()
        record["procedures"]["PROB_Extra"] = "0" * 64
        self.assert_names(F.stale(record, self.root), "PROB_Extra: recorded but no longer bound")

    def test_line_endings_bind_identically(self):
        path = os.path.join(self.root, F.SOURCE)
        with open(path, "rb") as f:
            data = f.read()
        self.assertIn(b"\r\n", data)
        with open(path, "wb") as f:
            f.write(data.replace(b"\r\n", b"\n"))
        self.assertEqual(self.stale(), [])


class Record(Copy):
    def test_refuses_while_the_model_differs(self):
        with open(self.record_path, "rb") as f:
            before = f.read()
        self.replace(F.SOURCE, "PROB_CDF_SPLIT As Double = 7.07106781186547", "PROB_CDF_SPLIT As Double = 7.5")
        problems = F.record("synthetic", self.root, self.record_path, "2026-11-05")
        self.assertTrue(any("PROB_CDF_SPLIT" in p for p in problems), problems)
        with open(self.record_path, "rb") as f:
            self.assertEqual(f.read(), before)

    def test_records_a_consistent_change(self):
        self.replace(F.SOURCE, "PROB_NORMAL_TAIL_CF_TERMS As Long = 16", "PROB_NORMAL_TAIL_CF_TERMS As Long = 24")
        self.replace(F.MODEL, "NORMAL_TAIL_CF_TERMS = 16", "NORMAL_TAIL_CF_TERMS = 24")
        problems = self.stale()
        self.assertEqual(len(problems), 2, problems)
        self.assertEqual(F.record("synthetic: depth 24", self.root, self.record_path, "2026-11-05"), [])
        self.assertEqual(self.stale(), [])
        with open(self.record_path, encoding="utf-8") as f:
            entry = json.load(f)["history"][-1]
        self.assertEqual(entry["reason"], "synthetic: depth 24")
        self.assertEqual(len(entry["identities_changed"]), 2)
        self.assertEqual(len(F.load_record(self.record_path)["history"]), len(F.load_record()["history"]) + 1)


class FindingFixtureFailsClosed(Copy):
    def run_fixture(self):
        return subprocess.run([sys.executable, os.path.join(self.study, "test_binary64_route_model.py")],
                              capture_output=True, text=True, cwd=self.study)

    def test_stale_source_fails_the_f1_fixture(self):
        proc = self.run_fixture()
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.replace(F.SOURCE, "PROB_NORMAL_TAIL_CF_TERMS As Long = 16", "PROB_NORMAL_TAIL_CF_TERMS As Long = 24")
        proc = self.run_fixture()
        self.assertNotEqual(proc.returncode, 0)
        self.assertIn("source binding: " + F.SOURCE + " constant PROB_NORMAL_TAIL_CF_TERMS", proc.stdout)


if __name__ == "__main__":
    unittest.main(verbosity=1)
