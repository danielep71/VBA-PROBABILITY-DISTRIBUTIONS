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
     procedures and constants. Its statement inventory covers every
     executable statement of the bound procedures, in order, each either
     cross-checked or unchecked with a reason; the translation of the checked
     ones equals the model.
  4. Mutation fixtures, on a copy of the source and model: a Hart
     coefficient, the 7.07 split, the continued-fraction depth, a route
     predicate, the density, a comment-only edit and a model edit each make
     the binding stale, naming what changed; a missing or duplicated
     procedure and a missing record fail closed; CRLF and LF checkouts bind
     identically.
  5. Statement inventory: a new statement in neither group, a removed or
     reordered statement and an unchecked entry without a reason are
     reported. With the inventory updated to accept a VBA change but the
     model left as it is, the translation still reports the difference for a
     Hart sign flip (in VBA, or as a Python unary sign in the model), the
     continued-fraction initializer Denominator = 0#, a loop bound, the
     density, and unsupported syntax; a consistent change of both passes.
     The density's unchecked guard is harmless only while its constant
     exceeds the modelled T_MAX, and lowering it is reported.
  6. Re-recording: --record refuses while the model does not mirror the
     source, including for a statement in neither inventory group and for
     each translation difference above, and leaves the record untouched;
     after a consistent change of source and model it records the changed
     identities in the history, and the binding is current again.
  7. The F1 fixture itself (test_binary64_route_model.py) fails on a copy
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
from unittest import mock

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
        self.assertEqual(list(record["constants"]), list(F.BOUND_CONSTANTS))
        self.assertEqual((record["source"], record["model"]), (F.SOURCE, F.MODEL))
        self.assertTrue(record["history"])
        self.assertEqual(F.extraction_problems(), [])

    def test_inventory_covers_every_statement(self):
        source = F.read_text(os.path.join(F.ROOT, F.SOURCE))
        self.assertEqual(list(F.INVENTORY), list(F.PROCEDURES))
        for name in F.PROCEDURES:
            self.assertEqual([text for text, _ in F.INVENTORY[name]],
                             F.statements(F.procedure_text(source, name)), name)
            for text, reason in F.INVENTORY[name]:
                self.assertTrue(reason is F.CHECKED or (isinstance(reason, str) and reason.strip()), text)
        self.assertEqual(F.inventory_counts(), (51, 4))
        unchecked = [text for rows in F.INVENTORY.values() for text, reason in rows if reason is not F.CHECKED]
        self.assertEqual(unchecked, ["If Abs(Z) >= PROB_PDF_UNDERFLOW_Z Then", "PROB_NormalPDF = 0#", "Else",
                                     "End If"])

    def test_translation_reads_as_the_model(self):
        checked = [text for text, reason in F.INVENTORY["PROB_NormalCDF"] if reason is F.CHECKED]
        self.assertEqual(F._render(F.normalize(F.translate("PROB_NormalCDF", checked, "z"))),
                         "return vba_upper_tail(-z) if z < 0.0 else 1.0 - vba_upper_tail(z)")
        checked = [text for text, reason in F.INVENTORY["PROB_NormalUpperTailPositive"] if reason is F.CHECKED]
        text = F._render(F.normalize(F.translate("PROB_NormalUpperTailPositive", checked, "z")))
        for line in ("a = a * z + 220.206867912376", "denominator = z\n",
                     "for k in range(NORMAL_TAIL_CF_TERMS, 0.0, -1.0):"):
            self.assertIn(line, text)


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

    def assert_names(self, problems, *fragments):
        for fragment in fragments:
            self.assertTrue(any(fragment in p for p in problems), (fragment, problems))

    def accepted(self, procedure, old, new, reason=F.CHECKED):
        """INVENTORY as a reviewer would update it to accept a statement change (new=None: removal)."""
        rows = []
        for text, entry in F.INVENTORY[procedure]:
            if text == old:
                if new is not None:
                    rows.append((new, reason))
            else:
                rows.append((text, entry))
        self.assertNotEqual(tuple(rows), F.INVENTORY[procedure], f"{old!r} not in the {procedure} inventory")
        return mock.patch.dict(F.INVENTORY, {procedure: tuple(rows)})

    def assert_refused(self):
        with open(self.record_path, "rb") as f:
            before = f.read()
        problems = F.record("synthetic", self.root, self.record_path, "2026-11-05")
        self.assertNotEqual(problems, [])
        with open(self.record_path, "rb") as f:
            self.assertEqual(f.read(), before)
        return problems


UPPER = "PROB_NormalUpperTailPositive"
UPPER_MISMATCH = "PROB_NormalUpperTailPositive does not match model vba_upper_tail"


class Mutations(Copy):
    def test_hart_coefficient(self):
        old, new = "SumA = SumA * Z + 220.206867912376", "SumA = SumA * Z + 220.206867912377"
        self.replace(F.SOURCE, old, new)
        problems = self.stale()
        self.assert_names(problems, "procedure PROB_NormalUpperTailPositive", f"statement '{new}' is neither",
                          f"INVENTORY lists '{old}'")
        self.assertEqual(len(problems), 3, problems)
        with self.accepted(UPPER, old, new):
            self.assert_names(self.stale(), UPPER_MISMATCH)

    def test_hart_coefficient_above_observed_range(self):
        old = "SumB = 8.83883476483184E-02 * Z + 1.75566716318264"
        new = "SumB = 8.83883476483185E-02 * Z + 1.75566716318264"
        self.replace(F.SOURCE, old, new)
        self.assert_names(self.stale(), "procedure PROB_NormalUpperTailPositive", f"statement '{new}' is neither")
        with self.accepted(UPPER, old, new):
            self.assert_names(self.stale(), UPPER_MISMATCH)

    def test_split(self):
        self.replace(F.SOURCE, "PROB_CDF_SPLIT As Double = 7.07106781186547", "PROB_CDF_SPLIT As Double = 7.5")
        problems = self.stale()
        self.assert_names(problems, "constant PROB_CDF_SPLIT:", "constant PROB_CDF_SPLIT = 7.5 in VBA")
        self.assertEqual(len(problems), 2, problems)

    def test_continued_fraction_depth(self):
        self.replace(F.SOURCE, "PROB_NORMAL_TAIL_CF_TERMS As Long = 16", "PROB_NORMAL_TAIL_CF_TERMS As Long = 24")
        problems = self.stale()
        self.assert_names(problems, "constant PROB_NORMAL_TAIL_CF_TERMS:",
                          "constant PROB_NORMAL_TAIL_CF_TERMS = 24 in VBA")
        self.assertEqual(len(problems), 2, problems)

    def test_route_predicate(self):
        old, new = "If Z < PROB_CDF_SPLIT Then", "If Z <= PROB_CDF_SPLIT Then"
        self.replace(F.SOURCE, old, new)
        self.assert_names(self.stale(), "procedure PROB_NormalUpperTailPositive", f"statement '{new}' is neither")
        with self.accepted(UPPER, old, new):
            self.assert_names(self.stale(), UPPER_MISMATCH)
        self.replace(F.SOURCE, new, old)
        old, new = ("PROB_NormalSurvival = 1# - PROB_NormalUpperTailPositive(-Z)",
                    "PROB_NormalSurvival = 1# - PROB_NormalUpperTailPositive(Z)")
        self.replace(F.SOURCE, old, new)
        self.assert_names(self.stale(), "procedure PROB_NormalSurvival", f"statement '{new}' is neither")
        with self.accepted("PROB_NormalSurvival", old, new):
            self.assert_names(self.stale(), "PROB_NormalSurvival does not match model vba_survival")

    def test_density(self):
        old = "PROB_NormalPDF = PROB_INV_SQRT_TWO_PI * Exp(-0.5 * Z * Z)"
        new = "PROB_NormalPDF = PROB_INV_SQRT_TWO_PI * Exp(0.5 * Z * Z)"
        self.replace(F.SOURCE, old, new)
        self.assert_names(self.stale(), "procedure PROB_NormalPDF", f"statement '{new}' is neither")
        with self.accepted("PROB_NormalPDF", old, new):
            self.assert_names(self.stale(), "PROB_NormalPDF does not match phi in model route")

    def test_comment_only_edit_fails_closed(self):
        self.replace(F.SOURCE, "'Use the Hart rational approximation below the tail split",
                     "'Use the Hart rational approximation below the split")
        self.assertEqual(len(self.stale()), 1)
        self.assert_names(self.stale(), "procedure PROB_NormalUpperTailPositive")

    def test_model_edit(self):
        self.replace(F.MODEL, "NORMAL_TAIL_CF_TERMS = 16", "NORMAL_TAIL_CF_TERMS = 17")
        self.assert_names(self.stale(), F.MODEL + ": recorded", "the model's NORMAL_TAIL_CF_TERMS = 17")
        self.replace(F.MODEL, "NORMAL_TAIL_CF_TERMS = 17", "NORMAL_TAIL_CF_TERMS = 16")
        self.replace(F.MODEL, "denominator = z + float(k) / denominator", "denominator = z + k / denominator")
        self.assert_names(self.stale(), UPPER_MISMATCH)
        self.replace(F.MODEL, "denominator = z + k / denominator", "denominator = z + float(k) / denominator")
        self.replace(F.MODEL, "phi = INV_SQRT_TWO_PI * math.exp(-0.5 * t * t)",
                     "phi = INV_SQRT_TWO_PI * math.exp(-0.5 * t)")
        self.assert_names(self.stale(), "PROB_NormalPDF does not match phi in model route")

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


class Inventory(Copy):
    """Every statement must be cross-checked or explicitly unchecked; changes force a decision."""

    def test_statement_in_neither_group(self):
        self.replace(F.SOURCE, "Exponential = Exp(-0.5 * Z * Z)\r\n",
                     "Exponential = Exp(-0.5 * Z * Z)\r\n            Exponential = Exponential * 1#\r\n")
        problems = self.stale()
        self.assert_names(problems, "PROB_NormalUpperTailPositive: statement 'Exponential = Exponential * 1#' is "
                                    "neither cross-checked nor listed as intentionally unchecked")
        self.assert_names(self.assert_refused(), "'Exponential = Exponential * 1#' is neither")

    def test_new_statement_in_a_partly_unchecked_procedure(self):
        self.replace(F.SOURCE, "If Abs(Z) >= PROB_PDF_UNDERFLOW_Z Then", "If Abs(Z) >= 5# Then")
        self.assert_names(self.stale(), "PROB_NormalPDF: statement 'If Abs(Z) >= 5# Then' is neither")
        self.assert_refused()

    def test_removed_and_reordered_statements(self):
        self.replace(F.SOURCE, "PROB_NormalUpperTailPositive = 0#\r\n            Exit Function\r\n",
                     "PROB_NormalUpperTailPositive = 0#\r\n")
        self.assert_names(self.stale(), "INVENTORY lists 'Exit Function', which the source no longer contains")
        with self.accepted(UPPER, "Exit Function", None):
            self.assert_names(self.stale(), UPPER_MISMATCH)
            self.assert_refused()
        self.replace(F.SOURCE, "PROB_NormalUpperTailPositive = 0#\r\n",
                     "PROB_NormalUpperTailPositive = 0#\r\n            Exit Function\r\n")
        self.replace(F.SOURCE, "SumA = SumA * Z + 6.37396220353165\r\n            SumA = SumA * Z + 33.912866078383",
                     "SumA = SumA * Z + 33.912866078383\r\n            SumA = SumA * Z + 6.37396220353165")
        self.assert_names(self.stale(), "PROB_NormalUpperTailPositive: the statements are in a different order")

    def test_unchecked_needs_a_reason(self):
        with self.accepted("PROB_NormalPDF", "PROB_NormalPDF = 0#", "PROB_NormalPDF = 0#", reason=" "):
            self.assert_names(self.stale(), "PROB_NormalPDF: 'PROB_NormalPDF = 0#' is unchecked without a reason")

    def test_guard_must_stay_outside_the_modelled_domain(self):
        self.replace(F.SOURCE, "PROB_PDF_UNDERFLOW_Z As Double = 38.6", "PROB_PDF_UNDERFLOW_Z As Double = 5")
        problems = self.stale()
        self.assert_names(problems, "constant PROB_PDF_UNDERFLOW_Z:",
                          "guard constant PROB_PDF_UNDERFLOW_Z = 5.0 must exceed the model's T_MAX = 8.0")
        self.assert_refused()


class Translation(Copy):
    """With INVENTORY updated to accept a VBA change, the model must still mirror it exactly."""

    def accept_vba(self, procedure, old, new):
        self.replace(F.SOURCE, old, new)
        return self.accepted(procedure, old, new)

    def test_hart_sign_flip(self):
        with self.accept_vba(UPPER, "SumA = SumA * Z + 220.206867912376", "SumA = SumA * Z - 220.206867912376"):
            problems = self.stale()
            self.assert_names(problems, UPPER_MISMATCH)
            self.assert_names(problems, "a = a * z - 220.206867912376", "a = a * z + 220.206867912376")
            self.assert_refused()

    def test_hart_sign_as_a_python_unary_sign(self):
        self.replace(F.MODEL, "a = a * z + 220.206867912376", "a = a * z + -220.206867912376")
        self.assert_names(self.stale(), F.MODEL + ": recorded", UPPER_MISMATCH)
        self.assert_refused()
        self.replace(F.MODEL, "a = a * z + -220.206867912376", "a = a * z + 220.206867912376")
        with self.accept_vba(UPPER, "SumB = 8.83883476483184E-02 * Z + 1.75566716318264",
                             "SumB = -8.83883476483184E-02 * Z + 1.75566716318264"):
            self.assert_names(self.stale(), UPPER_MISMATCH)
            self.assert_refused()

    def test_consistent_sign_flip_passes(self):
        with self.accept_vba(UPPER, "SumA = SumA * Z + 220.206867912376", "SumA = SumA * Z - 220.206867912376"):
            self.replace(F.MODEL, "a = a * z + 220.206867912376", "a = a * z - 220.206867912376")
            self.assertEqual(F.extraction_problems(self.root), [])
            self.assertEqual(len(self.stale()), 2, "only the procedure and model digests differ")

    def test_continued_fraction_initializer(self):
        self.replace(F.SOURCE, "Denominator = Z\r\n", "Denominator = 0#\r\n")
        self.assert_names(self.stale(), "statement 'Denominator = 0#' is neither")
        with self.accepted(UPPER, "Denominator = Z", "Denominator = 0#"):
            problems = self.stale()
            self.assert_names(problems, UPPER_MISMATCH, "denominator = 0.0")
            self.assert_refused()
        self.replace(F.SOURCE, "Denominator = 0#\r\n", "Denominator = Z\r\n")
        self.replace(F.MODEL, "    denominator = z\n", "    denominator = 0.0\n")
        self.assert_names(self.stale(), UPPER_MISMATCH)

    def test_loop_bound_and_step(self):
        old = "For TermIndex = PROB_NORMAL_TAIL_CF_TERMS To 1 Step -1"
        with self.accept_vba(UPPER, old, "For TermIndex = PROB_NORMAL_TAIL_CF_TERMS To 0 Step -1"):
            self.assert_names(self.stale(), UPPER_MISMATCH, "range(NORMAL_TAIL_CF_TERMS, -1.0, -1.0)")
        self.replace(F.SOURCE, "For TermIndex = PROB_NORMAL_TAIL_CF_TERMS To 0 Step -1", old)
        with self.accept_vba(UPPER, old, "For TermIndex = PROB_NORMAL_TAIL_CF_TERMS To 1 Step -2"):
            self.assert_names(self.stale(), UPPER_MISMATCH)

    def test_unsupported_syntax_and_names_fail_closed(self):
        for old, new, fragment in (
                ("SumA = SumA * Z + 220.206867912376", "SumA = SumA * Z ^ 1 + 220.206867912376",
                 "unsupported VBA syntax"),
                ("Denominator = Z\r\n", "Denominator = Abs(Z)\r\n", "identifier Abs has no model counterpart"),
                ("Dim SumA                As Double", "Dim SumA                As Single",
                 "unsupported statement 'Dim SumA As Single'")):
            with self.subTest(new=new):
                self.replace(F.SOURCE, old, new)
                with self.accepted(UPPER, " ".join(old.split()), " ".join(new.split())):
                    self.assert_names(self.stale(), "cannot translate the checked statements", fragment)
                    self.assert_refused()
                self.replace(F.SOURCE, new, old)


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
