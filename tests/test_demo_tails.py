"""The Tails sheet (Extreme Tails and Numerical Limits): its scenarios, its references and its plan.

The references of tests/../examples/demo/tails-cases.json are made with mpmath. Here each one is checked again
by a different method: exact integer arithmetic for the discrete families, a Decimal sum for the Poisson
tail, closed forms for the Exponential and the Weibull, math.erfc for the Normal where it does not underflow,
and scipy's logsf for the other continuous families where scipy is installed.
"""
import json
import math
import re
import sys
import unittest
from decimal import Decimal, getcontext
from fractions import Fraction
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools/demobuild"))
sys.path.insert(0, str(ROOT / "examples/demo"))
import comparison_plan
import explorer_plan
import parameters_plan
import probabilities_plan
import tails_plan as T
import tails_references as R

PLAN = T.plan()
DATA = T.load()
SCENARIOS = {item["name"]: item for item in DATA["scenarios"]}
EXCEL_LIMIT = 8192
MODULE = (ROOT / "src/workbook/M_STATS_PROBDIST_TAILS.bas").read_text(encoding="utf-8")
PUBLIC = set(re.findall(r"Public Function (K_STATS_\w+)",
                        "\n".join(path.read_text(encoding="utf-8", errors="replace")
                                  for path in (ROOT / "src/modules").glob("*.bas"))))
LN10 = math.log(10)


def depth_of(reference):
    """-log10 of a positive reference given as text, by Decimal arithmetic."""
    getcontext().prec = 50
    return float(-Decimal(reference).log10())


class ScenarioTests(unittest.TestCase):
    def test_the_scenarios_are_the_documented_ones(self):
        self.assertEqual(len(DATA["scenarios"]), 13)
        self.assertEqual(sum(item["kind"] == "tail" for item in DATA["scenarios"]), 8)
        self.assertEqual(sum(item["kind"] == "log" for item in DATA["scenarios"]), 5)
        self.assertEqual(DATA["default_scenario"], "Standard Normal")
        self.assertEqual((DATA["grades"]["reference"], DATA["grades"]["report"]), (12, 6))

    def test_every_call_is_a_public_function_and_the_alternative_is_the_expected_one(self):
        for item in DATA["scenarios"]:
            for key in ("direct", "alternative"):
                self.assertIn("{x}", item[key], item["name"])
            self.assertIn(item["direct"].split("(")[0], PUBLIC, item["name"])
            inner = re.findall(r"K_STATS_\w+", item["alternative"])
            self.assertEqual(len(inner), 1, item["name"])
            self.assertIn(inner[0], PUBLIC, item["name"])
            if item["kind"] == "tail":
                self.assertTrue(item["direct"].startswith(f"K_STATS_{inner[0].split('_')[2]}_"), item["name"])
                self.assertTrue(item["direct"].split("(")[0].endswith("_Survival"), item["name"])
                self.assertTrue(item["alternative"].startswith("1-") and inner[0].endswith("_Cumulative"), item["name"])
            else:
                self.assertTrue(item["direct"].split("(")[0].endswith("_LogPMF"), item["name"])
                self.assertTrue(item["alternative"].startswith("LN(") and inner[0].endswith("_PMF"), item["name"])

    def test_the_points_go_deeper_and_cross_the_limit_of_the_double_range(self):
        for item in DATA["scenarios"]:
            depths = [point["depth"] for point in item["points"]]
            self.assertEqual(depths, sorted(depths), item["name"])
            self.assertEqual(len({point["x"] for point in item["points"]}), len(item["points"]), item["name"])
            self.assertLess(depths[0], 12, f"{item['name']} starts in the body")
            self.assertGreater(depths[-1], 308, f"{item['name']} goes past 2.2E-308")
            beyond = [point for point in item["points"] if point["depth"] > 308]
            self.assertGreaterEqual(len(beyond), 2, item["name"])

    def test_a_reference_that_no_double_holds_is_stored_as_zero_and_as_text(self):
        for item in DATA["scenarios"]:
            for point in item["points"]:
                if item["kind"] == "tail":
                    value = float(point["reference"])
                    if point["depth"] > 308:
                        self.assertEqual(point["reference_value"], 0.0, f"{item['name']} {point['x']}")
                    else:
                        self.assertEqual(point["reference_value"], value, f"{item['name']} {point['x']}")
                    if value >= DATA["smallest_normal"]:
                        self.assertGreater(point["reference_value"], 0)
                    self.assertAlmostEqual(point["depth"], depth_of(point["reference"]), places=9)
                else:
                    self.assertEqual(point["reference_value"], float(point["reference"]))
                    self.assertAlmostEqual(point["depth"], -point["reference_value"] / LN10, places=9)

    def test_the_provenance_is_recorded(self):
        for key in ("python", "mpmath"):
            self.assertRegex(DATA["environment"][key], r"^\d+\.\d+")
        self.assertEqual(DATA["reference"]["digits"], 34)
        for item in DATA["scenarios"]:
            self.assertTrue(item["method"] and item["contract"] and item["parameters"] and item["title"], item["name"])
            for point in item["points"]:
                self.assertGreaterEqual(len(point["reference"].replace(".", "").replace("-", "").split("e")[0]), 30)

    def test_the_stored_references_are_what_mpmath_returns_now(self):
        if R.mp is None:
            self.skipTest("mpmath is not installed")
        again = R.build()
        for item, then in zip(again["scenarios"], DATA["scenarios"]):
            self.assertEqual(item["points"], then["points"], item["name"])


class IndependentReferenceTests(unittest.TestCase):
    """Each reference again, by another method."""

    def points(self, name):
        return SCENARIOS[name]["points"]

    def test_normal_tail_against_erfc_where_it_does_not_underflow(self):
        for point in self.points("Standard Normal"):
            if point["reference_value"] > 1e-290:
                expected = math.erfc(point["x"] / math.sqrt(2)) / 2
                # erfc of a rounded argument: the relative error grows with z squared.
                self.assertLess(abs(expected - point["reference_value"]) / expected, 4e-16 * point["x"] ** 2 + 1e-14, point["x"])

    def test_exponential_and_weibull_are_closed_forms(self):
        for point in self.points("Exponential"):
            self.assertAlmostEqual(point["depth"], point["x"] / LN10, places=9)
        for point in self.points("Weibull"):
            self.assertAlmostEqual(point["depth"], point["x"] ** 2 / LN10, places=9)

    def test_binomial_tail_by_exact_integers(self):
        for point in self.points("Binomial"):
            k = int(point["x"])
            total = sum(math.comb(2000, j) for j in range(k + 1, 2001))
            expected = 2000 * math.log10(2) - math.log10(total)
            self.assertAlmostEqual(point["depth"], expected, places=9, msg=str(k))

    def test_poisson_tail_by_a_decimal_sum(self):
        getcontext().prec = 60
        scale = (-Decimal(10)).exp()
        for point in self.points("Poisson"):
            k = int(point["x"])
            total = sum(Fraction(10 ** j, math.factorial(j)) for j in range(k + 1, k + 260))
            tail = Decimal(total.numerator) / Decimal(total.denominator) * scale
            self.assertAlmostEqual(point["depth"], float(-tail.log10()), places=9, msg=str(k))

    def test_log_mass_by_exact_integers(self):
        for point in self.points("Binomial LogPMF"):
            k = int(point["x"])
            expected = math.log(math.comb(2000, k)) - 2000 * math.log(2)
            self.assertAlmostEqual(point["reference_value"], expected, places=9, msg=str(k))
        for point in self.points("Poisson LogPMF"):
            k = int(point["x"])
            expected = k * math.log(10) - 10 - math.lgamma(k + 1)
            self.assertAlmostEqual(point["reference_value"], expected, places=7, msg=str(k))
        for point in self.points("Geometric LogPMF"):
            self.assertAlmostEqual(point["reference_value"], -(point["x"] + 1) * math.log(2), places=9)
        for point in self.points("Negative Binomial LogPMF"):
            f = int(point["x"])
            expected = math.log(math.comb(f + 4, f)) - (5 + f) * math.log(2)
            self.assertAlmostEqual(point["reference_value"], expected, places=9, msg=str(f))
        for point in self.points("Hypergeometric LogPMF"):
            k = int(point["x"])
            expected = math.log(math.comb(5000, k) * math.comb(5000, 2000 - k)) - math.log(math.comb(10000, 2000))
            self.assertAlmostEqual(point["reference_value"], expected, places=9, msg=str(k))

    def test_continuous_tails_against_scipy_logsf(self):
        try:
            from scipy import stats
        except ImportError:
            self.skipTest("SciPy is not installed")
        checks = {"Standard Normal": lambda x: stats.norm.logsf(x),
                  "Exponential": lambda x: stats.expon.logsf(x),
                  "Weibull": lambda x: stats.weibull_min.logsf(x, 2),
                  "Gamma": lambda x: stats.gamma.logsf(x, 5),
                  "Chi-square": lambda x: stats.chi2.logsf(x, 3),
                  "Student t": lambda x: stats.t.logsf(x, 5)}
        for name, logsf in checks.items():
            checked = 0
            for point in self.points(name):
                value = float(logsf(point["x"]))
                if not math.isfinite(value):
                    continue    # scipy's gamma survival gives -inf far in the tail: it is the reference's job
                self.assertAlmostEqual(point["depth"], -value / LN10, delta=1e-7, msg=f"{name} {point['x']}")
                checked += 1
            self.assertGreaterEqual(checked, len(self.points(name)) // 2, name)


class PlanTests(unittest.TestCase):
    def test_every_formula_and_name_fits_excels_limit(self):
        for name, size in T.formula_lengths(PLAN):
            self.assertLess(size, EXCEL_LIMIT, name)

    def test_cells_are_written_once(self):
        addresses = [item["address"] for item in PLAN["formulas"]] + [item["address"] for item in PLAN["texts"]]
        addresses += [item["address"] for item in PLAN["numbers"]]
        self.assertEqual(len(addresses), len(set(addresses)), "a cell is written twice")

    def test_the_stored_columns_have_one_value_per_row(self):
        for column in PLAN["columns"]:
            first, last = column["address"].split(":")
            rows = int(re.sub(r"\D", "", last)) - int(re.sub(r"\D", "", first)) + 1
            self.assertEqual(len(column["values"]), rows, column["address"])
        scenario = next(item for item in PLAN["columns"] if item["address"].startswith("R"))
        self.assertEqual(scenario["values"][0], "Standard Normal")
        self.assertEqual(len(scenario["values"]), sum(len(item["points"]) for item in DATA["scenarios"]))

    def test_the_sheet_names_are_its_own(self):
        others = set()
        for module in (explorer_plan, probabilities_plan, parameters_plan, comparison_plan):
            others |= {item["name"] for item in module.plan()["names"]}
        sheet = {item["name"] for item in PLAN["names"]}
        self.assertEqual(others & sheet, set())
        self.assertTrue(all(name.startswith("Tails") for name in sheet))
        text = " ".join(item["formula"] for item in PLAN["formulas"]) + " " + " ".join(
            item["refers"] for item in PLAN["names"]) + " " + " ".join(
            item["x"] + " " + item["y"] for chart in PLAN["charts"] for item in chart["series"])
        used = set(re.findall(r"\bTails[A-Z][A-Za-z]*\b", text))
        self.assertLessEqual(used, sheet, "a formula reads a name the plan does not define")

    def test_every_scenario_has_both_branches_in_both_switches(self):
        formulas = {item["address"]: item["formula"] for item in PLAN["formulas"]}
        for item in DATA["scenarios"]:
            direct = item["direct"].format(x="x")
            alternative = item["alternative"].format(x="x")
            self.assertIn(f'"{item["name"]}",IFERROR({direct},"error")', formulas["F38"])
            failure = "underflow" if item["kind"] == "log" else "error"
            self.assertIn(f'"{item["name"]}",IFERROR({alternative},"{failure}")', formulas["G38"])

    def test_lambda_parameters_are_not_function_or_cell_names(self):
        formulas = " ".join(item["formula"] for item in PLAN["formulas"])
        for parameter in re.findall(r"LAMBDA\(([A-Za-z0-9_,]+),", formulas):
            for name in parameter.split(","):
                self.assertNotRegex(name, r"^[A-Za-z]{1,3}\d+$|^[rRcC]$", name)
                self.assertNotIn(name.upper(), {"VALUE", "DIGITS", "ROW", "COLUMN"})

    def test_digits_and_statuses_follow_the_sheet_contract(self):
        formulas = {item["address"]: item["formula"] for item in PLAN["formulas"]}
        digits = formulas["H38"]
        self.assertIn("truth=0", digits)
        self.assertIn("-LOG10(", digits)
        self.assertIn(f"MIN({T.FULL_DIGITS}", digits)
        for address in ("K38", "L38"):
            for word in T.STATUS_WORDS.values():
                if address == "K38" and word in ("zero: cancellation",):
                    continue
                if address == "L38" and word in ("zero: unexpected",):
                    continue
                self.assertIn(f'"{word}"', formulas[address], f"{address} lacks {word}")
        self.assertIn(T.REFERENCE_GRADE, formulas["K38"])
        self.assertIn(T.REPORT_GRADE, formulas["K38"])

    def test_the_two_chart_lines_of_the_limit_are_at_the_smallest_normal_double(self):
        numbers = {item["address"]: item["value"] for item in PLAN["numbers"]}
        depth = -math.log10(T.SMALLEST_NORMAL)
        self.assertAlmostEqual(numbers["AL38"], depth, places=12)
        self.assertAlmostEqual(numbers["AO38"], -depth, places=12)
        self.assertEqual(T.SMALLEST_NORMAL, DATA["smallest_normal"])

    def test_nothing_reads_as_a_formula_unless_it_is_one(self):
        for item in PLAN["texts"]:
            self.assertNotIn(item["value"][:1], "=+-@", item["address"])
        # The stored columns are written into cells formatted as text, so a call kept as text stays text.
        self.assertTrue(PLAN["layout"]["text_columns"])

    def test_the_layout_is_the_documented_one(self):
        self.assertEqual(PLAN["layout"]["inputs"], "D6:D8")
        self.assertEqual(PLAN["layout"]["first_row"], 38)
        self.assertTrue(all(21 <= row <= 35 for row in T.DETAIL_ROWS.values()))
        self.assertEqual(PLAN["layout"]["last_row"], 37 + max(len(item["points"]) for item in DATA["scenarios"]))
        self.assertLess(PLAN["layout"]["store_last_row"], 300)
        self.assertEqual(PLAN["layout"]["chart_area"], "K5:O35")
        self.assertEqual(len(PLAN["charts"]), 2)

    def test_the_module_defaults_match_the_data(self):
        self.assertIn(f'DEFAULT_SCENARIO As String = "{DATA["default_scenario"]}"', MODULE)
        self.assertIn(f'DEFAULT_REFERENCE_GRADE As Double = {DATA["grades"]["reference"]}', MODULE)
        self.assertIn(f'DEFAULT_REPORT_GRADE As Double = {DATA["grades"]["report"]}', MODULE)
        self.assertIn(f'BUILD_CELL As String = "D{T.DETAIL_ROWS["excel_build"]}"', MODULE)
        for chart in PLAN["charts"]:
            self.assertIn(chart["name"], MODULE)

    def test_no_cell_needs_formulatext_or_python(self):
        code = re.sub(r'"[^"]*"', '""', " ".join(item["formula"] for item in PLAN["formulas"]))
        self.assertNotIn("FORMULATEXT(", code)
        self.assertNotIn("PY(", code)

    def test_the_statuses_are_coloured_by_what_went_wrong(self):
        rules = {item["text"]: item for item in PLAN["conditional"]["status_rules"]}
        self.assertEqual(rules["zero: cancellation"]["fill"], rules["wrong"]["fill"])
        self.assertEqual(rules["zero: underflow"]["fill"], rules["PMF underflow"]["fill"])
        self.assertNotEqual(rules["wrong"]["fill"], rules["zero: underflow"]["fill"])
        self.assertNotIn("reference grade", rules)
        self.assertEqual(json.loads(json.dumps(PLAN))["schema"], "tails-plan-1")


if __name__ == "__main__":
    unittest.main()
