"""The Discrete sheet (Binomial and Poisson counts): its cases, its references and its plan.

The references of examples/demo/discrete-cases.json are made with mpmath. Here each one is checked again by
a different method: exact rational arithmetic for the Binomial (the probability is the exact value of the
Double the sheet holds), a Decimal series for the Poisson, and scipy where it is installed.
"""
import json
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
import discrete_plan as D
import discrete_references as R
import explorer_plan
import parameters_plan
import probabilities_plan
import tails_plan

PLAN = D.plan()
DATA = D.load()
CASES = {item["name"]: item for item in DATA["cases"]}
EXCEL_LIMIT = 8192
MODULE = (ROOT / "src/workbook/M_STATS_PROBDIST_DISCRETESHEET.bas").read_text(encoding="utf-8")
PUBLIC = set(re.findall(r"Public Function (K_STATS_\w+)",
                        "\n".join(path.read_text(encoding="utf-8", errors="replace")
                                  for path in (ROOT / "src/modules").glob("*.bas"))))


def binomial_masses(case):
    """Every mass of the Binomial as an exact fraction, from the exact value of the Double p = a / den:
    integer numerators C(n, j) a^j (den - a)^(n - j) over the common denominator den^n, so the sums stay
    integer sums. Each numerator follows from the one before by an exact integer division."""
    n = case["inputs"]["trials"]
    p = Fraction(float(case["inputs"]["probability"]))
    a, b = p.numerator, p.denominator - p.numerator
    numerators = [b ** n]
    for j in range(n):
        numerators.append(numerators[-1] * (n - j) * a // ((j + 1) * b))
    return numerators, p.denominator ** n


def prefix(values):
    """Running sums: prefix[j] is the sum of values[0..j]."""
    total, sums = 0, []
    for value in values:
        total += value
        sums.append(total)
    return sums


def poisson_masses(case, last):
    """The Poisson masses 0..last as Decimals with 60 digits, from the Double mean rate x exposure."""
    getcontext().prec = 60
    mean = Decimal(float(case["inputs"]["rate"]) * float(case["inputs"]["exposure"]))
    term = (-mean).exp()
    masses = [term]
    for j in range(1, last + 1):
        term = term * mean / j
        masses.append(term)
    return masses


def close(value, expected, tolerance=1e-14):
    expected = float(expected)
    if expected == 0:
        return value == 0
    return abs(value - expected) <= tolerance * abs(expected)


class CaseTests(unittest.TestCase):
    def test_the_cases_are_the_documented_ones(self):
        self.assertEqual(len(DATA["cases"]), 9)
        self.assertEqual({item["model"] for item in DATA["cases"]}, set(D.MODELS))
        self.assertEqual(DATA["default_case"], "Loan defaults, Binomial")
        default = CASES[DATA["default_case"]]["inputs"]
        self.assertEqual((default["trials"], default["probability"]), (100, 0.02))
        # The Poisson inputs of the default give the Binomial's mean: the same portfolio as a rate.
        self.assertEqual(default["rate"] * default["exposure"], default["trials"] * default["probability"])

    def test_every_case_is_valid_for_the_sheet(self):
        limits = D.LIMITS
        for item in DATA["cases"]:
            inputs = item["inputs"]
            self.assertTrue(1 <= inputs["trials"] <= limits["trials"] and inputs["trials"] == int(inputs["trials"]))
            self.assertTrue(0 < inputs["probability"] < 1 and 0 < inputs["level"] < 1, item["name"])
            self.assertTrue(0 < inputs["rate"] <= limits["rate"] and 0 < inputs["exposure"] <= limits["exposure"])
            self.assertTrue(limits["mean_min"] <= inputs["rate"] * inputs["exposure"] <= limits["mean_max"])
            self.assertEqual(inputs["threshold"], int(inputs["threshold"]))
            if item["model"] == "Binomial":
                self.assertLessEqual(inputs["threshold"], inputs["trials"], item["name"])
            window = item["window"]
            self.assertLessEqual(window["upper"] - window["lower"] + 1, D.MAX_ROWS)
            self.assertTrue(window["lower"] <= inputs["threshold"] <= window["upper"], item["name"])
            self.assertTrue(window["lower"] <= item["results"]["quantile"] <= window["upper"], item["name"])

    def test_the_boundaries_are_covered(self):
        thresholds = {item["name"]: item["inputs"]["threshold"] for item in DATA["cases"]}
        self.assertEqual(thresholds["Threshold zero"], 0)
        self.assertEqual(thresholds["Threshold at n"], CASES["Threshold at n"]["inputs"]["trials"])
        zero = CASES["Threshold zero"]["results"]
        self.assertEqual((zero["fewer_than"], zero["at_least"]), (0.0, 1.0))
        at_n = CASES["Threshold at n"]["results"]
        self.assertEqual((at_n["at_most"], at_n["more_than"]), (1.0, 0.0))
        # A step boundary: q lies just below the step F(1) = 0.75, so the quantile is 1.
        step = CASES["Step boundary"]
        self.assertEqual(step["results"]["quantile"], 1)
        self.assertEqual((step["results"]["before_quantile"], step["results"]["at_quantile"]), (0.25, 0.75))
        # The far quantile lies beyond mean + 5 standard deviations: the window widens to hold it.
        far = CASES["Far quantile"]
        mean, sd = R.model_mean_sd({"model": far["model"], **far["inputs"]})
        self.assertGreater(far["results"]["quantile"], mean + D.SPREAD * sd)

    def test_the_quantile_is_the_smallest_count_with_f_at_least_q(self):
        for item in DATA["cases"]:
            results = item["results"]
            q = item["inputs"]["level"]
            self.assertLess(results["before_quantile"], q, item["name"])
            self.assertGreaterEqual(results["at_quantile"], q, item["name"])

    def test_the_readings_of_a_threshold_are_consistent(self):
        for item in DATA["cases"]:
            results = item["results"]
            self.assertAlmostEqual(results["at_most"] + results["more_than"], 1.0, places=14)
            self.assertAlmostEqual(results["fewer_than"] + results["at_least"], 1.0, places=14)
            self.assertAlmostEqual(results["at_least"] - results["more_than"], results["mass"], places=14)
            self.assertAlmostEqual(results["at_most"] - results["fewer_than"], results["mass"], places=14)

    def test_the_provenance_is_recorded(self):
        for key in ("python", "mpmath"):
            self.assertRegex(DATA["environment"][key], r"^\d+\.\d+")
        self.assertEqual(DATA["window"], {"spread": D.SPREAD, "margin": R.MARGIN, "max_rows": D.MAX_ROWS})
        for item in DATA["cases"]:
            self.assertTrue(item["purpose"], item["name"])

    def test_the_stored_references_are_what_mpmath_returns_now(self):
        if R.mp is None:
            self.skipTest("mpmath is not installed")
        again = R.build()
        self.assertEqual(again["cases"], DATA["cases"])


class IndependentReferenceTests(unittest.TestCase):
    """Each reference again, by another method."""

    def test_binomial_by_exact_fractions(self):
        for item in DATA["cases"]:
            if item["model"] != "Binomial":
                continue
            numerators, whole = binomial_masses(item)
            below = prefix(numerators)
            total = below[-1]
            self.assertEqual(total, whole)

            def exact(numerator):
                return float(Fraction(numerator, whole))

            n = item["inputs"]["trials"]
            for row in item["table"]:
                j = row["k"]
                self.assertTrue(close(row["mass"], exact(numerators[j])), f'{item["name"]} {j}')
                self.assertTrue(close(row["at_most"], exact(below[j])), f'{item["name"]} {j}')
                self.assertTrue(close(row["more_than"], exact(total - below[j])), f'{item["name"]} {j}')
            k = item["inputs"]["threshold"]
            results = item["results"]
            self.assertTrue(close(results["at_least"], exact(total - (below[k - 1] if k else 0))), item["name"])
            self.assertTrue(close(results["fewer_than"], exact(below[k - 1] if k else 0)), item["name"])
            q = Fraction(float(item["inputs"]["level"]))
            quantile = next(j for j in range(n + 1) if below[j] * q.denominator >= q.numerator * whole)
            self.assertEqual(results["quantile"], quantile, item["name"])

    def test_poisson_by_a_decimal_series(self):
        for item in DATA["cases"]:
            if item["model"] != "Poisson":
                continue
            last = max(row["k"] for row in item["table"])
            masses = poisson_masses(item, last + 400)
            below = prefix(masses)
            for row in item["table"]:
                j = row["k"]
                above = below[-1] - below[j]        # the terms beyond last + 400 are below 1E-300
                self.assertTrue(close(row["mass"], masses[j]), f'{item["name"]} {j}')
                self.assertTrue(close(row["at_most"], below[j], 1e-13), f'{item["name"]} {j}')
                self.assertTrue(close(row["more_than"], above, 1e-13), f'{item["name"]} {j}')
            q = Decimal(float(item["inputs"]["level"]))
            quantile = next(j for j in range(len(below)) if below[j] >= q)
            self.assertEqual(item["results"]["quantile"], quantile, item["name"])

    def test_against_scipy(self):
        try:
            from scipy import stats
        except ImportError:
            self.skipTest("SciPy is not installed")
        for item in DATA["cases"]:
            inputs = item["inputs"]
            if item["model"] == "Binomial":
                law = stats.binom(inputs["trials"], inputs["probability"])
            else:
                law = stats.poisson(inputs["rate"] * inputs["exposure"])
            for row in item["table"]:
                self.assertTrue(close(row["mass"], law.pmf(row["k"]), 1e-9), f'{item["name"]} {row["k"]}')
                self.assertTrue(close(row["more_than"], law.sf(row["k"]), 1e-9), f'{item["name"]} {row["k"]}')
            self.assertEqual(item["results"]["quantile"], law.ppf(inputs["level"]), item["name"])


class PlanTests(unittest.TestCase):
    def formulas(self):
        return {item["address"]: item["formula"] for item in PLAN["formulas"]}

    def test_every_formula_and_name_fits_excels_limit(self):
        for name, size in D.formula_lengths(PLAN):
            self.assertLess(size, EXCEL_LIMIT, name)

    def test_cells_are_written_once(self):
        addresses = [item["address"] for item in PLAN["formulas"]] + [item["address"] for item in PLAN["texts"]]
        self.assertEqual(len(addresses), len(set(addresses)), "a cell is written twice")

    def test_the_sheet_names_are_its_own(self):
        others = set()
        for module in (explorer_plan, probabilities_plan, parameters_plan, comparison_plan, tails_plan):
            others |= {item["name"] for item in module.plan()["names"]}
        sheet = {item["name"] for item in PLAN["names"]}
        self.assertEqual(others & sheet, set())
        self.assertTrue(all(name.startswith("Discrete") for name in sheet))
        text = " ".join(item["formula"] for item in PLAN["formulas"]) + " " + " ".join(
            item["refers"] for item in PLAN["names"]) + " " + " ".join(
            item["x"] + " " + item["y"] for chart in PLAN["charts"] for item in chart["series"]) + " " + " ".join(
            PLAN["validation"]["custom"].values()) + " " + " ".join(
            value for value in PLAN["conditional"].values() if isinstance(value, str))
        used = set(re.findall(r"\bDiscrete[A-Z][A-Za-z]*\b", text))
        self.assertLessEqual(used, sheet, "a formula reads a name the plan does not define")

    def test_every_call_is_a_public_function_of_the_two_models(self):
        code = " ".join(self.formulas().values()) + " ".join(item["refers"] for item in PLAN["names"])
        calls = set(re.findall(r"K_STATS_\w+", re.sub(r'"[^"]*"', '""', code)))
        self.assertLessEqual(calls, PUBLIC)
        families = {name.split("_")[2] for name in calls}
        self.assertEqual(families, {"Binomial", "Poisson", "NormalStandard"})

    def test_the_formula_shown_beside_each_result_is_the_call_it_makes(self):
        formulas = self.formulas()
        for key in ("mean", "sd", "mass", "at_most", "fewer_than", "at_least", "more_than", "quantile",
                    "before_quantile", "at_quantile"):
            row = D.RESULT_ROWS[key]
            value, shown = formulas[f"D{row}"], formulas[f"F{row}"]
            for model in D.MODELS:
                text = re.search(rf'"=(K_STATS_{model}_\w+)\(([^)]*)\)"', shown)
                self.assertIsNotNone(text, f"{key} {model}")
                name, arguments = text.group(1), text.group(2)
                self.assertIn(name + "(", value, f"{key} {model}")
                # The cells the text names, with the dollar signs of the formula.
                absolute = re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", arguments)
                self.assertIn(f"{name}({absolute})", value, f"{key} {model}")

    def test_a_survival_value_is_never_one_minus_the_cdf(self):
        code = re.sub(r'"[^"]*"', '""', " ".join(self.formulas().values()))
        self.assertNotRegex(code, r"1-K_STATS_")

    def test_the_threshold_readings_translate_the_inequalities(self):
        formulas = self.formulas()
        rows = D.RESULT_ROWS
        self.assertIn("IF($D$11=0,0,", formulas[f"D{rows['fewer_than']}"])
        self.assertIn("Cumulative($D$11-1,", formulas[f"D{rows['fewer_than']}"])
        self.assertIn("IF($D$11=0,1,", formulas[f"D{rows['at_least']}"])
        self.assertIn("Survival($D$11-1,", formulas[f"D{rows['at_least']}"])
        self.assertIn("Survival($D$11,", formulas[f"D{rows['more_than']}"])
        self.assertIn("Cumulative($D$11,", formulas[f"D{rows['at_most']}"])
        check = formulas[f"D{rows['check']}"]
        self.assertIn("<$D$12,$D$12<=", check)
        self.assertIn("is the smallest k with F(k) >= q", check)

    def test_the_window_is_the_one_of_the_references(self):
        names = {item["name"]: item["refers"] for item in PLAN["names"]}
        self.assertIn(f"{D.SPREAD}*DiscreteSpread", names["DiscreteLower"])
        self.assertIn(f"{D.SPREAD}*DiscreteSpread", names["DiscreteUpper"])
        self.assertIn(f">{D.MAX_ROWS}", names["DiscreteTooWide"])
        self.assertEqual(R.SPREAD, D.SPREAD)
        self.assertEqual(R.MAX_ROWS, D.MAX_ROWS)

    def test_lambda_parameters_are_not_function_or_cell_names(self):
        formulas = " ".join(self.formulas().values())
        for parameter in re.findall(r"LAMBDA\(([A-Za-z0-9_,]+),", formulas):
            for name in parameter.split(","):
                self.assertNotRegex(name, r"^[A-Za-z]{1,3}\d+$|^[rRcC]$", name)
                self.assertNotIn(name.upper(), {"VALUE", "COUNT", "ROW", "COLUMN", "K"})

    def test_nothing_reads_as_a_formula_unless_it_is_one(self):
        for item in PLAN["texts"]:
            self.assertNotIn(item["value"][:1], "=+-@", item["address"])

    def test_no_cell_needs_formulatext_or_python(self):
        code = re.sub(r'"[^"]*"', '""', " ".join(self.formulas().values()))
        self.assertNotIn("FORMULATEXT(", code)
        self.assertNotIn("PY(", code)

    def test_conditional_formats_need_no_list_separator(self):
        # FormatConditions.Add reads its formula with the regional list separator: none may hold a comma.
        for key, value in PLAN["conditional"].items():
            self.assertNotIn(",", value, key)
        for value in PLAN["validation"]["custom"].values():
            self.assertRegex(value, r"^=Discrete[A-Za-z]+$")

    def test_the_layout_is_the_documented_one(self):
        self.assertEqual(PLAN["layout"]["inputs"], "D6:D12")
        self.assertEqual(PLAN["layout"]["first_row"], 38)
        self.assertEqual(PLAN["layout"]["last_row"], 37 + D.MAX_ROWS)
        self.assertTrue(all(21 <= row <= 35 for row in D.RESULT_ROWS.values()))
        self.assertEqual(PLAN["layout"]["chart_area"], "K5:O35")
        self.assertEqual(len(PLAN["charts"]), 2)
        self.assertEqual(json.loads(json.dumps(PLAN))["schema"], "discrete-plan-1")

    def test_the_module_defaults_match_the_data(self):
        defaults = PLAN["defaults"]
        self.assertIn(f'DEFAULT_MODEL As String = "{defaults["model"]}"', MODULE)
        for key, constant in (("trials", "TRIALS"), ("probability", "PROBABILITY"), ("rate", "RATE"),
                              ("exposure", "EXPOSURE"), ("threshold", "THRESHOLD"), ("level", "LEVEL")):
            self.assertIn(f"DEFAULT_{constant} As Double = {defaults[key]}", MODULE)
        self.assertIn(f'BUILD_CELL As String = "D{D.RESULT_ROWS["excel_build"]}"', MODULE)
        self.assertIn(f'QUANTILE_CELL As String = "D{D.RESULT_ROWS["quantile"]}"', MODULE)
        for chart in PLAN["charts"]:
            self.assertIn(chart["name"], MODULE)


if __name__ == "__main__":
    unittest.main()
