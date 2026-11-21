"""The Risk sheet (value at risk, expected shortfall, tails and default counts): its cases, its references and
its plan.

The references of examples/demo/risk-cases.json are made with mpmath from closed forms (the Normal and Student-t
quantiles, the expected shortfall formulas, the incomplete beta for the Student-t tail). Here each one is
checked again by a different method: the expected shortfall by numerical integration of the loss beyond the
VaR, the Student-t tail and quantile by integrating the density, the Binomial by a direct sum of masses, and
SciPy where it is installed.
"""
import json
import math
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools/demobuild"))
sys.path.insert(0, str(ROOT / "examples/demo"))
import comparison_plan
import discrete_plan
import explorer_plan
import parameters_plan
import probabilities_plan
import risk_plan as P
import risk_references as R
import tails_plan

PLAN = P.plan()
DATA = P.load()
CASES = {item["name"]: item for item in DATA["cases"]}
EXCEL_LIMIT = 8192
MODULE = (ROOT / "src/workbook/M_STATS_PROBDIST_RISK.bas").read_text(encoding="utf-8")
PUBLIC = set(re.findall(r"Public Function (K_STATS_\w+)",
                        "\n".join(path.read_text(encoding="utf-8", errors="replace")
                                  for path in (ROOT / "src/modules").glob("*.bas"))))

try:
    from mpmath import mp, mpf
except ImportError:
    mp = mpf = None


def close(value, expected, tolerance=1e-12, scale=0.0):
    size = max(abs(float(expected)), scale)
    if size == 0:
        return value == 0
    return abs(float(value) - float(expected)) <= tolerance * size


class CaseTests(unittest.TestCase):
    def test_the_cases_are_the_documented_ones(self):
        self.assertEqual(len(DATA["cases"]), 7)
        self.assertEqual(DATA["default_case"], "Default portfolio")
        default = CASES["Default portfolio"]["inputs"]
        self.assertEqual((default["value"], default["volatility"], default["horizon"], default["confidence"],
                          default["df"]), (1000000, 0.01, 10, 0.99, 4))
        self.assertEqual((default["loans"], default["default_probability"]), (100, 0.02))

    def test_every_case_is_valid_for_the_sheet(self):
        limits = P.LIMITS
        for item in DATA["cases"]:
            inputs = item["inputs"]
            self.assertTrue(0 < inputs["value"] <= limits["value"], item["name"])
            self.assertLessEqual(abs(inputs["mean_return"]), limits["mean_return"])
            self.assertTrue(0 < inputs["volatility"] <= limits["volatility"])
            self.assertTrue(1 <= inputs["horizon"] <= limits["horizon"] and inputs["horizon"] == int(inputs["horizon"]))
            self.assertTrue(limits["confidence_min"] <= inputs["confidence"] <= limits["confidence_max"])
            self.assertTrue(2 < inputs["df"] <= limits["df"], item["name"])
            self.assertTrue(1 <= inputs["loans"] <= limits["loans"] and inputs["loans"] == int(inputs["loans"]))
            self.assertTrue(0 < inputs["default_probability"] < 1)
            self.assertTrue(0 <= inputs["loss_per_default"] <= limits["loss_per_default"])
            self.assertEqual(len(item["table"]), P.ROWS)

    def test_the_cases_cover_what_the_sheet_explains(self):
        results = {name: item["results"] for name, item in CASES.items()}
        # Heavy tails: at 99.99% with nu = 3 the Student-t VaR is several times the Normal's.
        self.assertGreater(results["Rare tail"]["t_var"], 3 * results["Rare tail"]["normal_var"])
        # Close to the Normal when nu is large.
        thin = results["Thin tails"]
        self.assertLess(abs(thin["t_var"] / thin["normal_var"] - 1), 0.01)
        # At a moderate level the Student t at the same variance can give the lower VaR: its body is narrower.
        self.assertLess(results["Expected loss"]["t_var"], results["Expected loss"]["normal_var"])
        # A positive mean loss, and nu just above the variance limit.
        self.assertGreater(results["Expected loss"]["mean_loss"], 0)
        self.assertLess(CASES["Near the variance limit"]["inputs"]["df"], 3)
        # A zero mean return over one day: the VaR is the volatility times the Normal quantile.
        one_day = CASES["One day, zero mean"]
        self.assertEqual(one_day["results"]["mean_loss"], 0)
        # The table widens to hold the far expected shortfall of a heavy tail.
        self.assertGreater(CASES["Rare tail"]["top"], P.MIN_TOP)

    def test_var_and_es_are_ordered(self):
        for item in DATA["cases"]:
            results = item["results"]
            self.assertGreaterEqual(results["normal_es"], results["normal_var"], item["name"])
            self.assertGreaterEqual(results["t_es"], results["t_var"], item["name"])
            # The Binomial quantile leaves at most 1 - confidence above it.
            self.assertLessEqual(results["default_exceed"], 1 - item["inputs"]["confidence"], item["name"])
            self.assertEqual(results["default_var"], results["default_quantile"] * item["inputs"]["loss_per_default"])

    def test_the_scale_matches_the_variance(self):
        for item in DATA["cases"]:
            nu = item["inputs"]["df"]
            results = item["results"]
            self.assertTrue(close(results["t_scale"] ** 2 * nu / (nu - 2), results["sd_loss"] ** 2, 1e-14), item["name"])

    def test_the_top_of_the_table_is_not_on_a_boundary(self):
        # The sheet forms INT((ES - m) / sd) from its own Double results: a ratio within 1E-6 of a whole
        # number could round the other way and move every row of the table.
        for item in DATA["cases"]:
            results = item["results"]
            ratio = (max(results["normal_es"], results["t_es"]) - results["mean_loss"]) / results["sd_loss"]
            self.assertGreater(abs(ratio - round(ratio)), 1e-6, item["name"])
            self.assertEqual(item["top"], max(P.MIN_TOP, math.floor(ratio) + 1), item["name"])

    def test_the_table_is_the_sheets_grid(self):
        for item in DATA["cases"]:
            mean, sd = item["results"]["mean_loss"], item["results"]["sd_loss"]
            top = item["top"]
            for j, row in enumerate(item["table"]):
                z = j * (top - P.FIRST_Z) / (P.ROWS - 1) + P.FIRST_Z
                self.assertEqual(row["z"], z)
                self.assertEqual(row["x"], mean + sd * z)
            self.assertEqual(item["table"][0]["z"], P.FIRST_Z)
            self.assertEqual(item["table"][-1]["z"], top)

    def test_the_provenance_is_recorded(self):
        for key in ("python", "mpmath"):
            self.assertRegex(DATA["environment"][key], r"^\d+\.\d+")
        for item in DATA["cases"]:
            self.assertTrue(item["purpose"], item["name"])

    def test_the_stored_references_are_what_mpmath_returns_now(self):
        if R.mp is None:
            self.skipTest("mpmath is not installed")
        again = R.build()
        self.assertEqual(again["cases"], DATA["cases"])


@unittest.skipIf(mp is None, "mpmath is not installed")
class IndependentReferenceTests(unittest.TestCase):
    """Each reference again, by another method."""

    def setUp(self):
        mp.dps = 30

    def density(self, nu):
        constant = mp.gamma((nu + 1) / 2) / (mp.sqrt(nu * mp.pi) * mp.gamma(nu / 2))
        return lambda u: constant * (1 + u * u / nu) ** (-(nu + 1) / 2)

    def test_expected_shortfall_by_integration(self):
        for item in DATA["cases"]:
            results = item["results"]
            level = mpf(item["inputs"]["confidence"])
            mean, sd, scale = (mpf(results[key]) for key in ("mean_loss", "sd_loss", "t_scale"))
            nu = mpf(item["inputs"]["df"])
            z = (mpf(results["normal_var"]) - mean) / sd
            normal = mp.quad(lambda u: u * mp.npdf(u), [z, mp.inf]) / (1 - level)
            self.assertTrue(close(results["normal_es"], mean + sd * normal, 1e-11, float(sd)), item["name"])
            t = (mpf(results["t_var"]) - mean) / scale
            f = self.density(nu)
            student = mp.quad(lambda u: u * f(u), [t, 10 * t, mp.inf]) / (1 - level)
            self.assertTrue(close(results["t_es"], mean + scale * student, 1e-9, float(sd)), item["name"])

    def test_student_tail_and_quantile_by_integrating_the_density(self):
        for item in DATA["cases"]:
            results = item["results"]
            nu = mpf(item["inputs"]["df"])
            mean, scale = mpf(results["mean_loss"]), mpf(results["t_scale"])
            f = self.density(nu)
            # The quantile: the tail beyond the stored VaR is 1 - confidence.
            t = (mpf(results["t_var"]) - mean) / scale
            beyond = mp.quad(f, [t, 10 * t, mp.inf])
            self.assertTrue(close(beyond, 1 - item["inputs"]["confidence"], 1e-10), item["name"])
            # The exceedance probability of the threshold.
            u = (mpf(item["inputs"]["threshold"]) - mean) / scale
            tail = mp.quad(f, [u, abs(u) + 10, mp.inf]) if u > 0 else 1 - mp.quad(f, [-mp.inf, u])
            self.assertTrue(close(results["t_exceed"], tail, 1e-10), item["name"])
            # A few rows of the table.
            for row in item["table"][::25]:
                u = (mpf(row["x"]) - mean) / scale
                tail = mp.quad(f, [u, abs(u) + 10, mp.inf]) if u > 0 else 1 - mp.quad(f, [-mp.inf, u])
                self.assertTrue(close(row["t_exceed"], tail, 1e-10), f'{item["name"]} {row["z"]}')
                self.assertTrue(close(row["t_density"], f(u) / scale, 1e-12), f'{item["name"]} {row["z"]}')

    def test_normal_by_the_complementary_error_function(self):
        for item in DATA["cases"]:
            results = item["results"]
            mean, sd = mpf(results["mean_loss"]), mpf(results["sd_loss"])
            z = (mpf(results["normal_var"]) - mean) / sd
            self.assertTrue(close(mp.erfc(z / mp.sqrt(2)) / 2, 1 - item["inputs"]["confidence"], 1e-12), item["name"])
            threshold = mpf(item["inputs"]["threshold"])
            self.assertTrue(close(results["normal_exceed"], mp.erfc((threshold - mean) / (sd * mp.sqrt(2))) / 2, 1e-12))
            for row in item["table"][::10]:
                u = (mpf(row["x"]) - mean) / sd
                self.assertTrue(close(row["normal_exceed"], mp.erfc(u / mp.sqrt(2)) / 2, 1e-12), item["name"])
                self.assertTrue(close(row["normal_density"], mp.exp(-u * u / 2) / (sd * mp.sqrt(2 * mp.pi)), 1e-12))

    def test_binomial_by_a_direct_sum(self):
        for item in DATA["cases"]:
            inputs, results = item["inputs"], item["results"]
            n, p = inputs["loans"], mpf(inputs["default_probability"])
            masses = [mp.binomial(n, j) * p ** j * (1 - p) ** (n - j) for j in range(n + 1)]
            q = int(results["default_quantile"])
            below = mp.fsum(masses[: q + 1])
            self.assertGreaterEqual(below, mpf(inputs["confidence"]), item["name"])
            if q > 0:
                self.assertLess(mp.fsum(masses[:q]), mpf(inputs["confidence"]), item["name"])
            self.assertTrue(close(results["default_exceed"], mp.fsum(masses[q + 1:]), 1e-12), item["name"])
            self.assertEqual(results["expected_defaults"], n * inputs["default_probability"])

    def test_against_scipy(self):
        try:
            from scipy import stats
        except ImportError:
            self.skipTest("SciPy is not installed")
        for item in DATA["cases"]:
            inputs, results = item["inputs"], item["results"]
            mean, sd, scale = results["mean_loss"], results["sd_loss"], results["t_scale"]
            self.assertTrue(close(results["normal_var"], stats.norm.ppf(inputs["confidence"], mean, sd), 1e-9, sd))
            self.assertTrue(close(results["t_var"], mean + scale * stats.t.ppf(inputs["confidence"], inputs["df"]), 1e-9, sd))
            self.assertEqual(results["default_quantile"],
                             stats.binom.ppf(inputs["confidence"], inputs["loans"], inputs["default_probability"]))
            for row in item["table"][::10]:
                self.assertTrue(close(row["t_exceed"], stats.t.sf((row["x"] - mean) / scale, inputs["df"]), 1e-9))


class PlanTests(unittest.TestCase):
    def formulas(self):
        return {item["address"]: item["formula"] for item in PLAN["formulas"]}

    def test_every_formula_and_name_fits_excels_limit(self):
        for name, size in P.formula_lengths(PLAN):
            self.assertLess(size, EXCEL_LIMIT, name)

    def test_cells_are_written_once(self):
        addresses = [item["address"] for item in PLAN["formulas"]] + [item["address"] for item in PLAN["texts"]]
        self.assertEqual(len(addresses), len(set(addresses)), "a cell is written twice")

    def test_the_sheet_names_are_its_own(self):
        others = set()
        for module in (explorer_plan, probabilities_plan, parameters_plan, comparison_plan, tails_plan, discrete_plan):
            others |= {item["name"] for item in module.plan()["names"]}
        sheet = {item["name"] for item in PLAN["names"]}
        self.assertEqual(others & sheet, set())
        self.assertTrue(all(name.startswith("Risk") for name in sheet))
        text = " ".join(item["formula"] for item in PLAN["formulas"]) + " " + " ".join(
            item["refers"] for item in PLAN["names"]) + " " + " ".join(
            item["x"] + " " + item["y"] for chart in PLAN["charts"] for item in chart["series"]) + " " + " ".join(
            PLAN["validation"]["custom"].values()) + " " + " ".join(PLAN["conditional"].values())
        used = set(re.findall(r"\bRisk[A-Z][A-Za-z]*\b", re.sub(r'"[^"]*"', '""', text)))
        self.assertLessEqual(used, sheet, "a formula reads a name the plan does not define")

    def test_every_call_is_a_public_function(self):
        code = " ".join(self.formulas().values()) + " ".join(item["refers"] for item in PLAN["names"])
        calls = set(re.findall(r"K_STATS_\w+", re.sub(r'"[^"]*"', '""', code)))
        self.assertLessEqual(calls, PUBLIC)
        families = {name.split("_")[2] for name in calls}
        self.assertEqual(families, {"Normal", "NormalStandard", "StudentT", "Binomial"})

    def test_the_formula_shown_beside_each_result_is_the_formula_it_computes(self):
        formulas = self.formulas()
        for key in P.VALUE_KEYS:
            row = P.RESULT_ROWS[key]
            value, shown = formulas[f"D{row}"], formulas[f"F{row}"]
            expression = re.fullmatch(r'="=(.*)"', shown).group(1)
            self.assertEqual(value, f'=IF($D$19="OK",{expression},NA())', key)

    def test_the_definitions_are_the_textbook_ones(self):
        text = P.expressions()
        self.assertEqual(text["t_scale"], "D22*SQRT((D11-2)/D11)")
        self.assertIn("/(1-D10)", text["normal_es"])
        self.assertIn("(D11+K_STATS_StudentT_InverseCumulative(D10,D11)^2)/(D11-1)", text["t_es"])
        self.assertEqual(text["normal_var"], "K_STATS_Normal_InverseCumulative(D10,D21,D22)")
        self.assertEqual(text["default_quantile"], "K_STATS_Binomial_InverseCumulative(D10,D13,D14)")
        # A tail is a survival call, never 1 - CDF: no result or table cell calls a cumulative function (the
        # macros cell D20 calls one only to see whether the API answers).
        code = re.sub(r'"[^"]*"', '""', " ".join(value for address, value in self.formulas().items() if address != "D20"))
        self.assertNotRegex(code, r"1-K_STATS_")
        self.assertNotIn("_Cumulative(", code.replace("_InverseCumulative(", ""))

    def test_the_reading_survives_a_zero_normal_value(self):
        # At confidence 0.5 with a zero mean return the Normal VaR is 0: no ratio divides by it.
        reading = self.formulas()[P.result("reading")]
        for normal, student in (("$D$23", "$D$26"), ("$D$24", "$D$27")):
            self.assertIn(f"IF({normal}=0,", reading)
            self.assertIn(f"ROUND({student}/{normal},3)", reading)
            self.assertLess(reading.index(f"IF({normal}=0,"), reading.index(f"ROUND({student}/{normal},3)"))

    def test_the_degrees_of_freedom_must_give_a_variance(self):
        names = {item["name"]: item["refers"] for item in PLAN["names"]}
        self.assertIn("Risk!$D$11>2", names["RiskDfOk"])
        self.assertIn("variance is not finite", P.status_cell_formula())

    def test_the_table_is_the_one_of_the_references(self):
        formulas = self.formulas()
        z = formulas[f"{P.HELPERS['z']}{P.FIRST_ROW}"]
        self.assertIn(f"SEQUENCE({P.ROWS},1,0,1)*(MAX({P.MIN_TOP},INT((MAX($D$24,$D$27)-$D$21)/$D$22)+1)-(-4))/100+(-4)", z)
        self.assertEqual((R.ROWS, R.FIRST_Z, R.MIN_TOP), (P.ROWS, P.FIRST_Z, P.MIN_TOP))

    def test_lambda_parameters_are_not_function_or_cell_names(self):
        formulas = " ".join(self.formulas().values())
        for parameter in re.findall(r"LAMBDA\(([A-Za-z0-9_,]+),", formulas):
            for name in parameter.split(","):
                self.assertNotRegex(name, r"^[A-Za-z]{1,3}\d+$|^[rRcC]$", name)

    def test_nothing_reads_as_a_formula_unless_it_is_one(self):
        for item in PLAN["texts"]:
            self.assertNotIn(item["value"][:1], "=+-@", item["address"])

    def test_every_quoted_constant_fits_excels_limit(self):
        # Excel refuses a string constant longer than 255 characters inside a formula.
        for item in PLAN["formulas"]:
            for constant in re.findall(r'"((?:[^"]|"")*)"', item["formula"]):
                self.assertLessEqual(len(constant), 255, item["address"])

    def test_no_cell_needs_formulatext_or_python(self):
        code = re.sub(r'"[^"]*"', '""', " ".join(self.formulas().values()))
        self.assertNotIn("FORMULATEXT(", code)
        self.assertNotIn("PY(", code)
        self.assertNotIn("TEXT(", code)

    def test_conditional_formats_need_no_list_separator(self):
        for key, value in PLAN["conditional"].items():
            self.assertNotIn(",", value, key)
        for value in PLAN["validation"]["custom"].values():
            self.assertRegex(value, r"^=Risk[A-Za-z]+$")

    def test_every_input_has_a_check_and_messages(self):
        self.assertEqual(set(PLAN["validation"]["custom"]), set(P.INPUTS.values()))
        self.assertEqual(set(PLAN["messages"]), set(P.INPUTS.values()))
        for address, texts in PLAN["messages"].items():
            self.assertEqual(len(texts), 4, address)
            self.assertTrue(all(texts), address)

    def test_the_layout_is_the_documented_one(self):
        self.assertEqual(PLAN["layout"]["inputs"], "D6:D15")
        self.assertEqual(PLAN["layout"]["first_row"], 41)
        self.assertEqual(PLAN["layout"]["last_row"], 40 + P.ROWS)
        self.assertTrue(all(21 <= row <= 38 for row in P.RESULT_ROWS.values()))
        self.assertEqual(PLAN["layout"]["chart_area"], "K5:O35")
        self.assertEqual(len(PLAN["charts"]), 2)
        self.assertEqual(json.loads(json.dumps(PLAN))["schema"], "risk-plan-1")

    def test_the_module_defaults_match_the_data(self):
        defaults = PLAN["defaults"]
        for key, constant in (("value", "VALUE"), ("mean_return", "MEAN_RETURN"), ("volatility", "VOLATILITY"),
                              ("horizon", "HORIZON"), ("confidence", "CONFIDENCE"), ("df", "DF"),
                              ("threshold", "THRESHOLD"), ("loans", "LOANS"),
                              ("default_probability", "DEFAULT_PROBABILITY"), ("loss_per_default", "LOSS_PER_DEFAULT")):
            self.assertIn(f"DEFAULT_{constant} As Double = {defaults[key]}", MODULE)
            self.assertIn(f'{constant}_CELL As String = "{P.INPUTS[key]}"', MODULE)
        self.assertIn(f'BUILD_CELL As String = "D{P.RESULT_ROWS["excel_build"]}"', MODULE)
        self.assertIn(f'NORMAL_VAR_CELL As String = "D{P.RESULT_ROWS["normal_var"]}"', MODULE)
        self.assertIn(f'T_VAR_CELL As String = "D{P.RESULT_ROWS["t_var"]}"', MODULE)
        for chart in PLAN["charts"]:
            self.assertIn(chart["name"], MODULE)


if __name__ == "__main__":
    unittest.main()
