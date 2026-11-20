"""The Parameter Effects sheet's plan: its data, formulas, names and texts, against the Explorer's table and the library."""
import json
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools/demobuild"))
import explorer_plan
import parameters_plan as PP
import probabilities_plan

PLAN = PP.plan()
EXCEL_LIMIT = 8192


class DataTests(unittest.TestCase):
    def test_every_family_has_four_values_per_parameter_within_its_limits(self):
        data = PP.load()          # raises on a value outside the limits or a missing family
        self.assertEqual(len(data["families"]), 16)

    def test_the_parameter_that_is_varied_first_is_one_of_the_family(self):
        families = {f["name"]: f for f in explorer_plan.load()["families"]}
        for name, entry in PP.load()["families"].items():
            self.assertLessEqual(entry["vary"], len(families[name]["parameters"]), name)

    def test_the_curves_of_a_parameter_are_four_distinct_values_in_order(self):
        for name, entry in PP.load()["families"].items():
            for parameter in entry["parameters"]:
                values = parameter["curves"]
                self.assertEqual(values, sorted(set(values)), f"{name}: {values}")

    def test_the_moments_are_written_in_the_parameters_only(self):
        for name, entry in PP.load()["families"].items():
            count = len(entry["parameters"])
            for moment, expression in entry["moments"].items():
                letters = set(re.findall(r"\b([a-z])\b", re.sub(r'"[^"]*"', "", expression)))
                self.assertLessEqual(letters, set("abcd"[:count]), f"{name} {moment}: {expression}")

    def test_every_effect_is_a_sentence_that_fits_a_cell(self):
        for name, entry in PP.load()["families"].items():
            for parameter in entry["parameters"]:
                self.assertTrue(parameter["effect"].endswith("."), name)
                self.assertLess(len(parameter["effect"]), 260, name)


class PlanTests(unittest.TestCase):
    def test_every_formula_and_name_fits_excels_limit(self):
        for name, size in PP.formula_lengths(PLAN):
            self.assertLess(size, EXCEL_LIMIT, name)

    def test_cells_are_written_once(self):
        addresses = [item["address"] for item in PLAN["formulas"]] + [item["address"] for item in PLAN["texts"]]
        self.assertEqual(len(addresses), len(set(addresses)), "a cell is both a formula and a text, or twice a formula")

    def test_formulas_and_names_are_balanced(self):
        for item in PLAN["formulas"]:
            self.assertEqual(item["formula"].count("("), item["formula"].count(")"), item["address"])
            self.assertEqual(item["formula"].count('"') % 2, 0, item["address"])
        for item in PLAN["names"]:
            self.assertEqual(item["refers"].count("("), item["refers"].count(")"), item["name"])

    def test_names_of_the_family_table_exist_on_the_explorer(self):
        defined = set(explorer_plan.NAMES) | {"ExplorerFamily", "ExplorerDecimals"}
        used = set(re.findall(r"\bExplorer[A-Za-z]+\b", " ".join(
            [item["formula"] for item in PLAN["formulas"]] + [item["refers"] for item in PLAN["names"]])))
        self.assertLessEqual(used - defined, set())

    def test_the_sheet_names_do_not_collide_with_the_other_sheets(self):
        sheet = {item["name"] for item in PLAN["names"]}
        for other in (explorer_plan.plan()["names"], probabilities_plan.plan()["names"]):
            self.assertEqual({item["name"] for item in other} & sheet, set())
        self.assertTrue(all(name.startswith("ParameterEffects") for name in sheet))

    def test_a_curve_replaces_only_the_varied_parameter(self):
        text = PP.curve(2, "pa")
        self.assertIn("pa,IF(ParameterEffectsSlot=1,$D$13,$D$7)", text)
        self.assertIn("pd,IF(ParameterEffectsSlot=4,$D$13,$D$10)", text)

    def test_the_grid_formulas_call_the_library_with_the_curve_parameters(self):
        density = next(item["formula"] for item in PLAN["formulas"] if item["address"] == "S20")
        self.assertIn("K_STATS_Normal_Density(x,pa,pb)", density)
        self.assertNotIn("$D$7,", density.split("SWITCH")[1], "the dispatch still reads the input cells")

    def test_the_curve_windows_share_one_axis(self):
        names = {item["name"]: item["refers"] for item in PLAN["names"]}
        self.assertEqual(names["ParameterEffectsLower"], "=MIN(Parameters!$AX$22:$AX$25)")
        self.assertEqual(names["ParameterEffectsUpper"], "=MAX(Parameters!$AY$22:$AY$25)")

    def test_every_library_function_called_exists(self):
        sys.path.insert(0, str(ROOT / "tests"))
        from test_demo_explorer import library_signatures
        known = library_signatures()
        called = set(re.findall(r"\bK_STATS_\w+", " ".join(item["formula"] for item in PLAN["formulas"])))
        self.assertTrue(called)
        self.assertLessEqual(called - set(known) - {"K_STATS_NormalStandard_Cumulative"}, set())

    def test_the_hidden_table_does_not_overlap_the_grid_or_the_helper_columns(self):
        columns = {re.match(r"[A-Z]+", item["address"]).group(0) for item in PLAN["table"]["cells"]}
        helpers = {PP.X_COLUMN, *PP.DENSITY_COLUMNS, *PP.CDF_COLUMNS, *PP.WINDOW_COLUMNS, PP.LEGEND_COLUMN}
        self.assertEqual(columns & helpers, set())
        self.assertEqual(len(PLAN["table"]["cells"]), len(PP.load()["families"]) * (16 + 4 + 1))

    def test_messages_fit_the_limits_of_a_validation_rule(self):
        for address, (title, prompt, error_title, error) in PLAN["messages"].items():
            self.assertLessEqual(len(title), explorer_plan.TITLE_LIMIT, address)
            self.assertLessEqual(len(prompt), explorer_plan.PROMPT_LIMIT, address)
            self.assertLessEqual(len(error_title), explorer_plan.TITLE_LIMIT, address)
            self.assertLessEqual(len(error), explorer_plan.ERROR_LIMIT, address)
        for address in ("D6", "D11", "D12", "D13", "D14", "D15", "D16"):
            self.assertIn(address, PLAN["messages"])

    def test_the_charts_follow_the_four_curves_through_names(self):
        for chart in PLAN["charts"]:
            self.assertEqual(len(chart["series"]), 4)
            for series in chart["series"]:
                self.assertTrue(series["y"].startswith("ParameterEffects"))

    def test_the_tail_column_reads_the_survival_function_at_mean_plus_three_sd(self):
        tail = next(item["formula"] for item in PLAN["formulas"] if item["address"] == "I22")
        self.assertIn("K_STATS_Normal_Survival(F22+3*G22,pa,pb)", tail)


if __name__ == "__main__":
    unittest.main()
