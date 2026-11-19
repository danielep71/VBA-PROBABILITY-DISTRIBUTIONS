"""The Probabilities sheet's plan: its formulas, names and texts, against the Explorer's table and the library."""
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools/demobuild"))
import explorer_plan
import probabilities_plan as P

PLAN = P.plan()
EXCEL_LIMIT = 8192


class PlanTests(unittest.TestCase):
    def test_every_formula_and_name_fits_excels_limit(self):
        for name, size in P.formula_lengths(PLAN):
            self.assertLess(size, EXCEL_LIMIT, name)

    def test_cells_are_written_once(self):
        addresses = [item["address"] for item in PLAN["formulas"]] + [item["address"] for item in PLAN["texts"]]
        self.assertEqual(len(addresses), len(set(addresses)), "a cell is both a formula and a text, or twice a formula")

    def test_names_of_the_family_table_exist_on_the_explorer(self):
        defined = {name for name in explorer_plan.NAMES}
        used = set(re.findall(r"\bExplorer[A-Za-z]+\b", " ".join(
            [item["formula"] for item in PLAN["formulas"]] + [item["refers"] for item in PLAN["names"]])))
        defined |= {"ExplorerFamily", "ExplorerDecimals"}  # defined by the Explorer plan beside the table
        self.assertLessEqual(used - defined, set(), "a name that the Explorer does not define")

    def test_the_sheet_names_do_not_collide_with_the_explorer(self):
        explorer = {item["name"] for item in explorer_plan.plan()["names"]}
        sheet = {item["name"] for item in PLAN["names"]}
        self.assertEqual(explorer & sheet, set())
        self.assertTrue(all(name.startswith("Probabilities") for name in sheet))

    def test_the_inputs_are_the_documented_cells(self):
        self.assertEqual(PLAN["layout"]["inputs"], "D6:D14")
        self.assertEqual((P.A, P.B, P.PROBABILITY, P.DECIMALS), ("$D$11", "$D$12", "$D$13", "$D$14"))
        for address in ("D6", "D11", "D12", "D13", "D14"):
            self.assertIn(address, PLAN["messages"])

    def test_messages_fit_the_limits_of_a_validation_rule(self):
        for address, (title, prompt, error_title, error) in PLAN["messages"].items():
            self.assertLessEqual(len(title), explorer_plan.TITLE_LIMIT, address)
            self.assertLessEqual(len(prompt), explorer_plan.PROMPT_LIMIT, address)
            self.assertLessEqual(len(error_title), explorer_plan.TITLE_LIMIT, address)
            self.assertLessEqual(len(error), explorer_plan.ERROR_LIMIT, address)

    def test_the_interval_uses_the_tail_that_keeps_its_digits(self):
        interval = next(item["formula"] for item in PLAN["formulas"] if item["address"] == P.INTERVAL)
        self.assertIn(f"IF({P.F_A}>0.5,{P.S_A}-{P.S_B},{P.F_B}-{P.F_A})", interval)
        self.assertIn("MAX(0,", interval)

    def test_every_function_called_exists_in_the_library(self):
        sys.path.insert(0, str(ROOT / "tests"))
        from test_demo_explorer import library_signatures
        known = library_signatures()
        called = set(re.findall(r"\bK_STATS_\w+", " ".join(item["formula"] for item in PLAN["formulas"])))
        self.assertTrue(called)
        self.assertLessEqual(called - set(known) - {"K_STATS_NormalStandard_Cumulative"}, set())

    def test_the_tolerance_is_declared_where_the_sheet_shows_it(self):
        verdict = next(item["formula"] for item in PLAN["formulas"] if item["address"] == P.VERDICT)
        self.assertIn("1E-08", verdict.upper())
        self.assertEqual(PLAN["roundtrip_tolerance"], 1e-8)

    def test_the_discrete_verdict_survives_a_quantile_on_the_lower_support_boundary(self):
        # Geometric(0.3) at p = 0.05 has Q = 0: F(Q - 1) asks the library for a count it refuses,
        # which must read as 0, not turn the verdict into #N/A.
        verdict = next(item["formula"] for item in PLAN["formulas"] if item["address"] == P.VERDICT)
        self.assertRegex(verdict, r"IFERROR\(SWITCH\(.*?D25-1.*?NA\(\)\),0\)<\$D\$13")


if __name__ == "__main__":
    unittest.main()
