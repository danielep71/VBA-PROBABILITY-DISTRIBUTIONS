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

    def verdict(self) -> str:
        return next(item["formula"] for item in PLAN["formulas"] if item["address"] == P.VERDICT)

    def test_the_verdict_never_reads_an_error_as_a_confirmation(self):
        # #483: F(Q - 1) was read through IFERROR(..., 0), so an error could confirm the quantile.
        verdict = self.verdict()
        self.assertNotRegex(verdict, r"IFERROR\([^()]*(\([^()]*\))*[^()]*,0\)")
        self.assertNotIn(",0)<$D$13", verdict)
        for guard in ("NOT(ISNUMBER(D25))", "NOT(ISNUMBER(D26))", "NOT(ISNUMBER(below))", "NOT(ISNUMBER(D27))"):
            self.assertIn(guard, verdict)
        self.assertTrue(verdict.endswith('"ERROR: the verdict could not be computed")'))

    def test_the_verdict_tells_invalid_input_from_the_ends_of_the_support(self):
        verdict = self.verdict()
        order = [verdict.index(text) for text in (
            "Invalid input: see the status", "Invalid input: p must be a number from 0 to 1", "End of the support",
            "ERROR: the quantile could not be computed")]
        self.assertEqual(order, sorted(order))
        self.assertIn('IF(OR($D$13=0,$D$13=1),"End of the support"', verdict)

    def test_the_lowest_count_is_decided_from_the_quantile_at_zero(self):
        # Q - 1 is below the support exactly when Q is the quantile at 0; F(Q - 1) is not computed there.
        verdict = self.verdict()
        lowest = verdict.index("Q is the smallest k with F(k) >= p: the lowest count")
        below = verdict.index("LET(below,")
        self.assertLess(lowest, below)
        self.assertRegex(verdict, r"IF\(D25=SWITCH\(.*?InverseCumulative\(0,")


if __name__ == "__main__":
    unittest.main()
