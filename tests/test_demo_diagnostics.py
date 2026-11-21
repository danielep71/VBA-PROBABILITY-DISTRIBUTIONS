"""The Diagnostics sheet (errors and diagnostics): its cases, its expected outcomes and its plan.

The outcomes of examples/demo/diagnostics-cases.json come from the documentation: each case cites the rule
that gives it, and here each rule is checked to exist where it is cited. The numbers are recomputed by another
method (SciPy where it is installed, closed forms otherwise), and every fragment of a Status the sheet expects is
checked to be in the library's own messages: the sheet shows no diagnostic the library does not write.
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
import diagnostics_plan as P
import diagnostics_references as R
import discrete_plan
import explorer_plan
import parameters_plan
import probabilities_plan
import risk_plan
import tails_plan

PLAN = P.plan()
DATA = P.load()
CATALOGUE = {item["name"]: item for item in DATA["functions"]}
EXCEL_LIMIT = 8192
MODULE = (ROOT / "src/workbook/M_STATS_PROBDIST_DIAGNOSTICS.bas").read_text(encoding="utf-8")
LIBRARY = "\n".join(path.read_text(encoding="utf-8", errors="replace") for path in (ROOT / "src/modules").glob("*.bas"))
PUBLIC = set(re.findall(r"Public Function (K_STATS_\w+)", LIBRARY))
MANUAL = ROOT / "docs/user-manual/v1.0.0"

try:
    from mpmath import mp
except ImportError:
    mp = None


def declarations():
    """{name: [parameter names without Status]} from docs/PUBLIC_API.txt."""
    out = {}
    for line in (ROOT / "docs/PUBLIC_API.txt").read_text(encoding="utf-8").splitlines():
        if line.startswith("#") or "\t" not in line:
            continue
        name, declaration = line.split("\t")[2:4]
        inside = re.search(r"\((.*)\)", declaration).group(1)
        names = [re.match(r"(?:Optional )?By(?:Val|Ref) (\w+) As", part).group(1) for part in inside.split(", ") if part]
        out[name] = [item for item in names if item != "Status"]
    return out


def select_cases(procedure):
    """The function names of each Case of a Select Case in a procedure of the module, in order, with the
    lines that follow each Case."""
    body = re.search(rf"Function {procedure}\((.*?)End Function", MODULE, re.S).group(1)
    return body


class CaseTests(unittest.TestCase):
    def test_the_cases_cover_every_kind_of_outcome(self):
        groups = {case["group"] for case in DATA["cases"]}
        self.assertEqual(groups, set(P.GROUPS))
        kinds = {case["expected"]["kind"] for case in DATA["cases"]}
        self.assertEqual(kinds, {"number", "#NUM!", "#VALUE!"})
        self.assertGreaterEqual(sum(case["group"] == "invalid" for case in DATA["cases"]), 12)

    def test_the_groups_match_the_outcomes(self):
        for case in DATA["cases"]:
            kind = case["expected"]["kind"]
            if case["group"] in ("valid", "limit", "underflow"):
                self.assertEqual(kind, "number", case["id"])
            if case["group"] == "invalid":
                self.assertEqual(kind, "#NUM!", case["id"])

    def test_every_rule_cites_a_place_of_the_documentation(self):
        headings = set()
        for path in MANUAL.glob("*.md"):
            headings |= set(re.findall(r"^## (\d+\.\d+) ", path.read_text(encoding="utf-8"), re.M))
        api = (ROOT / "docs/PUBLIC_API.md").read_text(encoding="utf-8")
        for case in DATA["cases"]:
            sections = re.findall(r"\b(\d+\.\d+):", case["rule"])
            self.assertTrue(sections or case["rule"].startswith("PUBLIC_API.md"), case["id"])
            for section in sections:
                self.assertIn(section, headings, case["id"])
        self.assertIn("| A text | `#VALUE!` |", api)
        self.assertIn("| An empty cell | What 0 returns |", api)

    def test_every_status_fragment_is_a_message_of_the_library(self):
        messages = " ".join(re.findall(r'"([^"]*)"', LIBRARY))
        fragments = [case["status_contains"] for case in DATA["cases"] + DATA["trials"]
                     if case["status_contains"] and case["status_contains"] != "Not called"]
        self.assertGreater(len(fragments), 10)
        for fragment in fragments:
            self.assertIn(fragment, messages)

    def test_a_number_has_no_status_and_an_error_may_name_one(self):
        for case in DATA["cases"] + DATA["trials"]:
            if case["expected"]["kind"] == "number":
                self.assertIn(case["status_contains"], (None, ""), case.get("id", case.get("name")))

    def test_every_call_has_the_arguments_of_its_declaration(self):
        declared = declarations()
        for item in DATA["functions"]:
            self.assertEqual(item["params"], declared[item["name"]], item["name"])
        for case in DATA["cases"] + DATA["trials"]:
            self.assertEqual(len(case["args"]), len(CATALOGUE[case["function"]]["params"]), case["function"])

    def test_every_example_is_inside_the_domain_and_sweeps_forward(self):
        for item in DATA["functions"]:
            self.assertLess(item["sweep"][0], item["sweep"][1], item["name"])
            self.assertLessEqual(max(abs(value) for value in item["sweep"]), P.SWEEP_LIMIT)
            if item["params"][0] == "Probability":
                self.assertTrue(0 < item["example"][0] < 1, item["name"])
                self.assertEqual(item["sweep"], [0, 1], item["name"])
            self.assertLessEqual(len(item["domain"]), 255, item["name"])

    def test_the_default_function_is_the_module_default(self):
        self.assertEqual(DATA["default_function"], "K_STATS_Gamma_InverseCumulative")
        self.assertIn(f'DEFAULT_FUNCTION As String = "{DATA["default_function"]}"', MODULE)
        self.assertEqual(DATA["trials"][0]["function"], DATA["default_function"])
        self.assertEqual(DATA["trials"][0]["args"], CATALOGUE[DATA["default_function"]]["example"])

    def test_the_provenance_is_recorded(self):
        self.assertIn("mpmath", DATA["environment"])
        self.assertEqual(len(DATA["reference"]["sources"]), 4)
        for source in DATA["reference"]["sources"]:
            self.assertTrue((ROOT / source).exists(), source)

    @unittest.skipIf(mp is None, "mpmath is not installed")
    def test_the_stored_outcomes_are_what_the_script_gives_now(self):
        fresh = R.build()
        for key in ("functions", "cases", "trials", "sweep_points", "default_function"):
            self.assertEqual(json.loads(json.dumps(fresh[key])), DATA[key], key)

    def test_the_numbers_by_another_method(self):
        values = {case["id"]: case["expected"].get("value") for case in DATA["cases"]}
        self.assertAlmostEqual(values["c01"], 0.5 * math.erfc(-1.5 / math.sqrt(2)), delta=1e-16)
        self.assertEqual(values["c06"], 3 ** 2 * math.exp(-3) / 2)
        self.assertEqual(values["c07"], 6 / 9)
        self.assertAlmostEqual(values["c11"] / (0.5 * math.erfc(10 / math.sqrt(2))), 1, delta=1e-14)
        # The Normal density at 40, exp(-800) / sqrt(2 pi), and the Gamma quantile, about 1E-30000: both below the
        # smallest positive Double, 4.9E-324.
        self.assertLess(-800 / math.log(10) - 0.5 * math.log10(2 * math.pi), -323.3)
        self.assertEqual((values["c09"], values["c12"]), (0.0, 0.0))
        self.assertEqual(values["c10"], 1.0)
        trials = {trial["name"]: trial["expected"].get("value") for trial in DATA["trials"]}
        # Binomial(10, 0.3): F(3) = 0.6496 < 0.8 <= F(4) = 0.8497, so the quantile is 4.
        self.assertEqual(trials["a discrete quantile"], 4)
        self.assertEqual(trials["an empty X"], 0.5)

    def test_against_scipy(self):
        try:
            from scipy import stats
        except ImportError:
            self.skipTest("SciPy is not installed")
        values = {case["id"]: case["expected"].get("value") for case in DATA["cases"]}
        self.assertAlmostEqual(values["c01"] / stats.norm.cdf(1.5), 1, delta=1e-15)
        self.assertAlmostEqual(values["c06"] / stats.poisson.pmf(2, 3), 1, delta=1e-14)
        self.assertAlmostEqual(values["c11"] / stats.norm.sf(10), 1, delta=1e-14)
        default = DATA["trials"][0]["expected"]["value"]
        self.assertAlmostEqual(default / stats.gamma.ppf(0.95, 2, scale=1.5), 1, delta=1e-13)
        self.assertEqual(DATA["trials"][-1]["expected"]["value"], stats.binom.ppf(0.8, 10, 0.3))


class PlanTests(unittest.TestCase):
    def formulas(self):
        return {item["address"]: item["formula"] for item in PLAN["formulas"]}

    def test_every_formula_and_name_fits_excels_limit(self):
        for name, size in P.formula_lengths(PLAN):
            self.assertLess(size, EXCEL_LIMIT, name)

    def test_every_quoted_constant_fits_excels_limit(self):
        # Excel refuses a string constant longer than 255 characters inside a formula.
        for item in PLAN["formulas"]:
            for constant in re.findall(r'"((?:[^"]|"")*)"', item["formula"]):
                self.assertLessEqual(len(constant), 255, item["address"])

    def test_cells_are_written_once(self):
        addresses = [item["address"] for item in PLAN["formulas"]] + [item["address"] for item in PLAN["texts"]]
        self.assertEqual(len(addresses), len(set(addresses)), "a cell is written twice")

    def test_the_empty_cell_stays_empty(self):
        addresses = {item["address"] for item in PLAN["formulas"]} | {item["address"] for item in PLAN["texts"]}
        self.assertNotIn(P.EMPTY_CELL, addresses)
        self.assertIn({"name": "DiagEmpty", "refers": "=Diagnostics!$AE$1"}, PLAN["names"])
        self.assertIn("DiagEmpty", self.formulas()["H69"])

    def test_the_sheet_names_are_its_own(self):
        others = set()
        for module in (explorer_plan, probabilities_plan, parameters_plan, comparison_plan, tails_plan, discrete_plan,
                       risk_plan):
            others |= {item["name"] for item in module.plan()["names"]}
        sheet = {item["name"] for item in PLAN["names"]}
        self.assertEqual(others & sheet, set())
        self.assertTrue(all(name.startswith("Diag") for name in sheet))
        text = " ".join(item["formula"] for item in PLAN["formulas"]) + " " + " ".join(
            item["refers"] for item in PLAN["names"]) + " " + " ".join(
            item["x"] + " " + item["y"] for chart in PLAN["charts"] for item in chart["series"]) + " " + " ".join(
            PLAN["validation"]["custom"].values()) + " " + PLAN["validation"]["function_list"] + " " + \
            PLAN["conditional"]["status_bad"]
        used = set(re.findall(r"\bDiag[A-Z][A-Za-z]*\b", re.sub(r'"[^"]*"', '""', text)))
        self.assertLessEqual(used, sheet, "a formula reads a name the plan does not define")

    def test_every_call_is_a_public_function(self):
        code = re.sub(r'"[^"]*"', '""', " ".join(self.formulas().values()))
        self.assertLessEqual(set(re.findall(r"K_STATS_\w+", code)), PUBLIC)
        self.assertLessEqual(set(CATALOGUE), PUBLIC)

    def test_the_choose_follows_the_catalogue(self):
        returned = self.formulas()[P.result("returned")]
        calls = re.findall(r"(K_STATS_\w+)\(([^()]*)\)", returned)
        self.assertEqual([name for name, _ in calls], [item["name"] for item in DATA["functions"]])
        for name, arguments in calls:
            self.assertEqual(len(arguments.split(",")), len(CATALOGUE[name]["params"]), name)
        sweep = self.formulas()[f"{P.HELPERS['raw']}{P.FIRST_ROW}"]
        self.assertEqual(re.findall(r"(K_STATS_\w+)\(xv", sweep), [item["name"] for item in DATA["functions"]])

    def test_the_status_reader_covers_the_catalogue(self):
        arity = select_cases("DiagnosticsArity")
        counts = {}
        for block in re.split(r"\n\s*Case ", arity)[1:]:
            names = re.findall(r'"(K_STATS_\w+)"', block)
            value = re.search(r"DiagnosticsArity = (\d)", block)
            for name in names:
                counts[name] = int(value.group(1))
        self.assertEqual(counts, {name: len(item["params"]) for name, item in CATALOGUE.items()})
        calls = re.findall(r'Case "(K_STATS_\w+)"\s+DiagnosticsCall = (K_STATS_\w+)\(([^\n]*)\)',
                           select_cases("DiagnosticsCall"))
        self.assertEqual({case for case, _, _ in calls}, set(CATALOGUE))
        for case, called, arguments in calls:
            self.assertEqual(case, called)
            values = re.findall(r"Values\((\d)\)", arguments)
            self.assertEqual(values, [str(slot) for slot in range(1, len(CATALOGUE[case]["params"]) + 1)], case)
            self.assertTrue(arguments.endswith("Status"), case)

    def test_the_module_reads_the_catalogue_where_the_plan_writes_it(self):
        def column(letter):
            number = 0
            for character in letter:
                number = number * 26 + ord(character) - 64
            return number
        name = column(P.CATALOGUE["name"])
        self.assertEqual(column(P.CATALOGUE_EXAMPLES[0]) - name, 6)
        self.assertIn("Found.Offset(0, 5 + Slot)", MODULE)
        self.assertEqual(column(P.CATALOGUE["sweep_from"]) - name, 10)
        self.assertEqual(column(P.CATALOGUE["sweep_to"]) - name, 11)
        self.assertIn("Found.Offset(0, 10)", MODULE)
        self.assertIn("Found.Offset(0, 11)", MODULE)
        self.assertIn('CATALOGUE_NAMES As String = "DiagFunctionList"', MODULE)

    def test_the_module_cells_match_the_plan(self):
        self.assertIn(f'FUNCTION_CELL As String = "{P.FUNCTION_CELL}"', MODULE)
        self.assertIn(f'FIRST_ARGUMENT_CELL As String = "{P.ARG_CELLS[0]}"', MODULE)
        self.assertIn(f'SWEEP_FROM_CELL As String = "{P.SWEEP_CELLS[0]}"', MODULE)
        self.assertIn(f'SWEEP_TO_CELL As String = "{P.SWEEP_CELLS[1]}"', MODULE)
        self.assertIn(f'BUILD_CELL As String = "D{P.RESULT_ROWS["excel_build"]}"', MODULE)
        self.assertIn(f'SWEEP_CHART As String = "{PLAN["charts"][0]["name"]}"', MODULE)
        self.assertIn('Range("C7")', MODULE)

    def test_the_status_is_the_librarys_text(self):
        # DiagnosticsStatus returns the Status the call wrote; its own lines all start "Not called", and the
        # sheet tells them apart by those words.
        own = re.findall(r'Const \w+ As String = "(Not called[^"]*)"', MODULE)
        self.assertEqual(len(own), 3)
        self.assertIn("DiagnosticsStatus = Status", MODULE)
        self.assertIn('LEFT(D24,10)="Not called"', self.formulas()[P.result("meaning")])

    def test_the_case_table_checks_each_documented_outcome(self):
        formulas = self.formulas()
        for offset, case in enumerate(DATA["cases"]):
            row = P.FIRST_ROW + offset
            call = P.call_text(case["function"], case["args"])
            self.assertEqual(formulas[f"H{row}"], "=" + call)
            self.assertEqual(formulas[f"D{row}"], '="=' + call.replace('"', '""') + '"')
            if case["expected"]["kind"] == "number":
                self.assertEqual(float(formulas[f"I{row}"][1:]), case["expected"]["value"])
            else:
                self.assertEqual(formulas[f"I{row}"], f'="{case["expected"]["kind"]}"')
            self.assertIn('"As documented"', formulas[f"K{row}"])
            self.assertIn('"UNEXPECTED"', formulas[f"K{row}"])
            self.assertIn(f'DiagnosticsStatus("{case["function"]}"', formulas[f"L{row}"])
        self.assertEqual(PLAN["layout"]["last_row"], P.FIRST_ROW + len(DATA["cases"]) - 1)

    def test_numbers_are_written_to_give_back_the_same_double(self):
        for value in (1e-300, 1e100, 0.9331927987311419, 7.619853024160525e-24, 2.9, -0.5, 95, 1e9):
            self.assertEqual(float(P.number_text(value)), value)

    def test_lambda_parameters_are_not_function_or_cell_names(self):
        formulas = " ".join(self.formulas().values())
        for parameter in re.findall(r"LAMBDA\(([A-Za-z0-9_,]+),", formulas):
            for name in parameter.split(","):
                self.assertNotRegex(name, r"^[A-Za-z]{1,3}\d+$|^[rRcC]$", name)

    def test_nothing_reads_as_a_formula_unless_it_is_one(self):
        for item in PLAN["texts"]:
            if isinstance(item["value"], str):
                self.assertNotIn(item["value"][:1], "=+-@", item["address"])

    def test_conditional_formats_need_no_list_separator(self):
        self.assertNotIn(",", PLAN["conditional"]["status_bad"])
        for value in PLAN["validation"]["custom"].values():
            self.assertRegex(value, r"^=Diag[A-Za-z]+$")
        self.assertEqual([rule["text"] for rule in PLAN["conditional"]["check_rules"]], ["As documented", "UNEXPECTED"])

    def test_every_input_has_a_message_and_the_arguments_take_any_value(self):
        self.assertEqual(set(PLAN["messages"]), {P.FUNCTION_CELL, *P.ARG_CELLS, *P.SWEEP_CELLS})
        self.assertEqual(PLAN["validation"]["input_only"], list(P.ARG_CELLS))
        for address in P.ARG_CELLS:
            self.assertTrue(PLAN["messages"][address][1].startswith("Any value"))

    def test_the_meanings_are_the_manuals(self):
        chapter = (MANUAL / "09-errors-and-diagnostics.md").read_text(encoding="utf-8")
        self.assertIn("`#NUM!` means that the library understood the request and declined to answer it.", chapter)
        self.assertIn("understood the request and declined it", P.MEANING["num"])
        self.assertIn("Do not hide errors with `IFERROR` by default.", chapter)
        self.assertIn("IFERROR", P.ACTION["num"])
        self.assertIn("Non-convergence", P.NOT_SHOWN)
        self.assertIn("#418", P.NOT_SHOWN)

    def test_the_layout_is_the_documented_one(self):
        self.assertEqual(PLAN["layout"]["inputs"], "D6:D12")
        self.assertEqual(PLAN["layout"]["first_row"], 41)
        self.assertTrue(all(21 <= row <= 33 for row in P.RESULT_ROWS.values()))
        self.assertEqual(PLAN["layout"]["chart_area"], "K5:O35")
        self.assertEqual(len(PLAN["charts"]), 1)
        self.assertEqual(json.loads(json.dumps(PLAN))["schema"], "diagnostics-plan-1")
        sheets = json.loads((ROOT / "examples/demo/sheets.json").read_text(encoding="utf-8"))
        entry = [sheet for section in sheets["sections"] for sheet in section["sheets"] if sheet["issue"] == 261][0]
        self.assertEqual((entry["tab"], entry["code"]), (P.SHEET, "shDiagnostics"))
        self.assertIn('"shDiagnostics"', MODULE)


if __name__ == "__main__":
    unittest.main()
