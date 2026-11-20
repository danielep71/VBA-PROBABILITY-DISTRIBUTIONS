"""The Comparison sheet (Excel vs VBA vs Python): its cases, its Python code and its plan."""
import csv
import re
import sys
import unittest
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools/demobuild"))
sys.path.insert(0, str(ROOT / "examples/demo"))
import comparison_plan as C
import comparison_references as R
import explorer_plan
import parameters_plan
import probabilities_plan

PLAN = C.plan()
DATA = C.load()
EXCEL_LIMIT = 8192
MODULE = (ROOT / "src/workbook/M_STATS_PROBDIST_COMPARISON.bas").read_text(encoding="utf-8")
PUBLIC = set(re.findall(r"Public Function (K_STATS_\w+)",
                        "\n".join(path.read_text(encoding="utf-8", errors="replace")
                                  for path in (ROOT / "src/modules").glob("*.bas"))))


class CaseTests(unittest.TestCase):
    def test_the_cases_are_the_study_grid_unchanged(self):
        with R.GRID.open(encoding="utf-8", newline="") as handle:
            grid = list(csv.DictReader(handle, delimiter="|"))
        self.assertEqual(len(DATA["cases"]), len(grid))
        for case, row in zip(DATA["cases"], grid):
            self.assertEqual(case["label"], row["label"].strip())
            self.assertEqual(case["library"], row["kstats_call"].strip())
            self.assertEqual(case["reference"], row["reference"].strip())
            self.assertEqual(case["excel"] or "-", row["excel_formula"].strip())

    def test_every_library_call_is_a_public_function(self):
        for case in DATA["cases"]:
            name = case["library"].split("(")[0]
            self.assertIn(name, PUBLIC, case["label"])
            if case["excel"]:
                self.assertNotIn("K_STATS_", case["excel"], case["label"])

    def test_every_case_has_its_python_code(self):
        self.assertEqual(set(R.PYTHON), {case["library"] for case in DATA["cases"]})
        for case in DATA["cases"]:
            self.assertEqual(case["python"], R.PYTHON[case["library"]])
            self.assertTrue(case["python"].startswith("from scipy.stats import "), case["label"])

    def test_the_stored_scipy_values_are_reference_grade(self):
        # SciPy is graded like Excel and the library; on these cases it is within 12 digits.
        for case in DATA["cases"]:
            reference = float(Decimal(case["reference"]))
            value = float(case["scipy"])
            relative = abs(value - reference) / abs(reference) if reference else abs(value)
            self.assertLess(relative, 1e-12, case["label"])

    def test_the_stored_values_are_what_scipy_returns_now(self):
        try:
            import scipy
        except ImportError:
            self.skipTest("SciPy is not installed")
        if scipy.__version__ != DATA["scipy"]["scipy"]:
            self.skipTest(f"SciPy {scipy.__version__}, the values were stored with {DATA['scipy']['scipy']}")
        for case in DATA["cases"]:
            self.assertEqual(repr(R.run(case["python"])), case["scipy"], case["label"])

    def test_the_provenance_is_recorded(self):
        for key in ("python", "numpy", "scipy", "mpmath"):
            self.assertRegex(DATA["scipy"][key], r"^\d+\.\d+")
        self.assertEqual(DATA["reference"]["digits"], 50)
        self.assertEqual((DATA["grades"]["reference"], DATA["grades"]["report"]), (12, 6))


class PlanTests(unittest.TestCase):
    def test_every_formula_and_name_fits_excels_limit(self):
        for name, size in C.formula_lengths(PLAN):
            self.assertLess(size, EXCEL_LIMIT, name)

    def test_cells_are_written_once(self):
        addresses = ([item["address"] for item in PLAN["formulas"]] + [item["address"] for item in PLAN["texts"]]
                     + [item["address"] for item in PLAN["cells"]])
        self.assertEqual(len(addresses), len(set(addresses)), "a cell is written twice")

    def test_the_sheet_names_are_its_own(self):
        others = set()
        for module in (explorer_plan, probabilities_plan, parameters_plan):
            others |= {item["name"] for item in module.plan()["names"]}
        sheet = {item["name"] for item in PLAN["names"]}
        self.assertEqual(others & sheet, set())
        self.assertTrue(all(name.startswith("Comparison") for name in sheet))
        used = set(re.findall(r"\bComparison[A-Z][A-Za-z]*\b", " ".join(item["formula"] for item in PLAN["formulas"])))
        self.assertLessEqual(used, sheet, "a formula reads a name the plan does not define")

    def test_no_text_or_value_reads_as_a_formula(self):
        # A text that starts with =, +, - or @ is written by Value2 as a formula; the reference
        # texts keep the apostrophe of a text constant.
        for item in PLAN["texts"] + [cell for cell in PLAN["cells"] if isinstance(cell["value"], str)]:
            self.assertNotIn(item["value"][:1], "=+-@", item["address"])

    def test_every_case_has_its_live_formulas(self):
        formulas = {item["address"]: item["formula"] for item in PLAN["formulas"]}
        cells = {item["address"]: item["value"] for item in PLAN["cells"]}
        for index, case in enumerate(DATA["cases"]):
            row = C.FIRST_ROW + index
            self.assertEqual(formulas[f"G{row}"], "=" + case["library"])
            if case["excel"]:
                self.assertEqual(formulas[f"F{row}"], "=" + case["excel"])
            else:
                self.assertEqual(cells[f"F{row}"], C.NONE_TEXT)
            for column in "IKLMNO":
                self.assertIn(f"{column}{row}", formulas)
            self.assertEqual(cells[f"R{row}"], "'" + case["reference"])
            # The detail rows show these texts: they are the live formulas, written as text.
            self.assertEqual(cells[f"V{row}"], "'" + formulas[f"G{row}"])
            self.assertEqual(cells[f"U{row}"], "'" + formulas[f"F{row}"] if case["excel"] else C.NO_EXCEL_TEXT)

    def test_digits_and_grades_follow_the_study(self):
        formulas = {item["address"]: item["formula"] for item in PLAN["formulas"]}
        row = C.FIRST_ROW
        self.assertIn("-LOG10(rel)", formulas[f"I{row}"])
        self.assertIn(f"MIN({C.FULL_DIGITS}", formulas[f"I{row}"])
        self.assertIn("IF(truth=0,IF(obs=0,0,1)", formulas[f"I{row}"])
        grade = formulas[f"M{row}"]
        for word in ("no function", "error", "wrong", "reference", "report", "inadequate"):
            self.assertIn(f'"{word}"', grade)
        self.assertIn(C.REFERENCE_GRADE, grade)
        self.assertIn(C.REPORT_GRADE, grade)

    def test_the_layout_is_the_documented_one(self):
        self.assertEqual(PLAN["layout"]["inputs"], "D6:D9")
        self.assertEqual(PLAN["layout"]["last_row"], C.FIRST_ROW + len(DATA["cases"]) - 1)
        self.assertTrue(all(21 <= row <= 35 for row in C.DETAIL_ROWS.values()))
        self.assertEqual(PLAN["layout"]["chart"], "K5:O35")

    def test_the_module_defaults_match_the_data(self):
        self.assertIn(f'DEFAULT_CASE As String = "{DATA["default_case"]}"', MODULE)
        self.assertIn(f'DEFAULT_REFERENCE_GRADE As Double = {DATA["grades"]["reference"]}', MODULE)
        self.assertIn(f'DEFAULT_REPORT_GRADE As Double = {DATA["grades"]["report"]}', MODULE)
        self.assertIn(f'BUILD_CELL As String = "D{C.DETAIL_ROWS["excel_build"]}"', MODULE)
        self.assertIn(f'CHART_NAME As String = "{PLAN["chart"]["name"]}"', MODULE)

    def test_the_python_route_is_optional(self):
        formulas = {item["address"]: item["formula"] for item in PLAN["formulas"]}
        line = formulas[f"D{C.DETAIL_ROWS['python_grade']}"]
        self.assertTrue(line.startswith(f"=IF(ISBLANK({C.PYTHON_VALUE}),"))
        code = re.sub(r'"[^"]*"', '""', " ".join(formulas.values()))
        self.assertNotIn("PY(", code, "no cell may need Python in Excel")
        self.assertNotIn("FORMULATEXT(", code, "recalculated at every opening, it leaves the workbook changed")


if __name__ == "__main__":
    unittest.main()
