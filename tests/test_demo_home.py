"""The plan of the Home and README sheets (tools/demobuild/home_plan.py)."""
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools" / "demobuild"))

import build_demo  # noqa: E402
import home_plan  # noqa: E402

EVOLVE = ROOT / "tools" / "demobuild" / "Evolve-Home.ps1"
WORKFLOW = ROOT / ".github" / "workflows" / "demo-explorer.yml"


class HomePlanTests(unittest.TestCase):
    def setUp(self):
        self.plan = home_plan.plan()
        self.sheets = {sheet["sheet"]: sheet for sheet in self.plan["sheets"]}

    def texts(self):
        for sheet in self.plan["sheets"]:
            yield sheet["subtitle"]
            for table in sheet["tables"]:
                yield table["heading"]
                for row in table["rows"]:
                    yield row["label"]
                    yield row["text"]

    def test_home_and_readme_are_planned(self):
        self.assertEqual(list(self.sheets), ["Home", "README"])
        self.assertEqual(self.plan["version"], (ROOT / "VERSION").read_text(encoding="utf-8").strip())

    def test_home_links_every_other_sheet_by_its_defined_name(self):
        links = [table for table in self.sheets["Home"]["tables"] if table["kind"] == "links"]
        self.assertEqual(len(links), 1)
        expected = [sheet for sheet in build_demo.sheet_map() if sheet.code != "shHome"]
        rows = links[0]["rows"]
        self.assertEqual([row["label"] for row in rows], [sheet.tab for sheet in expected])
        self.assertEqual([row["target"] for row in rows], ["nav_" + sheet.code for sheet in expected])
        # ci/Test-DemoPackage.ps1 follows nine links.
        self.assertEqual(len(rows), 9)

    def test_home_covers_what_issue_253_asks(self):
        headings = [table["heading"] for table in self.sheets["Home"]["tables"]]
        self.assertEqual(headings, ["What this workbook shows", "What it needs", "How to use", "Colour legend",
                                    "Sheets"])
        labels = {row["label"] for table in self.sheets["Home"]["tables"] for row in table["rows"]}
        for label in ("Demonstration", "Validation", "Version", "Excel", "Macros", "4. Reset inputs"):
            self.assertIn(label, labels)

    def test_the_first_heading_is_the_one_the_package_test_reads(self):
        self.assertEqual(self.sheets["Home"]["tables"][0]["heading"], "What this workbook shows")
        readme_labels = [row["label"] for table in self.sheets["README"]["tables"] for row in table["rows"]]
        self.assertIn("Source commit", readme_labels)
        commit = next(row for table in self.sheets["README"]["tables"] for row in table["rows"]
                      if row["label"] == "Source commit")
        self.assertEqual(commit["text"], "{commit}")

    def test_only_known_fields_are_used(self):
        for text in self.texts():
            for field in re.findall(r"\{[^}]*\}", text):
                self.assertIn(field, home_plan.FIELDS, text)

    def test_every_table_has_two_rows_and_a_height_that_fits(self):
        for sheet in self.plan["sheets"]:
            for table in sheet["tables"]:
                self.assertGreaterEqual(len(table["rows"]), 2)
                for row in table["rows"]:
                    self.assertEqual(row["height"], home_plan.height(row["text"]))
                    self.assertIn(row["height"], home_plan.HEIGHTS)
        with self.assertRaises(ValueError):
            home_plan.height("x" * (3 * home_plan.TEXT_PER_LINE + 1))

    def test_the_legend_uses_the_swatches_the_script_knows(self):
        legend = next(table for table in self.sheets["Home"]["tables"] if table["kind"] == "legend")
        swatches = [row["swatch"] for row in legend["rows"] if "swatch" in row]
        self.assertEqual(swatches, list(home_plan.SWATCHES))
        script = EVOLVE.read_text(encoding="utf-8")
        for swatch in home_plan.SWATCHES:
            self.assertIn(f"'{swatch}' {{", script)

    def test_cited_files_exist(self):
        for text in self.texts():
            for path in re.findall(r"(?:docs|src|tools)/[\w./-]+\.\w+", text):
                self.assertTrue((ROOT / path).exists(), path)

    def test_the_excel_functions_named_are_used_by_the_sheets(self):
        excel = next(row["text"] for table in self.sheets["Home"]["tables"] for row in table["rows"]
                     if row["label"] == "Excel")
        plans = "".join(path.read_text(encoding="utf-8") for path in (ROOT / "tools" / "demobuild").glob("*_plan.py"))
        for function in ("LET", "LAMBDA", "MAP", "VSTACK"):
            self.assertIn(function, excel)
            self.assertIn(function + "(", plans)

    def test_the_workflow_writes_home_after_every_other_sheet(self):
        steps = WORKFLOW.read_text(encoding="utf-8")
        home = steps.index("Evolve-Home.ps1")
        self.assertGreater(home, steps.index("Evolve-Diagnostics.ps1"))
        self.assertLess(home, steps.index("ci\\Test-UdfRegistration.ps1"))
        self.assertIn("home_plan.plan()", (ROOT / "tools" / "demobuild" / "explorer_build.py").read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
