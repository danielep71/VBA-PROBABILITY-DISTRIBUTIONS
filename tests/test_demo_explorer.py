"""The Distribution Explorer: its families, its plan, its references and its package."""
import importlib.util
import io
import json
import re
import sys
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools/demobuild"))
import explorer_build
import explorer_plan

DEMO = ROOT / "examples/demo"
FAMILIES = json.loads((DEMO / "explorer-families.json").read_text())
REFERENCES = json.loads((DEMO / "explorer-references.json").read_text())


def library_signatures():
    """Public K_STATS_ functions and the number of arguments before Status."""
    source = "\n".join(p.read_text(encoding="latin-1") for p in (ROOT / "src").glob("*.bas"))
    source = re.sub(r"_\s*\n\s*", "", source)
    found = {}
    for name, arguments in re.findall(r"Public Function (K_STATS_\w+)\((.*?)\)", source):
        names = [re.sub(r"(Optional |ByVal |ByRef )|\s+As\s+\w+.*", "", a).strip() for a in arguments.split(",")]
        found[name] = [n for n in names if n != "Status"]
    return found


class FamilyTests(unittest.TestCase):
    def test_sixteen_distinct_families_in_two_kinds(self):
        names = [family["name"] for family in FAMILIES["families"]]
        self.assertEqual(len(names), 16)
        self.assertEqual(len(set(names)), 16)
        kinds = {family["kind"] for family in FAMILIES["families"]}
        self.assertEqual(kinds, {"continuous", "discrete"})

    def test_each_function_exists_with_the_arity_the_call_uses(self):
        signatures = library_signatures()
        for family in FAMILIES["families"]:
            count = len(family["parameters"])
            self.assertEqual(family["call"].split(","), list("abcd"[:count]), family["name"])
            for operation in ("density", "cdf", "survival", "quantile"):
                name = family["functions"][operation]
                self.assertIn(name, signatures, family["name"])
                self.assertEqual(len(signatures[name]), count + 1, f"{name}: the point plus the parameters")

    def test_defaults_lie_inside_their_limits(self):
        for family in FAMILIES["families"]:
            for parameter in family["parameters"]:
                self.assertLessEqual(parameter["min"], parameter["default"], family["name"])
                self.assertLessEqual(parameter["default"], parameter["max"], family["name"])
                self.assertLessEqual(len(parameter["label"]), 26, "label must fit the label column")

    def test_references_cover_every_family_and_parameter_count(self):
        by_name = {family["name"]: family for family in FAMILIES["families"]}
        covered = {case["family"] for case in REFERENCES["cases"]}
        self.assertEqual(covered, set(by_name))
        self.assertEqual({grid["family"] for grid in REFERENCES["grids"]}, set(by_name))
        for case in REFERENCES["cases"]:
            self.assertEqual(len(case["parameters"]), len(by_name[case["family"]]["parameters"]))
            self.assertAlmostEqual(case["cdf"] + case["survival"], 1.0, places=12)
        self.assertEqual({case["family"] for case in REFERENCES["quantiles"]}, set(by_name))
        for grid in REFERENCES["grids"]:
            defaults = [parameter["default"] for parameter in by_name[grid["family"]]["parameters"]]
            self.assertEqual(grid["parameters"], defaults, "the grid reference is for the family's defaults")

    @unittest.skipUnless(importlib.util.find_spec("mpmath"), "mpmath is not installed")
    def test_reference_fixture_reproduces_independently(self):
        spec = importlib.util.spec_from_file_location("explorer_refs", DEMO / "explorer_references.py")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        self.assertEqual(module.references(), REFERENCES)


class PlanTests(unittest.TestCase):
    plan = explorer_plan.plan()

    def test_every_formula_fits_excel_and_has_no_implicit_intersection(self):
        formulas = [item["formula"] for item in self.plan["formulas"]]
        formulas += [cell["formula"] for cell in self.plan["table"]["cells"] if "formula" in cell]
        formulas += [item["refers"] for item in self.plan["names"]]
        for formula in formulas:
            self.assertLessEqual(len(formula), 8000)
            self.assertNotIn("@", formula)
        self.assertTrue(all(f["formula"].startswith("=") for f in self.plan["formulas"]))

    def test_output_formulas_spill_from_one_cell_each(self):
        by_address = {item["address"]: item["formula"] for item in self.plan["formulas"]}
        for address in ("F19", "G19", "H19", "I19"):
            self.assertIn(address, by_address)
        for address in ("G19", "H19", "I19"):
            self.assertIn("MAP(ExplorerX", by_address[address])
        self.assertIn("_Survival(x,", by_address["H19"])
        self.assertIn("_Cumulative(x,", by_address["I19"])
        self.assertIn("VSTACK", by_address["G7"])
        names = {item["name"]: item["refers"] for item in self.plan["names"]}
        self.assertEqual(names["ExplorerX"], "=Explorer!$F$19#")
        self.assertEqual(names["ExplorerPDF"], "=Explorer!$G$19#")
        self.assertEqual(names["ExplorerSurvival"], "=Explorer!$H$19#")
        self.assertEqual(names["ExplorerCDF"], "=Explorer!$I$19#")

    def test_dispatch_names_every_family_and_never_a_row_number(self):
        by_address = {item["address"]: item["formula"] for item in self.plan["formulas"]}
        for family in FAMILIES["families"]:
            self.assertIn(f'"{family["name"]}"', by_address["G19"])
            self.assertIn(family["functions"]["cdf"], by_address["I19"])
            self.assertIn(family["functions"]["quantile"], by_address["G10"])
        for formula in by_address.values():
            self.assertNotRegex(formula, r"[A-Z]+\$?(18|218)\b", "a row of the old fixed grid")

    def test_quantile_reads_only_p_and_the_status(self):
        by_address = {item["address"]: item["formula"] for item in self.plan["formulas"]}
        self.assertTrue(by_address["G10"].startswith('=IF(AND($G$5="OK",ISNUMBER($D$12),$D$12>=0,$D$12<=1),'))
        self.assertNotIn("$D$12", by_address["G5"], "a bad p must not stop the other results")

    def test_plot_window_and_decimals_cells(self):
        by_address = {item["address"]: item["formula"] for item in self.plan["formulas"]}
        self.assertIn("SEQUENCE($G$13,,0)", by_address["F19"])
        self.assertIn("-INT(-$G$11)", by_address["F19"])
        self.assertEqual(self.plan["decimals"]["cell"], "$G$14")
        self.assertEqual(self.plan["validation"]["decimals_list"], "=ExplorerDecimals")
        module = (DEMO / "M_STATS_PROBDIST_EXPLORER.bas").read_text(encoding="latin-1")
        self.assertIn(f'FORMATTED_CELLS As String = "{self.plan["decimals"]["cells"]}"', module)
        self.assertIn('DECIMALS_CELL As String = "G14"', module)
        self.assertIn(f'X_CELLS As String = "{self.plan["decimals"]["x_cells"]}"', module)
        values = [cell["value"] for cell in self.plan["table"]["cells"] if cell["address"].startswith("NC")]
        self.assertEqual(values, list(range(16)))

    def test_discrete_grid_starts_at_the_ceiling_of_a_negative_bound(self):
        by_address = {item["address"]: item["formula"] for item in self.plan["formulas"]}
        self.assertNotIn("ROUNDUP($G$11", by_address["F19"] + by_address["G5"])
        self.assertIn("-INT(-$G$11)", by_address["F19"])
        self.assertIn("-INT(-$G$11)", by_address["G5"])

    def test_texts_fit_the_template_cells(self):
        for text in self.plan["texts"]:
            if text["address"].startswith("C") and text["address"] != "C3":
                self.assertLessEqual(len(text["value"]), 110, text["address"])

    def test_family_table_is_complete_and_hidden_columns_are_consistent(self):
        cells = {cell["address"]: cell for cell in self.plan["table"]["cells"]}
        self.assertEqual(self.plan["table"]["rows"], 16)
        for index, family in enumerate(FAMILIES["families"], 1):
            self.assertEqual(cells[f"MA{index}"]["value"], family["name"])
        self.assertEqual(self.plan["table"]["first_column"], 339, "column MA, hidden in the baseline")
        for rule in ("ExplorerParameterOk1", "ExplorerParameterOk4", "ExplorerParameterUnused1", "ExplorerStatusBad"):
            self.assertIn(rule, {item["name"] for item in self.plan["names"]})


class PackageTests(unittest.TestCase):
    data, plan, hashes = explorer_build.build()

    def test_only_the_vba_project_changes_in_the_baseline(self):
        built = zipfile.ZipFile(io.BytesIO(self.data))
        baseline = zipfile.ZipFile(explorer_build.BASELINE)
        self.assertEqual(sorted(built.namelist()), sorted(baseline.namelist()))
        for name in baseline.namelist():
            if name != "xl/vbaProject.bin":
                self.assertEqual(built.read(name), baseline.read(name), name)
        self.assertNotEqual(built.read("xl/vbaProject.bin"), baseline.read("xl/vbaProject.bin"))

    def test_build_is_reproducible(self):
        self.assertEqual(explorer_build.build()[0], self.data)

    def test_ribbon_callbacks_exist_in_the_modules(self):
        ribbon = zipfile.ZipFile(explorer_build.BASELINE).read("customUI/customUI14.xml").decode()
        callbacks = set(re.findall(r'(?:onAction|getEnabled|getPressed|onLoad)="(\w+)"', ribbon))
        source = "\n".join(path.read_text(encoding="latin-1") for path in
                           (DEMO / "M_STATS_PROBDIST_DEMO.bas", DEMO / "M_STATS_PROBDIST_EXPLORER.bas"))
        for callback in callbacks:
            self.assertRegex(source, rf"Public Sub {callback}\(", callback)

    def test_explorer_module_holds_no_fixed_layout_or_numbers(self):
        text = (DEMO / "M_STATS_PROBDIST_EXPLORER.bas").read_text(encoding="latin-1")
        for stale in ("F18", "218", "K_STATS_", "G11 + ", "$MA$1"):
            self.assertNotIn(stale, text)
        self.assertIn('Range("ExplorerFamily")', text)

    def test_workbook_events_route_to_the_explorer_module(self):
        self.assertIn("M_STATS_PROBDIST_EXPLORER.ExplorerInputsChanged Target", explorer_build.EXPLORER_EVENTS)
        self.assertIn("M_STATS_PROBDIST_EXPLORER.ExplorerOpened", explorer_build.EXPLORER_EVENTS)


if __name__ == "__main__":
    unittest.main()
