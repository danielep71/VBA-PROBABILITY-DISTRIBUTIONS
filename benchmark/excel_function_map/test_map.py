"""
#330: tests of the Excel function map. The map is checked against the
public surface, the references against known values, and the renderer on
observations made up here.
"""
import csv
import io
import os
import re
import struct
import sys
import unittest

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, HERE)
import build_grid  # noqa: E402
import function_map  # noqa: E402
import references  # noqa: E402
import render_map  # noqa: E402


def public_api():
    """{name: number of required arguments, number of arguments} without Status."""
    out = {}
    with open(os.path.join(ROOT, "docs", "PUBLIC_API.txt"), encoding="utf-8") as handle:
        for line in handle:
            if line.startswith("#") or "\t" not in line:
                continue
            name, declaration = line.rstrip("\n").split("\t")[2:4]
            arguments = [a for a in re.search(r"\((.*)\)", declaration).group(1).split(", ")
                         if a and "Status As String" not in a]
            out[name] = (sum(not a.startswith("Optional") for a in arguments), len(arguments))
    return out


def calls(expression):
    """[(name, number of arguments)] of the library calls in an expression."""
    out = []
    for found in re.finditer(r"(K_STATS_\w+)\(", expression):
        depth, count, index = 1, 1, found.end()
        while depth:
            char = expression[index]
            depth += char == "("
            depth -= char == ")"
            count += char == "," and depth == 1
            index += 1
        out.append((found.group(1), count))
    return out


def bits(number):
    return f"{struct.unpack('<Q', struct.pack('<d', float(number)))[0]:016X}"


def grid(**change):
    """The design with both sides returning the reference, then changed by formula."""
    rows = list(csv.DictReader(io.StringIO(build_grid.render())))
    for row in rows:
        row["observed_excel"] = row["observed_kstats"] = bits(mp.mpf(row["reference"]))
        for formula, cells in change.items():
            if row["excel_formula"].replace(";", ",").startswith(formula.replace("_", ".")):
                row.update(cells)
    return rows


class Map(unittest.TestCase):
    def test_every_library_call_is_public_and_takes_its_arguments(self):
        api = public_api()
        expressions = [row["kstats"] for row in function_map.ROWS]
        expressions += [row["example"][1] for row in function_map.ROWS]
        expressions += [p[key] for p in function_map.PROBES for key in ("kstats", "truncated")]
        for expression in expressions:
            found = calls(expression)
            self.assertTrue(found, expression)
            for name, count in found:
                with self.subTest(expression=expression):
                    self.assertIn(name, api)
                    self.assertGreaterEqual(count, api[name][0])
                    self.assertLessEqual(count, api[name][1])

    def test_example_uses_the_functions_of_its_row(self):
        for row in function_map.ROWS:
            with self.subTest(excel=row["excel"]):
                self.assertEqual({name for name, _ in calls(row["kstats"])},
                                 {name for name, _ in calls(row["example"][1])})
                excel_name = re.search(r"[A-Z][A-Z0-9.]+(?=\()", row["excel"]).group(0)
                self.assertIn(excel_name + "(", row["example"][0])

    def test_every_family_with_an_excel_function_is_mapped(self):
        families = {name.split("_")[2] for name in public_api()}
        mapped = {name.split("_")[2] for row in function_map.ROWS for name, _ in calls(row["kstats"])}
        self.assertEqual(families - mapped, {"Uniform", "DiscreteUniform", "Geometric"})
        text = " ".join(functions for _, functions in function_map.LIBRARY_ONLY)
        for family in ("Uniform", "discrete uniform", "geometric"):
            self.assertIn(family, text)

    def test_no_example_is_repeated(self):
        formulas = [row["example"][0] for row in function_map.ROWS]
        self.assertEqual(len(set(formulas)), len(formulas))

    def test_file_is_the_design(self):
        with open(build_grid.FILE, encoding="utf-8") as handle:
            self.assertEqual(build_grid.identity(handle.read()),
                             build_grid.identity(build_grid.render()))


class References(unittest.TestCase):
    """Values known from tables or closed forms."""

    def close(self, expression, expected, digits=12):
        self.assertAlmostEqual(float(references.evaluate(expression)), expected, digits)

    def test_known_values(self):
        self.close("K_STATS_NormalStandard_Cumulative(0)", 0.5)
        self.close("K_STATS_NormalStandard_InverseCumulative(0.975)", 1.959963984540054)
        self.close("K_STATS_StudentT_InverseCumulative(0.975,10)", 2.2281388519649, 9)
        self.close("K_STATS_ChiSquare_InverseCumulative(0.95,4)", 9.487729036781154, 10)
        self.close("K_STATS_ChiSquare_Cumulative(2,2)", 1 - mp.e ** -1)
        self.close("K_STATS_Exponential_Cumulative(1,1)", 1 - mp.e ** -1)
        self.close("K_STATS_Beta_Cumulative(0.5,2,2)", 0.5)
        self.close("K_STATS_F_Cumulative(1,7,7)", 0.5)
        self.close("K_STATS_Gamma_Cumulative(3,1,1.5)", 1 - mp.e ** -2)
        self.close("K_STATS_Binomial_Cumulative(1,2,0.5)", 0.75)
        self.close("K_STATS_Poisson_PMF(0,2)", mp.e ** -2)
        self.close("K_STATS_Hypergeometric_PMF(1,1,1,2)", 0.5)
        self.close("K_STATS_NegativeBinomial_PMF(0,3,0.5)", 0.125)
        self.close("K_STATS_Weibull_Cumulative(1.5,1,1.5)", 1 - mp.e ** -1)

    def test_tails_add_to_one(self):
        for lower, upper in (("K_STATS_StudentT_Cumulative(1.5,10)", "K_STATS_StudentT_Survival(1.5,10)"),
                             ("K_STATS_ChiSquare_Cumulative(3,4)", "K_STATS_ChiSquare_Survival(3,4)"),
                             ("K_STATS_F_Cumulative(2.5,5,10)", "K_STATS_F_Survival(2.5,5,10)")):
            self.assertAlmostEqual(float(references.evaluate(lower) + references.evaluate(upper)), 1.0, 14)

    def test_inverses_invert(self):
        for inverse, forward in (
                ("K_STATS_Gamma_InverseCumulative(0.9,2,1.5)", "K_STATS_Gamma_Cumulative({},2,1.5)"),
                ("K_STATS_Beta_InverseCumulative(0.4,2,3)", "K_STATS_Beta_Cumulative({},2,3)"),
                ("K_STATS_F_InverseCumulative(0.95,5,10)", "K_STATS_F_Cumulative({},5,10)")):
            x = references.evaluate(inverse)
            p = references.NAMES[forward.split("(")[0]](x, *[float(v) for v in
                                                           forward.split("{},")[1].rstrip(")").split(",")])
            self.assertAlmostEqual(float(p), float(inverse.split("(")[1].split(",")[0]), 14)


class Render(unittest.TestCase):
    def test_agreement_is_same(self):
        text = render_map.render(grid())
        self.assertIn("| 0.933192798731 | same |", text)
        self.assertNotIn("off by", text)

    def test_excel_difference_is_shown(self):
        text = render_map.render(grid(**{"1-NORM_S_DIST": {"observed_excel": bits(0.0)}}))
        self.assertIn("| 1.12858840595e-19 | 0: the value is lost |", text)
        text = render_map.render(grid(**{"GAUSS": {"observed_excel": bits(0.4331)}}))
        self.assertIn("| 0.433192798731 | 0.4331, off by 0.00021 |", text)

    def test_excel_error_is_shown(self):
        text = render_map.render(grid(**{"GAUSS": {"observed_excel": "ERROR"}}))
        self.assertIn("| 0.433192798731 | an error |", text)

    def test_library_off_its_reference_is_refused(self):
        with self.assertRaises(render_map.MapError):
            render_map.render(grid(**{"GAUSS": {"observed_kstats": bits(0.4331927988)}}))
        with self.assertRaises(render_map.MapError):
            render_map.render(grid(**{"GAUSS": {"observed_kstats": "ERROR"}}))

    def test_blank_grid_is_refused(self):
        with self.assertRaises(render_map.MapError):
            render_map.render(list(csv.DictReader(io.StringIO(build_grid.render()))))

    def test_probe_findings(self):
        fraction, truncated = mp.mpf("0.35"), mp.mpf("0.44")
        self.assertEqual(render_map.probe_finding(bits(0.35), fraction, truncated), "uses the fraction")
        self.assertEqual(render_map.probe_finding(bits(0.44), fraction, truncated), "rounds it down")
        self.assertEqual(render_map.probe_finding("ERROR", fraction, truncated), "refuses it")
        self.assertEqual(render_map.probe_finding(bits(0.44), truncated, truncated), "rounds it down")
        self.assertIn("neither", render_map.probe_finding(bits(0.5), fraction, truncated))

    def test_section_is_replaced_between_markers(self):
        page = f"a\n{render_map.BEGIN}\nold\n{render_map.END}\nb\n"
        self.assertEqual(render_map.replace_section(page, "NEW"), "a\nNEW\nb\n")
        with self.assertRaises(render_map.MapError):
            render_map.replace_section("no markers", "NEW")


if __name__ == "__main__":
    unittest.main()
