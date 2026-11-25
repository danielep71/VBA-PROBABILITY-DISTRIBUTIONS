"""#493: tests of the Excel floor tools, on rows made up here."""
import os
import struct
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import floor_build as B  # noqa: E402
import floor_check as C  # noqa: E402


def bits(x):
    return struct.pack(">d", x).hex().upper()


def row(kind, excel, library, ref, formula="F"):
    return {"kind": kind, "excel_formula": formula, "kstats_call": "K", "reference": repr(ref),
            "observed_excel": "ERROR" if excel is None else bits(excel),
            "observed_kstats": "ERROR" if library is None else bits(library)}


class Build(unittest.TestCase):
    def test_token_keeps_only_exact_fifteen_digit_arguments(self):
        self.assertEqual(B.token("1.5"), "1.5")
        self.assertEqual(B.token("0.0"), "0")
        self.assertEqual(B.token("1e-300"), "1E-300")
        self.assertIsNone(B.token("0.30000000000000004"))
        self.assertIsNone(B.token(repr(1.0 / 3.0)))

    def test_every_function_is_public_and_its_formulas_use_its_arguments(self):
        with open(os.path.join(os.path.dirname(os.path.dirname(HERE)), "docs",
                               "PUBLIC_API.txt"), encoding="utf-8") as f:
            api = f.read()
        for fn, (arity, truncated, forms) in B.EXCEL.items():
            self.assertIn(f"K_STATS_{fn}(", api, fn)
            self.assertTrue(all(i < arity for i in truncated), fn)
            for form in forms:
                form.format(**{f"a{i}": "1" for i in range(arity)})

    def test_committed_grid_matches_the_generator(self):
        self.assertEqual(B.main(["--check"]), 0)


class Judge(unittest.TestCase):
    def verdict(self, rows):
        points, blank = C.judge(rows)
        self.assertEqual(blank, 0)
        return points[0]["verdict"]

    def test_inferior_needs_excel_closer_and_the_library_above_the_floor(self):
        ref = 0.25
        self.assertEqual(self.verdict([row("m:2:X", ref, ref * (1 + 1e-13), ref)]), "inferior")
        self.assertEqual(self.verdict([row("m:2:X", ref, ref * (1 + 5e-15), ref)]),
                         "not inferior")
        self.assertEqual(self.verdict([row("m:2:X", ref * (1 + 1e-12), ref * (1 + 1e-13), ref)]),
                         "not inferior")

    def test_the_most_accurate_excel_formulation_is_used(self):
        ref = 0.25
        rows = [row("m:2:X", ref * (1 + 1e-10), ref * (1 + 1e-13), ref, "A"),
                row("m:2:X", ref, ref * (1 + 1e-13), ref, "B")]
        points, _ = C.judge(rows)
        self.assertEqual(points[0]["verdict"], "inferior")
        self.assertEqual(points[0]["excel_formula"], "B")

    def test_errors_are_not_comparisons(self):
        self.assertEqual(self.verdict([row("m:2:X", None, 0.25, 0.25)]), "not compared")
        self.assertEqual(self.verdict([row("m:2:X", 0.25, None, 0.25)]), "library error")

    def test_discrete_inverse_compares_counts(self):
        kind = "m:2:Binomial_InverseCumulative"
        self.assertEqual(self.verdict([row(kind, 3.0, 4.0, 3.0)]), "inferior")
        self.assertEqual(self.verdict([row(kind, 3.0, 3.0, 3.0)]), "not inferior")
        self.assertEqual(self.verdict([row(kind, 4.0, 3.0, 3.0)]), "not inferior")

    def test_blank_rows_are_counted(self):
        r = row("m:2:X", 0.25, 0.25, 0.25)
        r["observed_excel"] = r["observed_kstats"] = ""
        self.assertEqual(C.judge([r])[1], 1)


if __name__ == "__main__":
    unittest.main()
