"""The user manual names only functions the library declares (#263)."""
from __future__ import annotations

import importlib.util
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MANUAL = ROOT / "docs" / "user-manual" / "v1.0.0"


def load():
    spec = importlib.util.spec_from_file_location("build_manual", MANUAL / "build_manual.py")
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class Manual(unittest.TestCase):
    def test_every_name_and_call_matches_the_public_api(self) -> None:
        self.assertEqual(load().problems(), [])

    def test_the_api_baseline_is_read(self) -> None:
        self.assertEqual(len(load().declarations()), 112)

    def test_a_wrong_argument_count_is_found(self) -> None:
        manual = load()
        self.assertEqual(list(manual.calls("=K_STATS_Gamma_Mean(3,2) and K_STATS_Gamma_Mean(...)")),
                         [("K_STATS_Gamma_Mean", 2)])
        self.assertEqual(list(manual.calls("K_STATS_F_Density(1,(2+3),4)")), [("K_STATS_F_Density", 3)])

    def test_the_chapters_are_numbered_in_order(self) -> None:
        chapters = load().chapters()
        self.assertEqual(len(chapters), 14)
        for number, path in enumerate(chapters):
            first = path.read_text(encoding="utf-8").splitlines()[0]
            expected = "# About this manual" if number == 0 else f"# {number}. "
            self.assertTrue(first.startswith(expected), f"{path.name}: {first!r}")


class ExcelExamples(unittest.TestCase):
    """Test_DS_ManualExamples asserts in Excel the figures the manual prints."""

    def group(self) -> str:
        source = (ROOT / "tests" / "M_STATS_PROBDIST_TEST.bas").read_text(encoding="ascii")
        start = source.index("Private Sub Test_DS_ManualExamples()")
        return source[start:source.index("End Sub", start)]

    def test_the_group_is_called(self) -> None:
        import re
        source = (ROOT / "tests" / "M_STATS_PROBDIST_TEST.bas").read_text(encoding="ascii")
        self.assertEqual(len(re.findall(r"(?m)^    Test_DS_ManualExamples\s*$", source)), 1)

    def test_every_asserted_figure_is_printed_in_the_manual(self) -> None:
        import re
        manual = "\n".join(path.read_text(encoding="utf-8") for path in load().chapters())
        code = "\n".join(line for line in self.group().splitlines() if not line.lstrip().startswith("'"))
        figures = set(re.findall(r"(?<![\w.])-?\d+\.\d{6,}(?:E-\d+)?(?![\w.])", code))
        figures.discard("0.0000000006")                # the tolerance
        figures.discard("1.64485362695147")            # an argument, printed in section 4.3
        self.assertGreaterEqual(len(figures), 30)
        missing = sorted(f for f in figures if f not in manual)
        self.assertEqual(missing, [])

    def test_the_first_formula_argument_is_the_one_printed(self) -> None:
        manual = "\n".join(path.read_text(encoding="utf-8") for path in load().chapters())
        self.assertIn("K_STATS_NormalStandard_Cumulative(1.64485362695147)", manual)
        self.assertIn("K_STATS_NormalStandard_Cumulative(1.64485362695147)", self.group())


if __name__ == "__main__":
    unittest.main()
