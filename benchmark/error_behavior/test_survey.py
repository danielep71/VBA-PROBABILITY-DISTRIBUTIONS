"""
#331: tests of the survey's tools: the grid against the public surface, and
the exporter's generated dispatch against the same.
"""
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_grid  # noqa: E402


def exporter():
    with open(os.path.join(HERE, "M_STATS_PROBDIST_ERB.bas"), encoding="latin-1") as handle:
        return handle.read().splitlines()


class Grid(unittest.TestCase):
    def test_every_public_function_has_a_valid_call_and_every_argument_every_class(self):
        declared = build_grid.declarations()
        rows = build_grid.rows()
        self.assertEqual([name for name, label, _, _ in rows if label == "valid"],
                         [name for name, _ in declared])
        for name, parameters in declared:
            for parameter in parameters:
                labels = {label for n, label, p, _ in rows if n == name and p == parameter}
                expected = {label for label, _ in build_grid.CLASSES}
                if parameter in build_grid.PROBABILITIES:
                    expected |= {label for label, _ in build_grid.PROBABILITY_CLASSES}
                with self.subTest(function=name, parameter=parameter):
                    self.assertEqual(labels, expected)

    def test_every_parameter_name_has_a_valid_value(self):
        names = {p for _, parameters in build_grid.declarations() for p in parameters}
        self.assertEqual(names - set(build_grid.VALID), set())

    def test_no_call_takes_more_arguments_than_the_exporter_passes(self):
        self.assertLessEqual(max(len(p) for _, p in build_grid.declarations()), 4)

    def test_file_is_the_design(self):
        with open(build_grid.FILE, encoding="utf-8") as handle:
            self.assertEqual(build_grid.identity(handle.read()),
                             build_grid.identity(build_grid.render()))


class Dispatch(unittest.TestCase):
    def test_exporter_holds_the_generated_cases(self):
        lines = exporter()
        start = lines.index("    Select Case FunctionName") + 1
        stop = lines.index("        Case Else", start)
        self.assertEqual(lines[start:stop], build_grid.dispatch_lines())


if __name__ == "__main__":
    unittest.main()
