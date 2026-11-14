"""
#35: tests of the production holdout's tools. No holdout row is evaluated:
the grid is checked by shape, and the rules on rows made up here.
"""
import os
import struct
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import analyze  # noqa: E402
import build_grid  # noqa: E402


def bits(x):
    return f"{struct.unpack('<Q', struct.pack('<d', x))[0]:016X}"


def row(value, ref="2", kind="holdout_f", func="F", point="x000"):
    cell = value if isinstance(value, str) else bits(value)
    return {"point_id": point, "kind": kind, "func": func, "p_bits": bits(0.5),
            "a_bits": bits(3.0), "b_bits": bits(4.0), "ref": ref, "value": cell}


def rules(base, cand):
    failures, _ = analyze.analyze([base], [cand])
    return sorted({text.split()[0] for text in failures})


class Grid(unittest.TestCase):
    def test_design(self):
        rows = build_grid.points()
        kinds = {}
        for kind, *_ in rows:
            kinds[kind] = kinds.get(kind, 0) + 1
        self.assertEqual(kinds, {"holdout_f": 16, "far_f": 10, "seam_f": 10,
                                 "t_invariance": 66, "beta": 24})
        self.assertEqual(len(set(rows)), len(rows))

    def test_holdout_is_the_reserved_one(self):
        reserved = {(p, d) for p in (0.1, 0.4, 0.6, 0.98)
                    for d in ((3e5, 7.0), (7.0, 3e5), (50.0, 1e8), (1e8, 50.0))}
        found = {(p, (a, b)) for kind, _, p, a, b in build_grid.points() if kind == "holdout_f"}
        self.assertEqual(found, reserved)

    def test_files_are_blank_and_equal(self):
        texts = []
        for path in build_grid.FILES.values():
            with open(path, encoding="utf-8") as handle:
                texts.append(build_grid.identity(handle.read()))
        self.assertEqual(texts[0], texts[1])
        self.assertEqual(len(texts[0]), 127)


class Rules(unittest.TestCase):
    def test_equal_passes(self):
        self.assertEqual(rules(row(2.0), row(2.0)), [])

    def test_blank_and_fault(self):
        self.assertEqual(rules(row(""), row(2.0)), ["R0"])
        self.assertEqual(rules(row(2.0), row("ERROR")), ["R0"])
        self.assertEqual(analyze.analyze([row("")])[0][0].split()[0], "R0")
        self.assertEqual(analyze.analyze([row(2.0)])[0], [])

    def test_regression_above_the_floor(self):
        self.assertEqual(rules(row(2.0), row(2.0 * (1 + 5e-14))), ["R1"])
        self.assertEqual(rules(row(2.0), row(2.0 * (1 + 5e-15))), [])
        self.assertEqual(rules(row(2.0 * (1 + 1e-9)), row(2.0 * (1 + 1e-13))), [])

    def test_guarantee(self):
        self.assertEqual(rules(row(2.0 * (1 + 1e-9)), row(2.0 * (1 + 1e-11))), ["R2"])
        self.assertEqual(rules(row("CVERR"), row(2.0 * (1 + 1e-11))), ["R2"])
        self.assertEqual(rules(row("CVERR"), row(2.0)), [])
        self.assertEqual(rules(row("CVERR"), row("CVERR")), [])

    def test_domain(self):
        self.assertEqual(rules(row(2.0), row("CVERR")), ["R3"])

    def test_shared_path(self):
        t = dict(kind="t_invariance", func="T")
        up = lambda n: build_grid.step(2.0, n)
        self.assertEqual(rules(row(2.0, **t), row(up(1), **t)), [])
        self.assertEqual(rules(row(2.0, **t), row(up(3), **t)), ["R4"])
        self.assertEqual(rules(row(up(9), **t), row(up(3), **t)), [])
        self.assertEqual(rules(row(2.0), row(up(3))), [])

    def test_zero_reference_is_absolute(self):
        t = dict(kind="t_invariance", func="T", ref="0")
        self.assertEqual(rules(row(0.0, **t), row(0.0, **t)), [])
        self.assertEqual(rules(row(0.0, **t), row(1e-3, **t)), ["R1", "R2", "R4"])

    def test_different_rows(self):
        self.assertEqual(rules(row(2.0), row(2.0, point="x001")), ["R0:"])

    def test_ordinal(self):
        self.assertEqual(analyze.ordinal(bits(0.0)), analyze.ordinal(bits(-0.0)))
        self.assertEqual(analyze.ordinal(bits(5e-324)) - analyze.ordinal(bits(-5e-324)), 2)


if __name__ == "__main__":
    unittest.main()
