"""
#359: tests of the cost study's tools. No holdout row is evaluated: the grid
is checked by shape, and the rules on rows and cost lines made up here.
"""
import os
import struct
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import analyze  # noqa: E402
import build_grid  # noqa: E402
import fit_sweep  # noqa: E402


def bits(x):
    return f"{struct.unpack('<Q', struct.pack('<d', x))[0]:016X}"


def row(value, ref="2", kind="direct_beta", func="Beta", point="x000"):
    cell = value if isinstance(value, str) else bits(value)
    return {"point_id": point, "kind": kind, "func": func, "p_bits": bits(0.5),
            "a_bits": bits(3.0), "b_bits": bits(4.0), "ref": ref, "value": cell}


def rules(base, cand):
    failures, _ = analyze.analyze([base], [cand])
    return sorted({text.split()[0] for text in failures})


def costs(scale=1.0, **override):
    """Every benchmark line at its recorded time after #35, times scale."""
    out = {key: (after * 1000.0 * scale, True) for key, (_, after) in analyze.RECORDED.items()}
    out.update({key: (30.0, True) for key in analyze.REPORTED})
    out.update(override.get("lines", {}))
    return out


def cost_rules(base, cand):
    return sorted({text.split()[0] for text in analyze.analyze_cost(base, cand)[0]})


class Grid(unittest.TestCase):
    def test_design(self):
        rows = build_grid.points()
        kinds = {}
        for kind, *_ in rows:
            kinds[kind] = kinds.get(kind, 0) + 1
        self.assertEqual(kinds, {"direct_beta": 42, "small_shape": 20, "balanced_f": 36,
                                 "large_f": 36, "branch_seam": 10, "half_seam": 10,
                                 "student_t": 24})
        self.assertEqual(len(set(rows)), len(rows))

    def test_no_row_was_seen_before(self):
        """Against the fitting sweep, the production holdout of #35 and the
        cost benchmark, with an F row read as the Beta problem it solves."""
        def key(func, p, a, b):
            return ("T", p, a, None) if func == "T" else \
                ("Beta", p, a / 2.0, b / 2.0) if func == "F" else ("Beta", p, a, b)

        seen = {key("Beta", p, a, b) for a in fit_sweep.BETA_SHAPES for b in fit_sweep.BETA_SHAPES
                for p in fit_sweep.PROBABILITIES}
        seen |= {key("F", p, float(a), float(b)) for a, b in fit_sweep.F_DF
                 for p in fit_sweep.PROBABILITIES}
        seen |= {key(func, p, a, b) for _, func, p, a, b in build_grid.production.points()}
        seen |= {key("Beta", p, a, b) for a, b in ((0.5, 0.5), (2.0, 5.0), (50.0, 80.0), (0.7, 25.0))
                 for p in (0.1, 0.9)}
        seen |= {key("F", p, a, b) for a, b in ((5.0, 10.0), (100.0, 100.0), (50.0, 1e8), (1e5, 2.0))
                 for p in (0.1, 0.9)}
        seen |= {key("T", 0.9, df, None) for df in (3.0, 10.0, 30.0)}
        here = {key(func, p, a, b) for _, func, p, a, b in build_grid.points()}
        self.assertEqual(here & seen, set())

    def test_files_hold_the_same_rows(self):
        texts = []
        for path in build_grid.FILES.values():
            with open(path, encoding="utf-8") as handle:
                texts.append(build_grid.production.identity(handle.read()))
        self.assertEqual(texts[0], texts[1])
        self.assertEqual(len(texts[0]) - 1, len(build_grid.points()))


class Rules(unittest.TestCase):
    def test_equal_passes(self):
        self.assertEqual(rules(row(2.0), row(2.0)), [])

    def test_blank_and_fault(self):
        self.assertEqual(rules(row(""), row(2.0)), ["R0"])
        self.assertEqual(rules(row(2.0), row("ERROR")), ["R0"])
        self.assertEqual(analyze.analyze([row("")])[0][0].split()[0], "R0")
        self.assertEqual(analyze.analyze([row(2.0)])[0], [])

    def test_regression_is_measured_from_the_baseline(self):
        self.assertEqual(rules(row(2.0), row(2.0 * (1 + 2e-14))), [])
        self.assertEqual(rules(row(2.0), row(2.0 * (1 + 5e-14))), ["R1"])
        self.assertEqual(rules(row(2.0 * (1 + 4e-14)), row(2.0 * (1 + 6e-14))), [])
        self.assertEqual(rules(row(2.0 * (1 + 1e-9)), row(2.0 * (1 + 1e-13))), [])

    def test_guarantee_where_the_baseline_holds_it(self):
        self.assertEqual(rules(row(2.0), row(2.0 * (1 + 1e-11))), ["R1", "R2", "R4"])
        self.assertEqual(rules(row(2.0 * (1 + 1e-9)), row(2.0 * (1 + 1e-11))), [])
        self.assertEqual(rules(row("CVERR"), row(2.0 * (1 + 1e-11))), [])
        self.assertEqual(rules(row("CVERR"), row("CVERR")), [])

    def test_worst_where_the_baseline_holds_it(self):
        self.assertEqual(rules(row(2.0 * (1 + 9e-14)), row(2.0 * (1 + 1.1e-13))), ["R4"])
        self.assertEqual(rules(row(2.0 * (1 + 2e-13)), row(2.0 * (1 + 2.1e-13))), [])

    def test_domain(self):
        self.assertEqual(rules(row(2.0), row("CVERR")), ["R3"])

    def test_different_rows(self):
        self.assertEqual(rules(row(2.0), row(2.0, point="x001")), ["R0:"])


class Cost(unittest.TestCase):
    DIRECT = ("beta_inv", "2", "5", "0.9")      # before 0.024 ms, after 0.591 ms
    PAIRED = ("beta_inv", "50", "80", "0.9")

    def test_points_are_split_twelve_and_four(self):
        self.assertEqual((len(analyze.DIRECT), len(analyze.PAIR)), (12, 4))

    def test_the_time_before_35_passes_everywhere_direct(self):
        cand = {key: (before * 1000.0, True) for key, (before, _) in analyze.RECORDED.items()}
        cand.update({key: (30.0, True) for key in analyze.REPORTED})
        self.assertEqual(cost_rules(costs(), cand), [])

    def test_unchanged_cost_fails_both(self):
        self.assertEqual(cost_rules(costs(), costs()), ["C1", "C2"])

    def test_direct_limit_is_five_times_before(self):
        fast = {key: (1.0, True) for key in analyze.RECORDED}
        fast.update({key: (30.0, True) for key in analyze.REPORTED})
        self.assertEqual(cost_rules(costs(), {**fast, self.DIRECT: (119.0, True)}), [])
        self.assertEqual(cost_rules(costs(), {**fast, self.DIRECT: (121.0, True)}), ["C1"])

    def test_direct_limit_follows_the_baseline_run(self):
        fast = {key: (1.0, True) for key in analyze.RECORDED}
        fast.update({key: (30.0, True) for key in analyze.REPORTED})
        cand = {**fast, self.DIRECT: (200.0, True)}
        self.assertEqual(cost_rules(costs(), cand), ["C1"])
        self.assertEqual(cost_rules(costs(scale=2.0), cand), [])

    def test_pair_limit_is_half_the_baseline(self):
        fast = {key: (1.0, True) for key in analyze.RECORDED}
        fast.update({key: (30.0, True) for key in analyze.REPORTED})
        self.assertEqual(cost_rules(costs(), {**fast, self.PAIRED: (3300.0, True)}), [])
        self.assertEqual(cost_rules(costs(), {**fast, self.PAIRED: (3400.0, True)}), ["C2"])

    def test_missing_and_failed_lines(self):
        base = costs()
        del base[self.DIRECT]
        self.assertEqual(cost_rules(base, None), ["C0"])
        self.assertEqual(cost_rules(costs(), costs(lines={self.PAIRED: (1.0, False)})), ["C0"])
        self.assertEqual(cost_rules(costs(), None), [])

    def test_log_lines_are_read(self):
        path = os.path.join(HERE, "_cost_lines_test.txt")
        with open(path, "w", encoding="utf-8") as handle:
            handle.write("2026-11-16 COST f_inv A=50 B=100000000 X=0.9 ok=1 calls=64 us_per_call=5463.000\n"
                         "COST t_inv A=3 X=0.9 ok=1 calls=32768 us_per_call=25.125\n"
                         "COST gamma_inv A=0.5 X=0.1 ok=1 calls=8 us_per_call=9.000\n")
        try:
            found = analyze.read_cost(path)
        finally:
            os.remove(path)
        self.assertEqual(found, {("f_inv", "50", "100000000", "0.9"): (5463.0, True),
                                 ("t_inv", "3", "", "0.9"): (25.125, True)})


class Record(unittest.TestCase):
    """The committed files give the verdict RESULT.md records."""

    def test_accuracy_rules_hold(self):
        failures, report = analyze.analyze(analyze.read(analyze.BASELINE),
                                           analyze.read(analyze.CANDIDATE))
        self.assertEqual(failures, [])
        self.assertEqual(report["counts"]["rows"], len(build_grid.points()))

    def test_cost_fails_on_one_pair_line_only(self):
        failures, _ = analyze.analyze_cost(analyze.read_cost(analyze.COST_BASELINE),
                                           analyze.read_cost(analyze.COST_CANDIDATE))
        self.assertEqual(len(failures), 1, failures)
        self.assertTrue(failures[0].startswith("C2 beta_inv (0.7, 25) p=0.9"), failures)


if __name__ == "__main__":
    unittest.main()
