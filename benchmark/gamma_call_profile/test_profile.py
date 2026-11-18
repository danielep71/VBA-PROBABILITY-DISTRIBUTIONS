"""
#348: tests of the profile's tools. The grid is checked by shape, the
analysis on rows made up here, and the copies of production constants in the
exporter against their source.
"""
import os
import re
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, HERE)
import analyze  # noqa: E402
import build_grid  # noqa: E402


def text_of(*parts):
    with open(os.path.join(ROOT, *parts), encoding="latin-1") as handle:
        return handle.read()


EXPORTER = text_of("benchmark", "gamma_call_profile", "M_STATS_PROBDIST_GCP.bas")
SPECIAL = text_of("src/modules", "M_STATS_PROBDIST_SPECIALFUNCS.bas")
CONTINUOUS = text_of("src/modules", "M_STATS_PROBDIST_CONTINUOUS.bas")


def constant(text, name):
    found = re.search(rf"Const {name}\s+As \w+ = ([^\s']+)", text)
    assert found, name
    return float(found.group(1).rstrip("#"))


def body(text, name):
    start = re.search(rf"^(?:Public|Private) Function {name}\(", text, re.M).start()
    lines = text[start:text.index("End Function", start)].splitlines()
    return [" ".join(line.split()) for line in lines
            if line.strip() and not line.strip().startswith("'")]


def row(part, us, note="ok=1", kind="call", a="2.5", x="1", calls="8"):
    return {"point_id": "x", "kind": kind, "part": part, "a": a, "x": x,
            "calls": calls, "us_per_call": us, "note": note}


class Grid(unittest.TestCase):
    def test_design(self):
        kinds = {}
        for kind, *_ in build_grid.rows():
            kinds[kind] = kinds.get(kind, 0) + 1
        self.assertEqual(kinds, {"call": 96, "fixed": 6, "inverse": 30})

    def test_points_are_the_cost_benchmark(self):
        bridge = text_of("ci", "Run-ExcelVbaTests.ps1")
        found = re.search(r"Points = Array\(([^)]*?10100#)\)", bridge, re.S)
        numbers = [float(v) for v in re.findall(r"[0-9.]+", found.group(1).replace("_", ""))]
        self.assertEqual([(float(a), float(x)) for a, x in build_grid.POINTS],
                         list(zip(numbers[0::2], numbers[1::2])))

    def test_file_is_the_design(self):
        with open(build_grid.FILE, encoding="utf-8") as handle:
            self.assertEqual(build_grid.identity(handle.read()),
                             build_grid.identity(build_grid.render()))

    def test_every_part_has_a_code(self):
        known = set(re.findall(r'Case "(\w+)": PartCode', EXPORTER)) | {"solver"}
        self.assertEqual({part for _, part, _, _ in build_grid.rows()} - known, set())


class Copies(unittest.TestCase):
    """The exporter's copies of private production constants and routines."""

    def test_constants(self):
        for copy, source in (("GCP_INV_MAX_ITER", "PROB_INV_MAX_ITER"),
                             ("GCP_FAST_MIN_SHAPE", "PROB_GAMMAINV_FAST_MIN_SHAPE"),
                             ("GCP_FAST_MAX_SHAPE", "PROB_GAMMA_FAST_MAX_SHAPE"),
                             ("GCP_FAST_FLOOR", "PROB_GAMMA_FAST_FLOOR"),
                             ("GCP_FAST_LOWER_MAX", "PROB_GAMMA_FAST_LOWER_MAX"),
                             ("GCP_FAST_UPPER_MAX", "PROB_GAMMA_FAST_UPPER_MAX")):
            with self.subTest(copy=copy):
                self.assertEqual(constant(EXPORTER, copy), constant(SPECIAL, source))

    def test_replica_is_the_solver_loop(self):
        """Every statement of the production loop is in the replica, in order."""
        def loop(lines):
            start = next(i for i, line in enumerate(lines) if line.startswith("For IterIdx"))
            return lines[start + 1:lines.index("Next IterIdx")]

        rename = {"PROB_INV_MAX_ITER": "GCP_INV_MAX_ITER",
                  "PROB_GAMMAINV_FAST_MIN_SHAPE": "GCP_FAST_MIN_SHAPE"}
        production = []
        for line in loop(body(SPECIAL, "PROB_TryGammaInvP")):
            for old, new in rename.items():
                line = line.replace(old, new)
            production.append(line)
        replica = loop(body(EXPORTER, "ReplicaInverse"))

        # The replica stores each iterate before it is used, and returns
        # where the production loop marks convergence and leaves.
        self.assertEqual(replica[:2], ["Stored = Stored + 1", "Xs(Stored) = X"])
        settle = "If Abs(XNew - X) <= PROB_EPS * Abs(XNew) Or XNew = X Then"
        at = production.index(settle)
        self.assertEqual(production[at + 1:at + 4], ["X = XNew", "Converged = True", "Exit For"])
        at_replica = replica.index(settle)
        self.assertEqual(replica[at_replica + 1:at_replica + 4],
                         ["Count = Stored", "ReplicaInverse = XNew", "Exit Function"])
        self.assertEqual(replica[2:at_replica + 1] + replica[at_replica + 4:],
                         production[:at + 1] + production[at + 4:])

    def test_replica_is_the_solver_seed(self):
        """The statements before the loop, from the orientation on."""
        def seed(lines):
            start = lines.index("If Probability <= ComplementProbability Then")
            stop = next(i for i, line in enumerate(lines) if line.startswith("For IterIdx"))
            return lines[start:stop]

        self.assertEqual(seed(body(EXPORTER, "ReplicaInverse")),
                         seed(body(SPECIAL, "PROB_TryGammaInvP")))

    def test_private_names_exist(self):
        for name in re.findall(r'"M_STATS_PROBDIST_(?:SPECIALFUNCS|CONTINUOUS)\.(\w+)"', EXPORTER):
            with self.subTest(name=name):
                self.assertRegex(SPECIAL + CONTINUOUS,
                                 rf"(?m)^Private Function {name}\(")


class Analysis(unittest.TestCase):
    def rows(self, **change):
        out = [row("empty", "0.100", kind="fixed", a="1", x="1"),
               row("run_pow2", "20.100", kind="fixed", a="1", x="1")]
        out += [row(part, us) for part, us in (
            ("public_cdf", "50.100"), ("public_sf", "60.100"), ("routed_p", "30.100"),
            ("routed_q", "35.100"), ("b64_kernel", "25.100"), ("legacy_loop", "5.100"),
            ("kernel_p", "900.100"), ("kernel_q", "950.100"), ("log_pdf", "2.100"),
            ("run_factor", "35.100"), ("run_validate", "22.100"), ("stirling", "1.100"))]
        out += [row("public_inv", "400.100", kind="inverse", x="0.1"),
                row("inv_p", "350.100", kind="inverse", x="0.1"),
                row("solver", "200.000", "same=1;lower=1;series=4;fraction=1;kernel=0;pdf_us=20.000",
                    kind="inverse", x="0.1", calls="5")]
        for item in out:
            if item["part"] in change:
                item.update(change[item["part"]])
        return out

    def test_complete_export_passes(self):
        self.assertEqual(analyze.problems(self.rows()), [])

    def test_blank_and_bad_rows_fail(self):
        self.assertEqual(len(analyze.problems(self.rows(stirling={"note": ""}))), 1)
        self.assertEqual(len(analyze.problems(self.rows(routed_p={"note": "ok=0"}))), 1)
        self.assertEqual(len(analyze.problems(self.rows(
            solver={"note": "same=0;lower=1;series=4;fraction=1;kernel=0;pdf_us=1.000"}))), 1)
        self.assertEqual(analyze.problems(self.rows(run_factor={"note": "ok=0"})), [])

    def test_layers_and_differences(self):
        text = "\n".join(analyze.tables(self.rows()))
        self.assertIn("| 2.5 | 1 | 50.00 | 60.00 | 30.00 | 35.00 | 25.00 | 5.00 | 900.00 | 950.00 |", text)
        # wrapper 50-30 and 60-35; route test 30-25 (X below A + 1); kernel 25-5;
        # prefactor 35.1-0.1-20 and validation 22.1-0.1-20 through Run
        self.assertIn("| 2.5 | 1 | 20.00 | 25.00 | 5.00 | 20.00 | 15.00 | 2.00 | 1.00 |", text)

    def test_refused_run_part_is_a_dash(self):
        text = "\n".join(analyze.tables(self.rows(run_factor={"note": "ok=0"})))
        self.assertIn("| 2.5 | 1 | 20.00 | 25.00 | 5.00 | 20.00 | - | 2.00 | 1.00 |", text)

    def test_inverse_row(self):
        text = "\n".join(analyze.tables(self.rows()))
        self.assertIn("| 2.5 | 0.1 | 400.0 | 350.0 | 5 | 4 | 1 | 0 | 200.0 | 20.0 | 130.0 |", text)


if __name__ == "__main__":
    unittest.main()
