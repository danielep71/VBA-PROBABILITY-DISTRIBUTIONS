"""
Fixtures for the result binding of the #72 normal-tail study (#79,
PREREGISTRATION.md Amendment 3).

  1. The committed results.json is current: its provenance block names every
     bound file with its role, every SHA-256 matches the file on disk, and the
     candidate matches coefficients_sha256 and holdout_exposure.candidate.
     The validator reads no input that is not bound.
  2. Mutations make the record stale. On a copy of the bound files, each of
     these is reported, naming the file: a change to the generator
     (derive_kernel.py: the iteration cap), the convergence record (its
     status), the models (kernel_model.py: the bound B; the #34 route model),
     the validator (validate_kernel.py: the seed of V; holdout_record.py;
     result_binding.py), the design and configuration (an appended
     preregistration amendment), and the inputs (a reference in a §7 grid, a
     P4 contract threshold, the #34 route points), as well as a changed
     candidate, a deleted file, a missing provenance block, and a missing or
     unknown entry.
  3. The modes fail closed on a stale record: on a synthetic study, --verify
     reports STALE without computing, evaluating or writing; --regenerate
     records the identities that changed and keeps the exposure record and
     A6; --verify then passes. validate_kernel.py --verify reports the same on
     a stale copy of the committed record without computing anything.

It never evaluates the holdout H and writes only in temporary directories.

Run: python3 test_result_binding.py   (exit 0 = pass, nonzero = fail)
"""
import copy
import json
import os
import shutil
import sys
import tempfile
import unittest
from unittest import mock

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import holdout_record as hr                                      # noqa: E402
import result_binding as rb                                      # noqa: E402
from test_holdout_record import Synthetic                        # noqa: E402

with open(os.path.join(HERE, "results.json"), encoding="utf-8") as f:
    RESULT = json.load(f)


class Committed(unittest.TestCase):
    def test_committed_record_is_current(self):
        self.assertEqual(rb.stale(RESULT), [])
        files = RESULT["provenance"]["files"]
        self.assertEqual(list(files), [name for name, _ in rb.BOUND])
        self.assertEqual({k: v["role"] for k, v in files.items()}, dict(rb.BOUND))
        self.assertEqual(RESULT["coefficients_sha256"], rb.file_sha256(os.path.join(rb.BENCH, rb.CANDIDATE)))
        self.assertEqual(RESULT["holdout_exposure"]["candidate"]["coefficients_sha256"],
                         RESULT["coefficients_sha256"])
        self.assertNotIn(rb.CANDIDATE, files, "the candidate is bound once, by coefficients_sha256")

    def test_validator_inputs_are_bound(self):
        import validate_kernel as vk
        bound = {name for name, _ in rb.BOUND}
        for grid in vk.GRIDS:
            self.assertIn(grid, bound)
        for path in (vk.PREREGISTRATION, vk.dk.CONVERGENCE_OUT, vk.km.__file__, vk.rm.__file__, vk.hr.__file__,
                     vk.__file__, vk.dk.__file__, rb.__file__):
            self.assertIn(os.path.relpath(os.path.abspath(path), rb.BENCH).replace(os.sep, "/"), bound)
        for name in ("accuracy_contracts.csv", "student_t_large_df_study/results.json",
                     "student_t_large_df_study/coefficients.json"):
            self.assertIn(name, bound)


class Mutations(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp()
        for name in [n for n, _ in rb.BOUND] + [rb.CANDIDATE]:
            target = os.path.join(self.root, name)
            os.makedirs(os.path.dirname(target), exist_ok=True)
            shutil.copyfile(os.path.join(rb.BENCH, name), target)
        self.assertEqual(rb.stale(RESULT, self.root), [])

    def tearDown(self):
        shutil.rmtree(self.root)

    def replace(self, name, old, new):
        path = os.path.join(self.root, name)
        with open(path, encoding="utf-8") as f:
            text = f.read()
        self.assertIn(old, text, f"mutation anchor missing in {name}")
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            f.write(text.replace(old, new, 1))

    def assert_stale(self, name, result=None):
        problems = rb.stale(RESULT if result is None else result, self.root)
        self.assertEqual(len(problems), 1, problems)
        self.assertTrue(problems[0].startswith(name), problems)
        return problems[0]

    def test_generator(self):
        self.replace("normal_tail_study/derive_kernel.py", "ITERATIONS = 120", "ITERATIONS = 121")
        self.assert_stale("normal_tail_study/derive_kernel.py")

    def test_convergence_record(self):
        self.replace("normal_tail_study/fit_convergence.json", '"status": "CONVERGED"', '"status": "CAP_EXHAUSTED"')
        self.assert_stale("normal_tail_study/fit_convergence.json")

    def test_model(self):
        self.replace("normal_tail_study/kernel_model.py", "BOUND = 2.0 ** -50", "BOUND = 2.0 ** -49")
        self.assert_stale("normal_tail_study/kernel_model.py")

    def test_route_model(self):
        self.replace("student_t_large_df_study/binary64_route_model.py", "ORDER = 5", "ORDER = 4")
        self.assert_stale("student_t_large_df_study/binary64_route_model.py")

    def test_validator(self):
        self.replace("normal_tail_study/validate_kernel.py", "random.Random(72001)", "random.Random(72005)")
        self.assert_stale("normal_tail_study/validate_kernel.py")

    def test_holdout_record_and_binding_code(self):
        for name in ("normal_tail_study/holdout_record.py", "normal_tail_study/result_binding.py"):
            with open(os.path.join(self.root, name), "a", encoding="utf-8") as f:
                f.write("\n")
            self.assert_stale(name)
            shutil.copyfile(os.path.join(rb.BENCH, name), os.path.join(self.root, name))

    def test_design(self):
        with open(os.path.join(self.root, "normal_tail_study/PREREGISTRATION.md"), "a", encoding="utf-8") as f:
            f.write("\n---\n\n## Amendment 4 — synthetic\n")
        self.assert_stale("normal_tail_study/PREREGISTRATION.md")

    def test_grid_input(self):
        path = os.path.join(self.root, "probability_accuracy_grid.csv")
        with open(path, encoding="utf-8") as f:
            lines = f.read().split("\n")
        header = lines[0].split(",")
        row = next(i for i, line in enumerate(lines) if line.startswith("NormalStandard_Survival,"))
        cells = lines[row].split(",")
        reference = cells[header.index("reference")]
        cells[header.index("reference")] = reference[:-1] + ("6" if reference[-1] == "5" else "5")
        lines[row] = ",".join(cells)
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            f.write("\n".join(lines))
        self.assert_stale("probability_accuracy_grid.csv")

    def test_contract_and_route_inputs(self):
        with open(os.path.join(self.root, "accuracy_contracts.csv"), "a", encoding="utf-8") as f:
            f.write("\n")
        self.assert_stale("accuracy_contracts.csv")
        shutil.copyfile(os.path.join(rb.BENCH, "accuracy_contracts.csv"),
                        os.path.join(self.root, "accuracy_contracts.csv"))
        self.replace("student_t_large_df_study/results.json", '"t"', '"t" ')
        self.assert_stale("student_t_large_df_study/results.json")

    def test_candidate(self):
        self.replace(rb.CANDIDATE, '"split": 3.0', '"split": 3.5')
        self.assert_stale(rb.CANDIDATE)

    def test_deleted_file(self):
        os.unlink(os.path.join(self.root, "holdout_older/holdout_older_grid.csv"))
        self.assertIn("missing", self.assert_stale("holdout_older/holdout_older_grid.csv"))

    def test_missing_block_and_entries(self):
        result = copy.deepcopy(RESULT)
        del result["provenance"]
        self.assertIn("no provenance block", rb.stale(result, self.root)[0])
        result = copy.deepcopy(RESULT)
        del result["provenance"]["files"]["accuracy_contracts.csv"]
        self.assertIn("bound but not recorded", self.assert_stale("accuracy_contracts.csv", result))
        result = copy.deepcopy(RESULT)
        result["provenance"]["files"]["normal_tail_study/extra.py"] = {"role": "x", "sha256": "0" * 64}
        self.assertIn("no longer bound", self.assert_stale("normal_tail_study/extra.py", result))
        self.assertEqual(rb.stale([]), ["the record is not a JSON object"])


class Modes(unittest.TestCase):
    """The binding hook of holdout_record.py on a synthetic study."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.s = Synthetic(self.tmp)
        self.code = os.path.join(self.tmp, "generator.py")
        with open(self.code, "w", encoding="utf-8") as f:
            f.write("ITERATIONS = 120\n")
        compute = self.s.compute

        def bound_compute():
            body = compute()
            body["provenance"] = {"files": {"generator.py": {"sha256": hr.file_sha256(self.code)}}}
            return body
        self.s.compute = bound_compute

    def tearDown(self):
        shutil.rmtree(self.tmp)

    def binding(self, record):
        recorded = record.get("provenance", {}).get("files", {}).get("generator.py", {}).get("sha256")
        now = hr.file_sha256(self.code)
        return [] if recorded == now else [f"generator.py: recorded {recorded}, now {now}"]

    def run_mode(self, mode, reason=None):
        study = self.s.study()
        study.binding = self.binding
        return hr.run(study, mode, reason)

    def test_stale_verify_then_regenerate(self):
        self.assertEqual(self.run_mode("init"), hr.OK)
        self.assertEqual(self.run_mode("first_evaluation"), hr.OK)
        self.assertEqual(self.run_mode("verify"), hr.OK)
        with open(self.code, "w", encoding="utf-8") as f:
            f.write("ITERATIONS = 121\n")
        before, computed, evaluated = self.s.bytes(), self.s.computed, len(self.s.evaluated)
        self.assertEqual(self.run_mode("verify"), hr.FAILED)
        self.assertIn("STALE", self.s.log())
        self.assertEqual((self.s.bytes(), self.s.computed, len(self.s.evaluated)), (before, computed, evaluated))
        exposure, a6 = self.s.record()["holdout_exposure"], self.s.record()["acceptance"]["A6"]
        self.assertEqual(self.run_mode("regenerate", "synthetic: generator changed"), hr.OK)
        rec = self.s.record()
        self.assertEqual(len(self.s.evaluated), evaluated, "regenerate evaluated the holdout")
        self.assertEqual((rec["holdout_exposure"], rec["acceptance"]["A6"]), (exposure, a6))
        entry = rec["regenerations"][-1]
        self.assertEqual(entry["holdout_evaluated"], False)
        self.assertEqual(len(entry["identities_changed"]), 1)
        self.assertTrue(entry["identities_changed"][0].startswith("generator.py"))
        self.assertEqual(hr.check_record(rec), [])
        self.assertEqual(self.run_mode("verify"), hr.OK)

    def test_validator_reports_a_stale_committed_record_without_computing(self):
        import validate_kernel as vk
        stale = copy.deepcopy(RESULT)
        stale["provenance"]["files"]["normal_tail_study/validate_kernel.py"]["sha256"] = "0" * 64
        out = os.path.join(self.tmp, "stale_results.json")
        with open(out, "w", encoding="utf-8") as f:
            json.dump(stale, f, indent=1)
            f.write("\n")
        with open(out, "rb") as f:
            raw = f.read()
        boom = mock.Mock(side_effect=AssertionError("must not compute"))
        logs = []
        with mock.patch.object(vk, "compute", boom), mock.patch.object(vk, "evaluate_holdout", boom), \
                mock.patch("builtins.print", logs.append):
            self.assertEqual(vk.main(["--verify", "--output", out]), hr.FAILED)
        boom.assert_not_called()
        self.assertTrue(any("STALE" in str(line) for line in logs), logs)
        with open(out, "rb") as f:
            self.assertEqual(f.read(), raw)


if __name__ == "__main__":
    unittest.main(verbosity=1)
