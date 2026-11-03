"""
Fixtures for the one-shot holdout exposure record of the #72 study (#80).

Every fixture runs holdout_record.py on a synthetic study in a temporary
directory: a stand-in preregistration and coefficient file, a five-point
holdout, and a computation that returns a small record. The real holdout H is
never evaluated here. They prove that:

  1. a first evaluation writes V with EVALUATION_STARTED, evaluates the
     holdout once, then records EVALUATED with the digest of A6; a failed
     condition before it leaves the holdout unevaluated;
  2. a second first-evaluation attempt is refused before anything is
     computed, and the output stays byte-identical;
  3. interrupted runs are recovered fail-closed: a crash during the holdout
     leaves EVALUATION_STARTED, which no mode evaluates again; a stale
     temporary file, a partial (truncated) output, a failed replace and a
     concurrent change are all refused without overwriting the output;
  4. --regenerate rewrites the non-holdout parts and carries the exposure
     record and A6 over unchanged; --verify is read-only and labels its
     result a replay;
  5. a changed candidate, holdout definition or frozen preregistration text,
     and every malformed or self-contradicting record, is refused, never
     reset;
  6. validate_kernel.py refuses a first evaluation of the committed record
     before computing anything, and the committed record passes the checks;
  7. two concurrent first evaluations evaluate the holdout exactly once: the
     second run meets the lock before it reads the record, at each point of
     the first run (computing V, evaluating H), in-process and from another
     thread; a stale lock blocks every mode until it is removed by hand, and
     a run releases its lock on every exit path.

Run: python3 test_holdout_record.py   (exit 0 = pass, nonzero = fail)
"""
import copy
import hashlib
import json
import os
import shutil
import sys
import tempfile
import threading
import unittest
from unittest import mock

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import holdout_record as hr  # noqa: E402

DEFINITION = [{"seed": 11, "count": 3, "low": 0, "high": 1}, {"seed": 12, "count": 2, "low": 0, "high": 2}]


class Synthetic:
    """A study whose computation and holdout evaluation are counted stand-ins."""

    def __init__(self, root):
        self.root = root
        self.output = os.path.join(root, "results.json")
        self.prereg = os.path.join(root, "PREREGISTRATION.md")
        self.coefficients = os.path.join(root, "coefficients.json")
        with open(self.prereg, "w", encoding="utf-8") as f:
            f.write("# Synthetic preregistration\n\nH is evaluated once.\n")
        with open(self.coefficients, "w", encoding="utf-8") as f:
            f.write('{"c": [1, 2, 3]}\n')
        self.problems = []
        self.prediction = 1.5
        self.holdout_pass = True
        self.holdout_error = None
        self.computed = 0
        self.evaluated = []
        self.logs = []
        self.definition = copy.deepcopy(DEFINITION)
        self.on_compute = None

    def compute(self):
        self.computed += 1
        if self.on_compute:
            self.on_compute()
        with open(self.coefficients, "rb") as f:
            sha = hashlib.sha256(f.read()).hexdigest()
        return {"study": "synthetic", "coefficients_sha256": sha, "problems": list(self.problems),
                "versions": {"python": "x"}, "acceptance": {"A1": {"pass": True, "worst": 1e-17}},
                "predictions": {"rows": [{"predicted": self.prediction}]}, "holdout_559_inspected": False}

    def evaluate(self, points):
        self.evaluated.append(list(points))
        if self.holdout_error:
            raise self.holdout_error
        return {"pass": self.holdout_pass, "points": len(points), "evaluated_once": True,
                "worst": {"q": {"value": 5e-16, "z": points[0]}}}

    def study(self):
        return hr.Study(output=self.output, preregistration=self.prereg, coefficients=self.coefficients,
                        holdout_definition=self.definition, compute=self.compute,
                        evaluate_holdout=self.evaluate, log=self.logs.append)

    def run(self, mode, reason=None):
        return hr.run(self.study(), mode, reason)

    def bytes(self):
        with open(self.output, "rb") as f:
            return f.read()

    def record(self):
        return json.loads(self.bytes())

    def write(self, record):
        with open(self.output, "w", encoding="utf-8") as f:
            json.dump(record, f, indent=1)
            f.write("\n")

    def log(self):
        return "\n".join(self.logs)


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.s = Synthetic(self.tmp)

    def tearDown(self):
        shutil.rmtree(self.tmp)

    def evaluated_study(self):
        self.assertEqual(self.s.run("init"), hr.OK)
        self.assertEqual(self.s.run("first_evaluation"), hr.OK)
        self.s.computed, self.s.evaluated, self.s.logs[:] = 0, [], []
        return self.s.bytes()

    def assert_refused(self, mode, before, reason=None):
        self.assertEqual(self.s.run(mode, reason), hr.REFUSED, self.s.log())
        self.assertEqual(self.s.bytes(), before, f"{mode} changed the output")
        self.assertEqual(self.s.evaluated, [], f"{mode} evaluated the holdout")


class FirstEvaluation(Base):
    def test_first_evaluation_records_one_exposure(self):
        self.assertEqual(self.s.run("init"), hr.OK)
        rec = self.s.record()
        self.assertEqual(rec["status"], hr.PENDING)
        self.assertEqual(rec["holdout_exposure"]["status"], hr.NOT_EVALUATED)
        self.assertEqual(rec["holdout_exposure"]["exposures"], 0)
        self.assertNotIn("A6", rec["acceptance"])
        self.assertEqual(self.s.evaluated, [])

        self.assertEqual(self.s.run("first_evaluation"), hr.OK)
        rec = self.s.record()
        self.assertEqual(len(self.s.evaluated), 1)
        self.assertEqual(self.s.evaluated[0], hr.holdout_points(DEFINITION))
        ex = rec["holdout_exposure"]
        self.assertEqual((rec["status"], ex["status"], ex["exposures"]), (hr.PASS, hr.EVALUATED, 1))
        self.assertEqual(ex["a6_sha256"], hr.digest(rec["acceptance"]["A6"]))
        self.assertEqual(ex["holdout"]["points_sha256"], hr.points_sha256(hr.holdout_points(DEFINITION)))
        self.assertEqual(ex["design"]["sha256"], hr.file_sha256(self.s.prereg))
        self.assertEqual(hr.check_record(rec), [])
        self.assertFalse(os.path.exists(hr.temporary_path(self.s.output)))

    def test_v_is_written_before_the_holdout(self):
        self.assertEqual(self.s.run("init"), hr.OK)
        seen = []

        def evaluate(points):
            seen.append(self.s.record())
            return {"pass": True, "points": len(points), "evaluated_once": True}
        self.s.evaluate = evaluate
        self.assertEqual(self.s.run("first_evaluation"), hr.OK)
        self.assertEqual(seen[0]["holdout_exposure"]["status"], hr.STARTED)
        self.assertEqual(seen[0]["predictions"], {"rows": [{"predicted": 1.5}]})

    def test_failed_condition_leaves_holdout_unevaluated(self):
        self.assertEqual(self.s.run("init"), hr.OK)
        self.s.problems = ["A5 failed"]
        self.assertEqual(self.s.run("first_evaluation"), hr.FAILED)
        rec = self.s.record()
        self.assertEqual(self.s.evaluated, [])
        self.assertEqual((rec["status"], rec["holdout_exposure"]["status"]), (hr.FAIL, hr.NOT_EVALUATED))

    def test_failed_holdout_is_recorded_and_not_retried(self):
        self.assertEqual(self.s.run("init"), hr.OK)
        self.s.holdout_pass = False
        self.assertEqual(self.s.run("first_evaluation"), hr.FAILED)
        rec = self.s.record()
        self.assertEqual((rec["status"], rec["problems"]), (hr.FAIL, ["A6 failed"]))
        self.assertEqual(hr.check_record(rec), [])
        before, self.s.evaluated = self.s.bytes(), []
        self.s.holdout_pass = True
        self.assert_refused("first_evaluation", before)


class SecondRun(Base):
    def test_second_first_evaluation_is_refused_before_computing(self):
        before = self.evaluated_study()
        self.assert_refused("first_evaluation", before)
        self.assertEqual(self.s.computed, 0, "the refusal must come before any computation")
        self.assertIn("second first-evaluation attempt refused", self.s.log())
        self.assertIn("is unchanged", self.s.log())

    def test_init_refuses_an_existing_record(self):
        before = self.evaluated_study()
        self.assert_refused("init", before)
        self.assertEqual(self.s.computed, 0)


class InterruptedRun(Base):
    def test_crash_during_holdout_leaves_exposure_recorded(self):
        self.assertEqual(self.s.run("init"), hr.OK)
        self.s.holdout_error = KeyboardInterrupt()
        with self.assertRaises(KeyboardInterrupt):
            self.s.run("first_evaluation")
        rec = self.s.record()
        ex = rec["holdout_exposure"]
        self.assertEqual((ex["status"], ex["exposures"], rec["status"]), (hr.STARTED, 1, hr.FAIL))
        self.assertNotIn("A6", rec["acceptance"])
        self.assertTrue(rec["problems"][0].startswith("A6 not recorded"))
        self.assertEqual(hr.check_record(rec), [])
        self.assertFalse(os.path.exists(hr.temporary_path(self.s.output)))

        self.s.holdout_error, self.s.evaluated = None, []
        before = self.s.bytes()
        self.assert_refused("first_evaluation", before)
        self.assertEqual(self.s.run("verify"), hr.FAILED)
        self.assertIn("INCOMPLETE", self.s.log())
        self.assertEqual(self.s.run("regenerate", "recompute V"), hr.FAILED)
        rec2 = self.s.record()
        self.assertEqual(rec2["holdout_exposure"], ex)
        self.assertEqual(self.s.evaluated, [])

    def test_stale_temporary_file_blocks_every_mode(self):
        before = self.evaluated_study()
        tmp = hr.temporary_path(self.s.output)
        with open(tmp, "w", encoding="utf-8") as f:
            f.write('{"status": "PA')
        for mode in ("verify", "regenerate", "first_evaluation", "init"):
            self.assert_refused(mode, before, "r" if mode == "regenerate" else None)
        with open(tmp, encoding="utf-8") as f:
            self.assertEqual(f.read(), '{"status": "PA')
        self.assertIn("interrupted", self.s.log())
        os.unlink(tmp)
        self.assertEqual(self.s.run("verify"), hr.OK)

    def test_partial_output_is_refused_not_reset(self):
        full = self.evaluated_study()
        with open(self.s.output, "wb") as f:
            f.write(full[: len(full) // 2])
        before = self.s.bytes()
        for mode in ("verify", "regenerate", "first_evaluation", "init"):
            self.assert_refused(mode, before, "r" if mode == "regenerate" else None)
        self.assertIn("not valid JSON", self.s.log())
        self.assertEqual(self.s.computed, 0)

    def test_missing_output_is_never_taken_as_untouched(self):
        for mode in ("verify", "regenerate", "first_evaluation"):
            self.assertEqual(self.s.run(mode, "r" if mode == "regenerate" else None), hr.REFUSED)
            self.assertFalse(os.path.exists(self.s.output))
        self.assertEqual((self.s.computed, self.s.evaluated), (0, []))

    def test_failed_replace_keeps_the_old_output(self):
        before = self.evaluated_study()
        with mock.patch.object(hr.os, "replace", side_effect=OSError("disk full")):
            with self.assertRaises(OSError):
                self.s.run("regenerate", "recompute V")
        self.assertEqual(self.s.bytes(), before)
        self.assertFalse(os.path.exists(hr.temporary_path(self.s.output)))

    def test_concurrent_change_is_not_overwritten(self):
        self.assertEqual(self.s.run("init"), hr.OK)

        def other_run():
            rec = self.s.record()
            rec["study"] = "changed by another run"
            self.s.write(rec)
        self.s.on_compute = other_run
        self.assertEqual(self.s.run("first_evaluation"), hr.REFUSED)
        self.assertEqual(self.s.evaluated, [])
        self.assertEqual(self.s.record()["study"], "changed by another run")


class EvidencePreservation(Base):
    def test_regenerate_keeps_exposure_and_a6(self):
        before = self.evaluated_study()
        old = json.loads(before)
        self.s.prediction = 2.5
        self.assertEqual(self.s.run("regenerate", "amend the conditioning rule"), hr.OK)
        new = self.s.record()
        self.assertEqual(self.s.evaluated, [])
        self.assertEqual(new["predictions"], {"rows": [{"predicted": 2.5}]})
        self.assertEqual(hr.canonical(new["holdout_exposure"]), hr.canonical(old["holdout_exposure"]))
        self.assertEqual(hr.canonical(new["acceptance"]["A6"]), hr.canonical(old["acceptance"]["A6"]))
        self.assertEqual(new["status"], hr.PASS)
        self.assertEqual(new["regenerations"], [{"utc": new["regenerations"][0]["utc"],
                                                 "reason": "amend the conditioning rule",
                                                 "holdout_evaluated": False}])
        self.assertEqual(list(new)[: len(old)], list(old))
        self.assertEqual(self.s.run("verify"), hr.OK)
        self.assertEqual(self.s.evaluated, [])

    def test_regenerate_needs_a_reason(self):
        before = self.evaluated_study()
        self.assert_refused("regenerate", before, "  ")
        self.assertEqual(self.s.computed, 0)

    def test_regenerate_reports_a_failed_condition_and_keeps_a6(self):
        before = self.evaluated_study()
        self.s.problems = ["A1 failed"]
        self.assertEqual(self.s.run("regenerate", "r"), hr.FAILED)
        new = self.s.record()
        self.assertEqual(new["status"], hr.FAIL)
        self.assertEqual(new["acceptance"]["A6"], json.loads(before)["acceptance"]["A6"])

    def test_verify_is_read_only_and_labelled_a_replay(self):
        before = self.evaluated_study()
        self.assertEqual(self.s.run("verify"), hr.OK)
        self.assertEqual(self.s.bytes(), before)
        self.assertEqual(self.s.evaluated, [])
        log = self.s.log()
        self.assertIn("REPLAY (verification only)", log)
        self.assertIn("not a new measurement", log)
        self.assertIn("not re-established by a new holdout evaluation", log)

        self.s.prediction = 9.0
        self.assertEqual(self.s.run("verify"), hr.FAILED)
        self.assertIn("NOT REPRODUCED", self.s.log())
        self.assertIn("predictions.rows[0].predicted", self.s.log())
        self.assertEqual(self.s.bytes(), before)


class Identity(Base):
    def test_changed_candidate_is_refused(self):
        before = self.evaluated_study()
        with open(self.s.coefficients, "a", encoding="utf-8") as f:
            f.write(" ")
        for mode in ("verify", "regenerate", "first_evaluation"):
            self.assert_refused(mode, before, "r" if mode == "regenerate" else None)
        self.assertIn("candidate changed", self.s.log())

    def test_changed_holdout_definition_is_refused(self):
        before = self.evaluated_study()
        self.s.definition[0]["seed"] = 99
        for mode in ("verify", "regenerate"):
            self.assert_refused(mode, before, "r" if mode == "regenerate" else None)
        self.assertIn("holdout definition changed", self.s.log())

    def test_rewritten_preregistration_is_refused_and_appended_one_accepted(self):
        before = self.evaluated_study()
        with open(self.s.prereg, "a", encoding="utf-8") as f:
            f.write("\n## Amendment, dated\n")
        self.assertEqual(self.s.run("verify"), hr.OK)
        with open(self.s.prereg, "r+", encoding="utf-8") as f:
            f.write("# Rewritten")
        self.assert_refused("verify", before)
        self.assertIn("not amended by appending", self.s.log())


class Malformed(Base):
    def mutations(self, rec):
        def drop_exposure(r):
            del r["holdout_exposure"]

        def a6_with_not_evaluated(r):
            r["holdout_exposure"]["status"] = hr.NOT_EVALUATED
            r["holdout_exposure"]["exposures"] = 0

        def drop_a6(r):
            del r["acceptance"]["A6"]

        def edit_a6(r):
            r["acceptance"]["A6"]["worst"]["q"]["value"] = 4e-16

        def two_exposures(r):
            r["holdout_exposure"]["exposures"] = 2

        def not_once(r):
            r["acceptance"]["A6"]["evaluated_once"] = False

        def unknown_status(r):
            r["holdout_exposure"]["status"] = "RESET"

        def pass_with_failed_a6(r):
            r["acceptance"]["A6"]["pass"] = False
            r["holdout_exposure"]["a6_sha256"] = hr.digest(r["acceptance"]["A6"])

        def status_pending(r):
            r["status"] = hr.PENDING

        def candidate_mismatch(r):
            r["coefficients_sha256"] = "0" * 64

        def evaluated_regeneration(r):
            r["regenerations"] = [{"utc": "x", "reason": "y", "holdout_evaluated": True}]

        def no_started(r):
            del r["holdout_exposure"]["started_utc"]

        def not_an_object(r):
            r.clear()
            r["x"] = 1
        return [drop_exposure, a6_with_not_evaluated, drop_a6, edit_a6, two_exposures, not_once,
                unknown_status, pass_with_failed_a6, status_pending, candidate_mismatch,
                evaluated_regeneration, no_started, not_an_object]

    def test_every_inconsistent_record_is_refused(self):
        good = json.loads(self.evaluated_study())
        for mutate in self.mutations(copy.deepcopy(good)):
            with self.subTest(mutation=mutate.__name__):
                rec = copy.deepcopy(good)
                mutate(rec)
                self.assertNotEqual(hr.check_record(rec), [], mutate.__name__)
                self.s.write(rec)
                before = self.s.bytes()
                for mode in ("verify", "regenerate", "first_evaluation"):
                    self.assert_refused(mode, before, "r" if mode == "regenerate" else None)
        self.assertEqual(self.s.computed, 0)


class Locking(Base):
    def lock(self):
        return hr.lock_path(self.s.output)

    def second_actor(self):
        """Another run on the same files, with its own counted evaluator."""
        other = Synthetic.__new__(Synthetic)
        other.__dict__.update(self.s.__dict__)
        other.evaluated, other.logs, other.computed, other.on_compute = [], [], 0, None
        return other

    @unittest.skipIf(os.name != "posix", "POSIX file modes")
    def test_lock_and_output_are_owner_only(self):
        # Neither the lock nor the atomically written output is readable by
        # other users (CodeQL: overly permissive file permissions).
        self.assertEqual(self.s.run("init"), hr.OK)
        modes = []

        def observe():
            self.s.on_compute = None
            modes.append(os.stat(self.lock()).st_mode & 0o777)
        self.s.on_compute = observe
        self.assertEqual(self.s.run("first_evaluation"), hr.OK)
        self.assertEqual(len(modes), 1)
        self.assertEqual(modes[0] & 0o077, 0, oct(modes[0]))
        output_mode = os.stat(self.s.output).st_mode & 0o777
        self.assertEqual(output_mode & 0o077, 0, oct(output_mode))

    def test_second_run_during_v_is_refused_and_h_evaluated_once(self):
        self.assertEqual(self.s.run("init"), hr.OK)
        other, outcome = self.second_actor(), []

        def race():
            # The first run has read the NOT_EVALUATED record and is computing V.
            self.s.on_compute = None
            outcome.append(other.run("first_evaluation"))
        self.s.on_compute = race
        self.assertEqual(self.s.run("first_evaluation"), hr.OK)
        self.assertEqual(outcome, [hr.REFUSED])
        self.assertEqual((other.computed, other.evaluated), (0, []))
        self.assertIn("holds the lock", other.log())
        self.assertEqual(len(self.s.evaluated), 1)
        ex = self.s.record()["holdout_exposure"]
        self.assertEqual((ex["status"], ex["exposures"]), (hr.EVALUATED, 1))
        self.assertFalse(os.path.lexists(self.lock()))

    def test_second_run_during_h_is_refused(self):
        self.assertEqual(self.s.run("init"), hr.OK)
        other, outcome = self.second_actor(), []
        evaluate = self.s.evaluate

        def evaluate_with_race(points):
            # EVALUATION_STARTED is written and the first run is evaluating H.
            for mode in ("first_evaluation", "regenerate", "init", "verify"):
                outcome.append(other.run(mode, "r" if mode == "regenerate" else None))
            return evaluate(points)
        self.s.evaluate = evaluate_with_race
        self.assertEqual(self.s.run("first_evaluation"), hr.OK)
        self.assertEqual(outcome, [hr.REFUSED] * 4)
        self.assertEqual((other.computed, other.evaluated), (0, []))
        self.assertEqual(len(self.s.evaluated), 1)

    def test_concurrent_threads_evaluate_h_once(self):
        self.assertEqual(self.s.run("init"), hr.OK)
        other = self.second_actor()
        inside, release, results = threading.Event(), threading.Event(), {}
        evaluate = self.s.evaluate

        def blocking_evaluate(points):
            inside.set()
            self.assertTrue(release.wait(10))
            return evaluate(points)
        self.s.evaluate = blocking_evaluate
        first = threading.Thread(target=lambda: results.setdefault("first", self.s.run("first_evaluation")))
        first.start()
        try:
            self.assertTrue(inside.wait(10))
            results["second"] = other.run("first_evaluation")
        finally:
            release.set()
            first.join(10)
        self.assertEqual(results, {"first": hr.OK, "second": hr.REFUSED})
        self.assertEqual(other.evaluated, [])
        self.assertEqual(len(self.s.evaluated), 1)
        self.assertEqual(self.s.record()["holdout_exposure"]["exposures"], 1)

    def test_stale_lock_blocks_every_mode_until_removed_by_hand(self):
        self.assertEqual(self.s.run("init"), hr.OK)
        before = self.s.bytes()
        with open(self.lock(), "w", encoding="utf-8") as f:
            f.write("pid 1, mode first_evaluation, killed\n")
        self.s.computed = 0
        for mode in ("first_evaluation", "regenerate", "init", "verify"):
            self.assert_refused(mode, before, "r" if mode == "regenerate" else None)
        self.assertEqual(self.s.computed, 0)
        with open(self.lock(), encoding="utf-8") as f:
            self.assertEqual(f.read(), "pid 1, mode first_evaluation, killed\n")
        self.assertIn("delete the lock file by hand", self.s.log())
        self.assertIn("killed", self.s.log())
        os.unlink(self.lock())
        self.assertEqual(self.s.run("first_evaluation"), hr.OK)
        self.assertEqual(len(self.s.evaluated), 1)

    def test_removing_a_stale_lock_does_not_allow_a_reevaluation(self):
        self.assertEqual(self.s.run("init"), hr.OK)
        self.s.holdout_error = SystemExit(9)          # a run that dies inside H
        with self.assertRaises(SystemExit):
            self.s.run("first_evaluation")
        with open(self.lock(), "w", encoding="utf-8") as f:
            f.write("left by a killed run\n")
        os.unlink(self.lock())                         # the documented manual step
        self.s.holdout_error, self.s.evaluated = None, []
        self.assert_refused("first_evaluation", self.s.bytes())
        self.assertEqual(self.s.record()["holdout_exposure"]["status"], hr.STARTED)

    def test_lock_is_released_on_every_exit_path(self):
        self.assertEqual(self.s.run("init"), hr.OK)                      # success
        self.assertFalse(os.path.lexists(self.lock()))
        self.assertEqual(self.s.run("init"), hr.REFUSED)                 # refusal inside the lock
        self.assertFalse(os.path.lexists(self.lock()))
        self.s.on_compute = mock.Mock(side_effect=RuntimeError("compute"))
        with self.assertRaises(RuntimeError):                            # crash in V
            self.s.run("regenerate", "r")
        self.assertFalse(os.path.lexists(self.lock()))
        self.s.on_compute = None
        self.s.problems = ["A5 failed"]
        self.assertEqual(self.s.run("first_evaluation"), hr.FAILED)      # failed condition
        self.assertFalse(os.path.lexists(self.lock()))
        self.s.problems = []
        self.s.holdout_error = KeyboardInterrupt()
        with self.assertRaises(KeyboardInterrupt):                       # crash in H
            self.s.run("first_evaluation")
        self.assertFalse(os.path.lexists(self.lock()))

    def test_verify_takes_no_lock(self):
        self.evaluated_study()
        seen = []
        self.s.on_compute = lambda: seen.append(os.path.lexists(self.lock()))
        self.assertEqual(self.s.run("verify"), hr.OK)
        self.assertEqual(seen, [False])
        self.assertEqual(sorted(os.listdir(self.tmp)), ["PREREGISTRATION.md", "coefficients.json", "results.json"])

    def test_a_replaced_lock_is_left_in_place(self):
        self.assertEqual(self.s.run("init"), hr.OK)

        def replace_lock():
            with open(self.lock(), "w", encoding="utf-8") as f:
                f.write("someone else\n")
        self.s.on_compute = replace_lock
        self.s.run("regenerate", "r")
        self.assertTrue(os.path.lexists(self.lock()))
        self.assertIn("left in place", self.s.log())


class CommittedRecord(unittest.TestCase):
    """validate_kernel.py on the committed results.json, without computing anything."""

    @classmethod
    def setUpClass(cls):
        import validate_kernel as vk
        cls.vk = vk
        with open(vk.OUT, "rb") as f:
            cls.raw = f.read()

    def test_committed_record_is_consistent(self):
        rec = json.loads(self.raw)
        self.assertEqual(hr.check_record(rec), [])
        self.assertEqual(hr.check_identity(rec["holdout_exposure"], self.vk.make_study()), [])
        self.assertEqual(rec["holdout_exposure"]["status"], hr.EVALUATED)

    def test_validator_refuses_a_second_first_evaluation(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = os.path.join(tmp, "results.json")
            with open(out, "wb") as f:
                f.write(self.raw)
            boom = mock.Mock(side_effect=AssertionError("must not compute"))
            with mock.patch.object(self.vk, "compute", boom), \
                    mock.patch.object(self.vk, "evaluate_holdout", boom), \
                    mock.patch.object(self.vk, "measure_point", boom), \
                    mock.patch("builtins.print"):
                self.assertEqual(self.vk.main(["--first-evaluation", "--output", out]), hr.REFUSED)
                self.assertEqual(self.vk.main(["--init", "--output", out]), hr.REFUSED)
                self.assertEqual(self.vk.main(["--regenerate", "--output", out]), hr.REFUSED)
            boom.assert_not_called()
            with open(out, "rb") as f:
                self.assertEqual(f.read(), self.raw)

    def test_validator_needs_an_explicit_mode(self):
        with mock.patch("sys.stderr"), self.assertRaises(SystemExit) as e:
            self.vk.main([])
        self.assertEqual(e.exception.code, 2)


if __name__ == "__main__":
    unittest.main(verbosity=1)
