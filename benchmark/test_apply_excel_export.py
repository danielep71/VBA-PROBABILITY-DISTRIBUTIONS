"""Fixture proof for apply_excel_export.py (#181).

Every case builds a private Git repository with the evidence policy, tiny
grids in the real schemas, and copies of the real apply, finalize, check and
excel_certification tools. refresh_evidence.py is replaced by a stub that
records its arguments, because the real one needs the whole evidence tree; the
finalizer and the certification checker run as committed. A fixture artifact
has the shape ci/Export-ExcelGrids.ps1 uploads.

Covered: the happy path (with a post-commit validation of the evidence commit),
a SHA mismatch, a red regression, a missing grid, a digest mismatch, a changed
out-of-scope row, an environment file of another commit, the accepted
pre-commit record failure of refresh_evidence.py and an open item, a
holdout_older-only export, the finalizer's own commit_sha check, and the grid
check itself.

Run: python3 test_apply_excel_export.py   (exit 0 = pass)
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import apply_excel_export as A  # noqa: E402
import excel_certification as E  # noqa: E402
import _manifest as M  # noqa: E402

TOOLS = ("apply_excel_export.py", "excel_certification.py", "_manifest.py",
         "finalize_excel_certification.py", "check_excel_certification.py")
MAIN, HOLDOUT, OLDER, BOUNDARY, TCHI = (A.GRIDS[n].path for n in (
    "main", "holdout", "holdout_older", "survival_boundary", "tchi_large_df"))
HEADER13 = ("function,vba_kernel,claim,metric,arg1,arg2,arg3,arg4,reference,"
            "observed_vba,regime,evidence_set,expected_error")
GRID_TEXT = {
    MAIN: HEADER13 + "\n"
    "Normal_Cumulative,K,rel<=1E-15,rel,0.5,,,,0.69,6.9E-001;0E+000,all,main grid,\n"
    "Normal_Survival,K,rel<=1E-15,rel,0.5,,,,0.31,,all,main grid,\n"
    "Gamma_Cumulative,K,rel<=1E-15,rel,2,3,,,0.1,1.0E-001;0E+000,small,gamma study,\n",
    HOLDOUT: HEADER13 + "\n"
    "Beta_Density,K,,,0.2,2,3,,1.5,1.4E+000;0E+000,balanced,holdout,\n"
    "F_Survival,K,,,1.5,4,7,,0.3,3.0E-001;0E+000,all,holdout,\n",
    OLDER: "function,vba_kernel,claim,metric,arg1,arg2,arg3,reference,observed_vba,regime,"
           "evidence_set\n"
           "Normal_Cumulative,K,holdout,rel,1.35,0.4,1.7,0.71,7.1E-001;0E+000,all,holdout_older\n",
    BOUNDARY: "function,vba_kernel,claim,metric,arg1,arg2,arg3,reference,observed_vba,regime,"
              "evidence_set\n"
              "NormalStandard_Survival,K,boundary,rel,3.5,,,0.00023,2.3E-004;0E+000,all,"
              "survival_boundary\n",
    TCHI: "function,vba_kernel,claim,metric,arg1,arg2,arg3,reference,observed_vba,regime,"
          "evidence_set\n"
          "StudentT_Survival,K,characterization,rel,2.0,100000000.0,,0.0227,2.27E-002;0E+000,"
          "meaningful,tchi_large_df\n",
}
STUB = '''import json, os, sys
with open(os.environ["REFRESH_STUB_LOG"], "w") as f:
    json.dump(sys.argv[1:], f)
for label in filter(None, os.environ.get("REFRESH_STUB_FAIL", "").split("|")):
    print("  FAIL  " + label)
sys.exit(1 if os.environ.get("REFRESH_STUB_FAIL") else 0)
'''


def git(root, *args):
    proc = subprocess.run(["git", "-C", root] + list(args), capture_output=True)
    if proc.returncode:
        raise AssertionError(proc.stderr.decode())
    return proc.stdout.decode()


def write(root, rel, data):
    path = os.path.join(root, *rel.split("/"))
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as f:
        f.write(data if isinstance(data, bytes) else data.encode())


def read(path):
    with open(path, "rb") as f:
        return f.read()


def fill(spec, mode, baseline, value=b"2.5E+000;0E+000", error_first=False):
    """What the exporter writes: every in-scope observed_vba cell filled, CRLF."""
    lines = A.grid_lines(baseline)
    first = True
    for i in range(1, len(lines)):
        if A.row_in_scope(spec, mode, lines[i]):
            cols = lines[i].split(b",")
            cols[spec.observed] = b"ERROR" if (error_first and first) else value
            first = False
            lines[i] = b",".join(cols)
    return A.printed(lines)


class Fixture:
    def __init__(self):
        self.root = tempfile.mkdtemp(prefix="apply-export-repo-")
        self.scratch = tempfile.mkdtemp(prefix="apply-export-artifact-")
        root = self.root
        git(root, "init", "-q")
        git(root, "config", "user.email", "t@t")
        git(root, "config", "user.name", "t")
        git(root, "config", "commit.gpgsign", "false")
        self.policy = {
            "schema_version": 1, "repository": "owner/repo",
            "workflow": ".github/workflows/excel.yml", "entry_point": "RunAll_CI",
            "expected_assertions": 3, "regression_sources": ["src/A.bas", "tests/T.bas"],
            "grid_exporters": {MAIN: A.GRIDS["main"].exporter,
                               HOLDOUT: A.GRIDS["holdout"].exporter},
        }
        write(root, ".gitattributes", "*.csv text eol=lf\n*.json text eol=lf\n")
        write(root, E.POLICY_PATH, json.dumps(self.policy, indent=2) + "\n")
        for rel in ["src/A.bas", "tests/T.bas"] + [s.exporter for s in A.GRIDS.values()]:
            write(root, rel, 'Attribute VB_Name = "%s"\n' % os.path.basename(rel)[:-4])
        for rel, text in GRID_TEXT.items():
            write(root, rel, text)
        write(root, A.ENV_PATH, '{\n  "excel_version": "16.0",\n  "excel_build": "1",\n'
                                '  "office_bitness": "64"\n}\n')
        write(root, A.RETAINED_RECORD, "{}\n")
        for tool in TOOLS:
            shutil.copy(os.path.join(HERE, tool), os.path.join(root, "benchmark", tool))
        write(root, "benchmark/refresh_evidence.py", STUB)
        self.bind_main()
        self.head = self.commit("base")

    def bind_main(self):
        """The part of observation_manifest.json a missing-only export reads."""
        write(self.root, "benchmark/" + M.MANIFEST_NAME, json.dumps(
            {"source_binding": M.source_hashes(self.root)}, indent=2) + "\n")

    def commit(self, message):
        git(self.root, "add", "-A")
        git(self.root, "commit", "-qm", message)
        return git(self.root, "rev-parse", "HEAD").strip()

    def close(self):
        shutil.rmtree(self.root, ignore_errors=True)
        shutil.rmtree(self.scratch, ignore_errors=True)

    def artifact(self, names=("main", "holdout"), mode="full", name="artifact"):
        """An artifact in the shape ci/Export-ExcelGrids.ps1 uploads."""
        root, head, out = self.root, self.head, os.path.join(self.scratch, name)
        os.makedirs(out)
        log = b"[2026-11-06 10:00:00] RESULT: ALL TESTS PASSED\r\n"
        write(out, A.LOG, log)
        record = {
            "schema_version": 1, "repository": "owner/repo", "candidate_sha": head,
            "execution": "automated", "started_at": "2026-11-06T10:00:00Z",
            "finished_at": "2026-11-06T10:01:00Z",
            "runner": {"class": "self-hosted-excel", "identity": "PROBDIST",
                       "workflow": {"repository": "owner/repo",
                                    "path": ".github/workflows/excel.yml", "sha": head,
                                    "run_id": 42, "run_attempt": 1}},
            "environment": {"excel_version": "16.0", "excel_build": "20228",
                            "office_bitness": "64-bit", "windows": "Windows 11",
                            "os_architecture": "X64", "automation_security": "1",
                            "vba_project_access": "preconfigured"},
            "sources": [{"path": p, "sha256": E.sha256_bytes(E.git_bytes(root, head, p))}
                        for p in sorted(self.policy["regression_sources"])],
            "stages": {s: {"status": "PASS", "detail": "ok"} for s in E.STAGES},
            "harness": {"entry_point": "RunAll_CI", "assertions": 3, "passed": 3, "failed": 0},
            "log": {"path": A.LOG, "sha256": E.sha256_bytes(log)},
            "grids": {g: {"exported": False, "sha256": None, "row_count": None,
                          "exported_utc": None} for g in self.policy["grid_exporters"]},
        }
        excel_env = b'{\r\n  "excel_version": "16.0",\r\n  "excel_build": "20228",\r\n' \
                    b'  "office_bitness": "64"\r\n}\r\n'
        env = ('{\n  "excel_version": "16.0",\n  "excel_build": "20228",\n'
               '  "office_bitness": "64",\n  "commit_sha": "%s",\n'
               '  "commit_sha_recorded_by": "%s"\n}\n' % (head, A.ADAPTER)).encode()
        write(out, A.EXCEL_WRITTEN_ENVIRONMENT, excel_env)
        write(out, A.ENVIRONMENT, env)
        grids, sources = [], ["src/A.bas"]
        for n in names:
            spec = A.GRIDS[n]
            grid_mode = mode if n == "main" else "full"
            baseline = E.git_bytes(root, head, spec.path)
            raw = fill(spec, grid_mode, baseline, error_first=(n == "main"))
            write(out, "grids/" + spec.path, raw)
            counts = A.verify_grid(spec, grid_mode, baseline, raw)
            assert not counts.pop("problems"), "fixture export must verify"
            grids.append({"name": n, "path": spec.path, "artifact_path": "grids/" + spec.path,
                          "mode": grid_mode, "macro": spec.macros[grid_mode],
                          "exporter": spec.exporter, **counts,
                          "baseline_sha256": E.sha256_bytes(baseline),
                          "raw_sha256": E.sha256_bytes(raw), "seconds": 1.5,
                          "verified": True, "problems": []})
            sources.append(spec.exporter)
        sources = sorted(set(sources) | {A.GRIDS["main"].exporter})
        self.save_record(out, record)
        manifest = {
            "schema_version": 1, "kind": "excel-grid-export", "adapter": A.ADAPTER,
            "repository": "owner/repo", "candidate_sha": head, "status": "PASS",
            "failure": None, "requested": list(names), "main_mode": mode,
            "workflow": {"repository": "owner/repo", "path": ".github/workflows/excel.yml",
                         "ref": "refs/heads/feat", "run_id": 42, "run_attempt": 1,
                         "runner": "PROBDIST"},
            "regression": {"record": A.RECORD, "record_sha256": None, "log": A.LOG,
                           "assertions": 3, "passed": 3, "failed": 0},
            "excel": {"version": "16.0", "build": "20228", "office_bitness": "64-bit",
                      "process_id": 1234},
            "environment": {"path": A.ENVIRONMENT, "sha256": E.sha256_bytes(env),
                            "excel_written": A.EXCEL_WRITTEN_ENVIRONMENT,
                            "excel_written_sha256": E.sha256_bytes(excel_env),
                            "commit_sha": head},
            "sources": [{"path": p, "sha256": E.sha256_bytes(E.git_bytes(root, head, p))}
                        for p in sources],
            "grids": grids, "macros": [], "cleanup": "Closed",
            "started_at": "2026-11-06T10:01:00Z", "finished_at": "2026-11-06T10:02:00Z",
            "total_seconds": 60.0,
        }
        self.save_manifest(out, manifest)
        return out

    @staticmethod
    def load(out, name):
        with open(os.path.join(out, name), encoding="utf-8-sig") as f:
            return json.load(f)

    def save_record(self, out, record):
        write(out, A.RECORD, json.dumps(record, indent=2) + "\n")
        manifest_path = os.path.join(out, A.MANIFEST)
        if os.path.exists(manifest_path):
            manifest = self.load(out, A.MANIFEST)
            manifest["regression"]["record_sha256"] = E.sha256_bytes(
                read(os.path.join(out, A.RECORD)))
            write(out, A.MANIFEST, json.dumps(manifest, indent=2) + "\n")

    def save_manifest(self, out, manifest):
        manifest["regression"]["record_sha256"] = E.sha256_bytes(
            read(os.path.join(out, A.RECORD)))
        write(out, A.MANIFEST, json.dumps(manifest, indent=2) + "\n")

    def apply(self, out, *extra, fail=""):
        stub_log = os.path.join(self.scratch, "refresh-args.json")
        if os.path.exists(stub_log):
            os.remove(stub_log)
        env = dict(os.environ, REFRESH_STUB_LOG=stub_log, REFRESH_STUB_FAIL=fail)
        proc = subprocess.run([sys.executable,
                               os.path.join(self.root, "benchmark", "apply_excel_export.py"),
                               "--artifact", out, *extra],
                              cwd=self.root, capture_output=True, text=True, env=env)
        refresh = None
        if os.path.exists(stub_log):
            with open(stub_log) as f:
                refresh = json.load(f)
        return proc.returncode, proc.stdout + proc.stderr, refresh

    def dirty(self):
        return git(self.root, "status", "--porcelain", "--untracked-files=no").strip()


class ApplyExcelExportTests(unittest.TestCase):
    def setUp(self):
        self.fx = Fixture()
        self.addCleanup(self.fx.close)

    def refused(self, out, needle, *extra):
        code, output, refresh = self.fx.apply(out, *extra)
        self.assertEqual(code, 1, output)
        self.assertIn("REFUSED", output)
        self.assertIn(needle, output)
        self.assertEqual(self.fx.dirty(), "", "a refusal must write nothing")
        self.assertIsNone(refresh, "a refusal must not run refresh_evidence.py")

    def test_happy_path_main_and_holdout(self):
        fx = self.fx
        out = fx.artifact()
        code, output, refresh = fx.apply(out, "--grid", "main", "--grid", "holdout")
        self.assertEqual(code, 0, output)
        self.assertEqual(refresh, ["--bind-exported-main", "--bind-exported-holdout"])
        for spec in (A.GRIDS["main"], A.GRIDS["holdout"]):
            raw = read(os.path.join(out, "grids", spec.path))
            self.assertIn(b"\r\n", raw)
            self.assertEqual(read(os.path.join(fx.root, spec.path)), E.canonical_grid_bytes(raw),
                             "the working tree holds the LF bytes Git stores")
        record = fx.load(fx.root, A.RETAINED_RECORD)
        self.assertEqual(record["grids"][MAIN]["raw_sha256"],
                         E.sha256_bytes(read(os.path.join(out, "grids", MAIN))),
                         "the finalizer saw the bytes Excel wrote")
        self.assertIn("Evidence ready to commit.", output)
        env = fx.load(fx.root, A.ENV_PATH)
        self.assertEqual(env["commit_sha"], fx.head)
        self.assertTrue(record["grids"][MAIN]["exported"])
        self.assertTrue(record["grids"][HOLDOUT]["exported"])
        self.assertIn("evidence: Excel export of main, holdout at " + fx.head[:7], output)
        self.assertIn("1 ERROR", output)
        self.assertIn("--require-grid " + MAIN, output)

        evidence = fx.commit("evidence")
        self.assertEqual(E.git_bytes(fx.root, evidence, MAIN),
                         E.canonical_grid_bytes(read(os.path.join(out, "grids", MAIN))),
                         "Git stores the export as LF")
        for grid in (MAIN, HOLDOUT):
            proc = subprocess.run(
                [sys.executable, os.path.join(fx.root, "benchmark",
                                                    "check_excel_certification.py"),
                 "--record", os.path.join(fx.root, A.RETAINED_RECORD),
                 "--candidate-sha", fx.head, "--evidence-ref", evidence, "--require-pass",
                 "--require-grid", grid, "--log-directory", out],
                cwd=fx.root, capture_output=True, text=True)
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_sha_mismatch_is_refused(self):
        out = self.fx.artifact()
        write(self.fx.root, "notes.md", "later\n")
        self.fx.commit("a later commit")
        self.refused(out, "but HEAD is")

    def test_red_regression_is_refused(self):
        out = self.fx.artifact()
        record = self.fx.load(out, A.RECORD)
        record["stages"]["regression"] = {"status": "FAIL", "detail": "1 of 3 failed"}
        record["harness"].update(passed=2, failed=1)
        self.fx.save_record(out, record)
        self.refused(out, "not green")

    def test_missing_grid_is_refused(self):
        out = self.fx.artifact()
        os.remove(os.path.join(out, "grids", HOLDOUT))
        self.refused(out, "missing from the artifact")

    def test_requested_grid_absent_from_artifact_is_refused(self):
        out = self.fx.artifact(names=("main",))
        self.refused(out, "not main, holdout", "--grid", "main", "--grid", "holdout")

    def test_digest_mismatch_is_refused(self):
        out = self.fx.artifact()
        path = os.path.join(out, "grids", MAIN)
        raw = read(path)
        write(out, "grids/" + MAIN, raw.replace(b"2.5E+000", b"2.6E+000", 1))
        self.refused(out, "differ from the raw SHA-256")

    def test_changed_out_of_scope_row_is_refused(self):
        out = self.fx.artifact()
        raw = read(os.path.join(out, "grids", MAIN))
        raw = raw.replace(b"1.0E-001;0E+000,small,gamma study", b"1.1E-001;0E+000,small,gamma study")
        write(out, "grids/" + MAIN, raw)
        manifest = self.fx.load(out, A.MANIFEST)
        manifest["grids"][0]["raw_sha256"] = E.sha256_bytes(raw)
        self.fx.save_manifest(out, manifest)
        self.refused(out, "out-of-scope row changed")

    def test_environment_of_another_commit_is_refused(self):
        out = self.fx.artifact()
        path = os.path.join(out, A.ENVIRONMENT)
        raw = read(path).replace(self.fx.head.encode(), b"f" * 40)
        write(out, A.ENVIRONMENT, raw)
        manifest = self.fx.load(out, A.MANIFEST)
        manifest["environment"]["sha256"] = E.sha256_bytes(raw)
        self.fx.save_manifest(out, manifest)
        self.refused(out, "records commit " + "f" * 40)

    def test_failed_export_manifest_is_refused(self):
        out = self.fx.artifact()
        manifest = self.fx.load(out, A.MANIFEST)
        manifest.update(status="FAIL", failure="watchdog: interactive dialog")
        self.fx.save_manifest(out, manifest)
        self.refused(out, "watchdog: interactive dialog")

    def test_pre_commit_record_failure_is_accepted(self):
        out = self.fx.artifact()
        code, output, _ = self.fx.apply(out, fail=A.PRE_COMMIT_STEP)
        self.assertEqual(code, 0, output)
        self.assertIn("fails until the evidence is committed", output)

    def test_other_refresh_failure_is_not_ready(self):
        out = self.fx.artifact()
        code, output, _ = self.fx.apply(out, fail=A.PRE_COMMIT_STEP + "|strict accuracy gate")
        self.assertEqual(code, 1, output)
        self.assertIn("NOT READY", output)
        self.assertIn("  - strict accuracy gate", output)
        self.assertNotIn("  - " + A.PRE_COMMIT_STEP, output)
        self.assertNotIn("Evidence ready to commit.", output)
        self.assertIn(MAIN, self.fx.dirty(), "the applied files stay for inspection")

    def test_holdout_older_only_binds_nothing(self):
        fx = self.fx
        out = fx.artifact(names=("holdout_older",))
        code, output, refresh = fx.apply(out)
        self.assertEqual(code, 0, output)
        self.assertEqual(refresh, [])
        self.assertNotIn("finalize_excel_certification.py", output.split("Commit message:")[0])
        changed = fx.dirty()
        self.assertIn(OLDER, changed)
        self.assertNotIn(A.RETAINED_RECORD, changed)

    def test_survival_boundary_only_binds_nothing(self):
        fx = self.fx
        out = fx.artifact(names=("survival_boundary",))
        code, output, refresh = fx.apply(out)
        self.assertEqual(code, 0, output)
        self.assertEqual(refresh, [])
        self.assertNotIn("finalize_excel_certification.py", output.split("Commit message:")[0])
        changed = fx.dirty()
        self.assertIn(BOUNDARY, changed)
        self.assertNotIn(A.RETAINED_RECORD, changed)
        self.assertIn("full export by Export_SurvivalBoundary", output)

    def test_tchi_large_df_only_binds_nothing(self):
        fx = self.fx
        out = fx.artifact(names=("tchi_large_df",))
        code, output, refresh = fx.apply(out)
        self.assertEqual(code, 0, output)
        self.assertEqual(refresh, [])
        changed = fx.dirty()
        self.assertIn(TCHI, changed)
        self.assertNotIn(A.RETAINED_RECORD, changed)
        self.assertIn("full export by Export_TChiLargeDf", output)

    def test_missing_only_main_export(self):
        out = self.fx.artifact(names=("main",), mode="missing-only")
        manifest = self.fx.load(out, A.MANIFEST)
        self.assertEqual(manifest["grids"][0]["in_scope_rows"], 1)
        code, output, refresh = self.fx.apply(out)
        self.assertEqual(code, 0, output)
        self.assertEqual(refresh, ["--bind-exported-main"])
        self.assertIn("missing-only export by Export_Accuracy_MissingOnly", output)

    def test_missing_only_after_a_vba_change_is_refused(self):
        fx = self.fx
        write(fx.root, "src/A.bas", 'Attribute VB_Name = "A"\n\' a numerical change\n')
        fx.head = fx.commit("change a kernel")
        out = fx.artifact(names=("main",), mode="missing-only")
        self.refused(out, "differ from the source they were bound to: src/A.bas")
        code, output, _ = fx.apply(fx.artifact(names=("main",), name="full"))
        self.assertEqual(code, 0, "a full export after the change is accepted:\n" + output)

    def test_missing_only_without_a_main_binding_is_refused(self):
        fx = self.fx
        os.remove(os.path.join(fx.root, "benchmark", M.MANIFEST_NAME))
        fx.head = fx.commit("drop the binding")
        self.refused(fx.artifact(names=("main",), mode="missing-only"),
                     "needs the committed observation_manifest.json")

    def test_finalizer_refuses_environment_of_another_commit(self):
        fx = self.fx
        out = fx.artifact()
        write(fx.root, A.ENV_PATH, ('{"excel_version": "16.0", "excel_build": "20228", '
                                    '"office_bitness": "64", "commit_sha": "%s"}\n') % ("e" * 40))
        proc = subprocess.run(
            [sys.executable, os.path.join(fx.root, "benchmark",
                                                "finalize_excel_certification.py"),
             "--record", os.path.join(out, A.RECORD), "--main", "--from-fresh-export"],
            cwd=fx.root, capture_output=True, text=True)
        self.assertEqual(proc.returncode, 1, proc.stdout)
        self.assertIn("commit_sha", proc.stdout)


class GridCheckTests(unittest.TestCase):
    spec = A.GRIDS["main"]
    baseline = GRID_TEXT[MAIN].encode()

    def test_filled_crlf_export_verifies(self):
        result = A.verify_grid(self.spec, "full", self.baseline, fill(self.spec, "full", self.baseline))
        self.assertEqual(result["problems"], [])
        self.assertEqual((result["rows"], result["in_scope_rows"], result["out_of_scope_rows"]),
                         (3, 2, 1))

    def test_lf_export_is_not_what_excel_writes(self):
        exported = fill(self.spec, "full", self.baseline).replace(b"\r\n", b"\n")
        problems = A.verify_grid(self.spec, "full", self.baseline, exported)["problems"]
        self.assertIn("the export is not CRLF-terminated lines as VBA Print # writes them", problems)

    def test_blank_in_scope_cell_is_reported(self):
        exported = A.printed(A.grid_lines(self.baseline))
        problems = A.verify_grid(self.spec, "full", self.baseline, exported)["problems"]
        self.assertIn("line 3: observed_vba left blank", problems)

    def test_missing_only_scope_and_untouched_rows(self):
        lines = A.grid_lines(self.baseline)
        self.assertEqual([A.row_in_scope(self.spec, "missing-only", line) for line in lines[1:]],
                         [False, True, False])
        exported = fill(self.spec, "full", self.baseline)
        problems = A.verify_grid(self.spec, "missing-only", self.baseline, exported)["problems"]
        self.assertIn("line 2: out-of-scope row changed", problems)

    def test_dropped_row_is_reported(self):
        lines = A.grid_lines(fill(self.spec, "full", self.baseline))
        problems = A.verify_grid(self.spec, "full", self.baseline, A.printed(lines[:-1]))["problems"]
        self.assertIn("line count changed from 4 to 3", problems)


if __name__ == "__main__":
    unittest.main()
