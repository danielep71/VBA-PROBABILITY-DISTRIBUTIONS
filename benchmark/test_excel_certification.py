"""Negative/positive fixtures for the exact-SHA Excel certification contract."""
import copy
import json
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import excel_certification as E

FAILS = []


def check(condition, message):
    if not condition:
        FAILS.append(message)


def run(root, *args):
    return subprocess.run(list(args), cwd=root, capture_output=True, check=False)


def write(root, rel, text):
    path = os.path.join(root, rel)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)


def commit(root, message):
    run(root, "git", "add", "-A")
    run(root, "git", "commit", "-qm", message)
    return run(root, "git", "rev-parse", "HEAD").stdout.decode().strip()


def base_repo():
    root = tempfile.mkdtemp()
    run(root, "git", "init", "-q")
    run(root, "git", "config", "user.email", "t@t")
    run(root, "git", "config", "user.name", "t")
    policy = {
        "schema_version": 1,
        "repository": "owner/repo",
        "workflow": ".github/workflows/excel.yml",
        "entry_point": "RunAll_CI",
        "expected_assertions": 3,
        "regression_sources": ["src/A.bas", "tests/T.bas"],
        "grid_exporters": {
            "benchmark/main.csv": "benchmark/MainExport.bas",
            "benchmark/holdout/holdout.csv": "benchmark/holdout/HoldoutExport.bas",
        },
    }
    write(root, E.POLICY_PATH, json.dumps(policy, indent=2) + "\n")
    write(root, "src/A.bas", "Attribute VB_Name = \"A\"\n")
    write(root, "tests/T.bas", "Attribute VB_Name = \"T\"\n")
    write(root, "benchmark/MainExport.bas", "Attribute VB_Name = \"MainExport\"\n")
    write(root, "benchmark/holdout/HoldoutExport.bas", "Attribute VB_Name = \"HoldoutExport\"\n")
    write(root, "benchmark/main.csv", "function,value\nA,1\nB,2\n")
    write(root, "benchmark/holdout/holdout.csv", "function,value\nC,3\n")
    sha = commit(root, "base")
    return root, sha, policy


def digest_at(root, sha, path):
    return E.sha256_bytes(E.git_bytes(root, sha, path))


def good_record(root, sha, policy, export_main=False):
    source_paths = list(policy["regression_sources"])
    grids = {
        grid: {"exported": False, "sha256": None, "row_count": None, "exported_utc": None}
        for grid in policy["grid_exporters"]
    }
    if export_main:
        grid = "benchmark/main.csv"
        source_paths.append(policy["grid_exporters"][grid])
        raw = E.git_bytes(root, sha, grid)
        grids[grid] = {"exported": True, "sha256": E.sha256_bytes(raw),
                       "row_count": E.row_count_bytes(raw),
                       "exported_utc": "2026-09-16T04:00:01Z"}
    sources = [{"path": p, "sha256": digest_at(root, sha, p)} for p in sorted(source_paths)]
    return {
        "schema_version": 1,
        "repository": policy["repository"],
        "candidate_sha": sha,
        "execution": "automated",
        "started_at": "2026-09-16T04:00:00Z",
        "finished_at": "2026-09-16T04:00:02Z",
        "runner": {"class": "self-hosted-excel", "identity": "PROBDIST",
                   "workflow": {"repository": policy["repository"],
                                "path": policy["workflow"], "sha": sha,
                                "run_id": 1, "run_attempt": 1}},
        "environment": {"excel_version": "16.0", "excel_build": "20228",
                        "office_bitness": "64-bit", "windows": "Windows 11 10.0.26200",
                        "os_architecture": "X64", "automation_security": "1 (isolated process)",
                        "vba_project_access": "preconfigured runner"},
        "sources": sources,
        "stages": {name: {"status": "PASS", "detail": "ok"} for name in E.STAGES},
        "harness": {"entry_point": policy["entry_point"], "assertions": 3,
                    "passed": 3, "failed": 0},
        "log": {"path": "test-result.txt", "sha256": "sha256:" + "0" * 64},
        "grids": grids,
    }


def expect_fail(root, record, message, **kwargs):
    try:
        E.validate_record(record, root, **kwargs)
    except E.CertificationError:
        return
    FAILS.append(message)


root, sha, policy = base_repo()
try:
    record = good_record(root, sha, policy)
    E.validate_record(record, root, policy_ref=sha, evidence_ref=sha,
                      expected_candidate_sha=sha, require_pass=True)

    short = copy.deepcopy(record)
    short["candidate_sha"] = sha[:7]
    expect_fail(root, short, "short candidate SHA must fail", policy_ref=sha)

    wrong_workflow_sha = copy.deepcopy(record)
    wrong_workflow_sha["runner"]["workflow"]["sha"] = "f" * 40
    expect_fail(root, wrong_workflow_sha, "workflow SHA must equal candidate SHA",
                policy_ref=sha)

    wrong_count = copy.deepcopy(record)
    wrong_count["harness"]["assertions"] = 2
    wrong_count["harness"]["passed"] = 2
    expect_fail(root, wrong_count, "assertion-count drift must fail", policy_ref=sha)

    bad_cleanup = copy.deepcopy(record)
    bad_cleanup["stages"]["cleanup"]["status"] = "FAIL"
    expect_fail(root, bad_cleanup, "green certification must reject cleanup failure",
                policy_ref=sha, require_pass=True)
    E.validate_record(bad_cleanup, root, policy_ref=sha, require_pass=False)

    failed_host = copy.deepcopy(record)
    failed_host["environment"]["excel_version"] = "unavailable"
    failed_host["environment"]["excel_build"] = "unavailable"
    failed_host["environment"]["office_bitness"] = "unavailable"
    failed_host["stages"]["import"] = {"status": "FAIL", "detail": "Excel unavailable"}
    failed_host["stages"]["compile"] = {"status": "NOT_RUN", "detail": "Import did not pass"}
    failed_host["stages"]["regression"] = {"status": "NOT_RUN", "detail": "Compile did not pass"}
    failed_host["stages"]["cleanup"] = {"status": "PASS", "detail": "No Excel process was created"}
    failed_host["harness"] = {"entry_point": policy["entry_point"], "assertions": 0,
                              "passed": 0, "failed": 0}
    E.validate_record(failed_host, root, policy_ref=sha, require_pass=False)
    expect_fail(root, failed_host, "failed-host record must not validate as green",
                policy_ref=sha, require_pass=True)

    bad_source = copy.deepcopy(record)
    bad_source["sources"][0]["sha256"] = "sha256:" + "f" * 64
    expect_fail(root, bad_source, "source digest drift must fail", policy_ref=sha)

    exported = good_record(root, sha, policy, export_main=True)
    E.validate_record(exported, root, policy_ref=sha, evidence_ref=sha,
                      require_pass=True, require_grid="benchmark/main.csv")

    wrong_grid = copy.deepcopy(exported)
    wrong_grid["grids"]["benchmark/main.csv"]["sha256"] = "sha256:" + "e" * 64
    expect_fail(root, wrong_grid, "exported grid digest mismatch must fail",
                policy_ref=sha, evidence_ref=sha, require_grid="benchmark/main.csv")

    wrong_rows = copy.deepcopy(exported)
    wrong_rows["grids"]["benchmark/main.csv"]["row_count"] += 1
    expect_fail(root, wrong_rows, "exported grid row-count mismatch must fail",
                policy_ref=sha, evidence_ref=sha, require_grid="benchmark/main.csv")

    regression_only = good_record(root, sha, policy)
    expect_fail(root, regression_only, "regression-only record cannot authorize grid binding",
                policy_ref=sha, evidence_ref=sha, require_grid="benchmark/main.csv")

    write(root, "benchmark/note.md", "evidence-only\n")
    evidence_sha = commit(root, "evidence-only")
    E.validate_record(record, root, policy_ref=evidence_sha, evidence_ref=evidence_sha,
                      require_pass=True)
    write(root, "src/A.bas", "Attribute VB_Name = \"A\"\n' drift\n")
    drift_sha = commit(root, "source drift")
    expect_fail(root, record, "later source drift must invalidate retained record",
                policy_ref=drift_sha, evidence_ref=drift_sha, require_pass=True)
finally:
    shutil.rmtree(root, ignore_errors=True)

if FAILS:
    print("FAIL: Excel certification fixtures")
    for item in FAILS:
        print("  - " + item)
    raise SystemExit(1)
print("PASS: Excel certification fixtures (exact candidate/workflow SHA, source inventory, "
      "assertion count, non-green host evidence, cleanup, grid digest/rows/target, and "
      "later-source drift)")
