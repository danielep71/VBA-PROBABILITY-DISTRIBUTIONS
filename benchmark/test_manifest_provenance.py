"""Fixtures for check_manifest_provenance.py and the certification bridge."""
import json
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import check_manifest_provenance as G  # noqa: E402
import excel_certification as E        # noqa: E402

fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


# Historical regression controls that motivated the per-commit guard.
check(G.check_commit("9fba175") != [], "9fba175 (the real rebind) must fail")
check(G.check_commit("228337e") == [], "228337e (a real export) must pass")
check(G.check_commit("c496f1b") == [], "c496f1b (the repair) must pass")

_out = G.git("show", "--name-only", "--format=", "c496f1b").stdout
_changed = {ln.strip() for ln in _out.splitlines() if ln.strip()}
_ok, _why = G.is_exact_restoration("c496f1b", "benchmark/observation_manifest.json",
                                   _changed)
check(_changed == {"benchmark/observation_manifest.json"} and _ok,
      f"c496f1b must pass only as an exact restoration, got: {_why}")

_out = G.git("show", "--name-only", "--format=", "9fba175").stdout
_ch = {ln.strip() for ln in _out.splitlines() if ln.strip()}
_ok, _ = G.is_exact_restoration("9fba175", "benchmark/observation_manifest.json", _ch)
check(not _ok, "9fba175 must not qualify as a restoration")


def run(cwd, *args):
    return subprocess.run(list(args), cwd=cwd, capture_output=True, check=False)


def write(tmp, rel, text):
    path = os.path.join(tmp, rel)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)


def commit(tmp, msg):
    run(tmp, "git", "add", "-A")
    run(tmp, "git", "commit", "-qm", msg)
    return run(tmp, "git", "rev-parse", "HEAD").stdout.decode().strip()


MAIN_M = "benchmark/observation_manifest.json"
MAIN_G = "benchmark/probability_accuracy_grid.csv"
HOLD_M = "benchmark/holdout/holdout_manifest.json"
HOLD_G = "benchmark/holdout/holdout_grid.csv"
POLICY = {
    "schema_version": 1,
    "repository": "owner/repo",
    "workflow": ".github/workflows/excel.yml",
    "entry_point": "RunAll_CI",
    "expected_assertions": 3,
    "regression_sources": ["src/A.bas", "tests/T.bas"],
    "grid_exporters": {
        MAIN_G: "benchmark/MainExport.bas",
        HOLD_G: "benchmark/holdout/HoldoutExport.bas",
    },
}


def build_repo(tmp):
    run(tmp, "git", "init", "-q")
    run(tmp, "git", "config", "user.email", "t@t")
    run(tmp, "git", "config", "user.name", "t")
    write(tmp, E.POLICY_PATH, json.dumps(POLICY, indent=2) + "\n")
    write(tmp, "src/A.bas", "Attribute VB_Name = \"A\"\n")
    write(tmp, "tests/T.bas", "Attribute VB_Name = \"T\"\n")
    write(tmp, "benchmark/MainExport.bas", "Attribute VB_Name = \"MainExport\"\n")
    write(tmp, "benchmark/holdout/HoldoutExport.bas", "Attribute VB_Name = \"HoldoutExport\"\n")
    write(tmp, MAIN_G, "function,arg1\nA,1\nB,2\n")
    write(tmp, HOLD_G, "function,arg1\nC,3\n")
    write(tmp, MAIN_M, json.dumps({"v": 1}) + "\n")
    write(tmp, HOLD_M, json.dumps({"v": 1}) + "\n")
    commit(tmp, "base")


def with_repo(fn):
    tmp = tempfile.mkdtemp()
    saved_root = G.ROOT
    try:
        build_repo(tmp)
        G.ROOT = tmp
        return fn(tmp)
    finally:
        G.ROOT = saved_root
        shutil.rmtree(tmp, ignore_errors=True)


def digest_at(tmp, sha, rel):
    return E.sha256_bytes(E.git_bytes(tmp, sha, rel))


def export_record(tmp, targets, break_hash=False, break_rows=False,
                  drop_field=None, mark_not_exported=False, source_drift=False):
    candidate = run(tmp, "git", "rev-parse", "HEAD").stdout.decode().strip()
    grids = {
        grid: {"exported": False, "sha256": None, "row_count": None,
               "exported_utc": None}
        for grid in POLICY["grid_exporters"]
    }
    source_paths = set(POLICY["regression_sources"])
    for rel in targets:
        raw = open(os.path.join(tmp, rel), "rb").read()
        digest = E.sha256_bytes(raw)
        rows = E.row_count_bytes(raw)
        if break_hash:
            digest = "sha256:" + "0" * 64
        if break_rows:
            rows += 7
        grids[rel] = {"exported": not mark_not_exported,
                      "sha256": digest if not mark_not_exported else None,
                      "row_count": rows if not mark_not_exported else None,
                      "exported_utc": "2026-09-16T04:00:01Z" if not mark_not_exported else None}
        if not mark_not_exported:
            source_paths.add(POLICY["grid_exporters"][rel])
    sources = [{"path": rel, "sha256": digest_at(tmp, candidate, rel)}
               for rel in sorted(source_paths)]
    if source_drift:
        sources[0]["sha256"] = "sha256:" + "f" * 64
    rec = {
        "schema_version": 1,
        "repository": POLICY["repository"],
        "candidate_sha": candidate,
        "execution": "automated",
        "started_at": "2026-09-16T04:00:00Z",
        "finished_at": "2026-09-16T04:00:02Z",
        "runner": {"class": "self-hosted-excel", "identity": "PROBDIST",
                   "workflow": {"repository": POLICY["repository"],
                                "path": POLICY["workflow"], "sha": candidate,
                                "run_id": 1, "run_attempt": 1}},
        "environment": {"excel_version": "16.0", "excel_build": "20228",
                        "office_bitness": "64-bit", "windows": "Windows 11",
                        "os_architecture": "X64", "automation_security": "1 (isolated process)",
                        "vba_project_access": "preconfigured runner"},
        "sources": sources,
        "stages": {name: {"status": "PASS", "detail": "ok"} for name in E.STAGES},
        "harness": {"entry_point": POLICY["entry_point"], "assertions": 3,
                    "passed": 3, "failed": 0},
        "log": {"path": "test-result.txt", "sha256": "sha256:" + "0" * 64},
        "grids": grids,
    }
    if drop_field:
        rec.pop(drop_field, None)
    return json.dumps(rec, indent=2) + "\n"


def case_grid_cochange(tmp):
    write(tmp, MAIN_G, "function,arg1\nA,1\nB,2\nD,4\n")
    write(tmp, MAIN_M, json.dumps({"v": 2}) + "\n")
    return G.check_commit(commit(tmp, "export"))


def case_record_valid(tmp):
    write(tmp, MAIN_M, json.dumps({"v": 3}) + "\n")
    write(tmp, G.EXPORT_RECORD, export_record(tmp, [MAIN_G]))
    return G.check_commit(commit(tmp, "reexport, identical grid"))


def case_record_bad_hash(tmp):
    write(tmp, MAIN_M, json.dumps({"v": 4}) + "\n")
    write(tmp, G.EXPORT_RECORD, export_record(tmp, [MAIN_G], break_hash=True))
    return G.check_commit(commit(tmp, "bad hash"))


def case_record_bad_rows(tmp):
    write(tmp, MAIN_M, json.dumps({"v": 5}) + "\n")
    write(tmp, G.EXPORT_RECORD, export_record(tmp, [MAIN_G], break_rows=True))
    return G.check_commit(commit(tmp, "bad row count"))


def case_record_missing_field(tmp):
    write(tmp, MAIN_M, json.dumps({"v": 6}) + "\n")
    write(tmp, G.EXPORT_RECORD, export_record(tmp, [MAIN_G], drop_field="environment"))
    return G.check_commit(commit(tmp, "missing field"))


def case_record_wrong_target(tmp):
    write(tmp, MAIN_M, json.dumps({"v": 7}) + "\n")
    write(tmp, G.EXPORT_RECORD, export_record(tmp, [HOLD_G]))
    return G.check_commit(commit(tmp, "wrong target"))


def case_record_not_exported(tmp):
    write(tmp, MAIN_M, json.dumps({"v": 8}) + "\n")
    write(tmp, G.EXPORT_RECORD, export_record(tmp, [MAIN_G], mark_not_exported=True))
    return G.check_commit(commit(tmp, "exported=false"))


def case_record_source_drift(tmp):
    write(tmp, MAIN_M, json.dumps({"v": 9}) + "\n")
    write(tmp, G.EXPORT_RECORD, export_record(tmp, [MAIN_G], source_drift=True))
    return G.check_commit(commit(tmp, "bad source binding"))


def case_holdout_valid(tmp):
    write(tmp, HOLD_M, json.dumps({"v": 2}) + "\n")
    write(tmp, G.EXPORT_RECORD, export_record(tmp, [HOLD_G]))
    return G.check_commit(commit(tmp, "holdout identical reexport"))


def case_holdout_rebind(tmp):
    write(tmp, HOLD_M, json.dumps({"v": 9}) + "\n")
    write(tmp, "benchmark/notes.md", "unrelated\n")
    return G.check_commit(commit(tmp, "holdout rebind"))


def case_holdout_record_bad_hash(tmp):
    write(tmp, HOLD_M, json.dumps({"v": 10}) + "\n")
    write(tmp, G.EXPORT_RECORD, export_record(tmp, [HOLD_G], break_hash=True))
    return G.check_commit(commit(tmp, "holdout bad hash"))


def case_restoration(tmp):
    with open(os.path.join(tmp, MAIN_M), encoding="utf-8") as f:
        original = f.read()
    write(tmp, MAIN_M, json.dumps({"v": 99}) + "\n")
    write(tmp, "benchmark/other.md", "x\n")
    bad = commit(tmp, "rebind")
    write(tmp, MAIN_M, original)
    good = commit(tmp, "restore")
    return G.check_commit(bad), G.check_commit(good)


def case_manifest_only_novel(tmp):
    write(tmp, MAIN_M, json.dumps({"v": 123}) + "\n")
    return G.check_commit(commit(tmp, "manifest-only novel content"))


def case_two_commit_push(tmp):
    write(tmp, MAIN_G, "function,arg1\nA,1\nB,2\nZ,9\n")
    c1 = commit(tmp, "grid only")
    write(tmp, MAIN_M, json.dumps({"v": 77}) + "\n")
    c2 = commit(tmp, "manifest only")
    return G.check_commit(c1), G.check_commit(c2)


check(with_repo(case_grid_cochange) == [], "grid co-change must pass")
check(with_repo(case_record_valid) == [], "canonical exact-SHA export record must pass")
check(with_repo(case_record_bad_hash) != [], "mismatched grid hash must fail")
check(with_repo(case_record_bad_rows) != [], "mismatched grid row count must fail")
check(with_repo(case_record_missing_field) != [], "missing certification field must fail")
check(with_repo(case_record_wrong_target) != [], "wrong exported target must fail")
check(with_repo(case_record_not_exported) != [], "regression-only record must not authorize rebind")
check(with_repo(case_record_source_drift) != [], "wrong certified source hash must fail")
check(with_repo(case_holdout_valid) == [], "canonical holdout export record must pass")
check(with_repo(case_holdout_rebind) != [], "holdout rebind must fail")
check(with_repo(case_holdout_record_bad_hash) != [], "holdout bad hash must fail")
_bad, _good = with_repo(case_restoration)
check(_bad != [], "synthetic rebind must fail")
check(_good == [], "exact restoration must pass")
check(with_repo(case_manifest_only_novel) != [], "novel manifest-only content must fail")
_c1, _c2 = with_repo(case_two_commit_push)
check(_c1 == [], "grid-only commit is fine on its own")
check(_c2 != [], "two-commit push must catch the manifest-only second commit")

_bare = subprocess.run([sys.executable, "write_manifest.py"], cwd=HERE,
                       capture_output=True, text=True)
check(_bare.returncode != 0, "bare write_manifest.py must exit non-zero")
check("from-fresh-export" in (_bare.stderr + _bare.stdout),
      "bare write_manifest.py must name the required flag")
_dry = subprocess.run([sys.executable, "write_manifest.py", "--dry-run"],
                      cwd=HERE, capture_output=True, text=True)
check(_dry.returncode == 0 and "nothing was written" in _dry.stdout,
      "--dry-run must preview and exit 0")

import refresh_evidence as R  # noqa: E402
check(all("write_manifest.py" not in argv for _, _, argv in R.REGENERATE),
      "ordinary regeneration must not invoke write_manifest.py")
check("--from-fresh-export" in R.MAIN_BINDING[2],
      "main binding must pass --from-fresh-export")
check("--from-fresh-export" in R.HOLDOUT_BINDING[2],
      "holdout binding must pass --from-fresh-export")

if fails:
    print("FAIL: manifest provenance guard")
    for item in fails:
        print("  - " + item)
    raise SystemExit(1)
print("PASS: manifest provenance guard (historical rebind/export/restoration controls; "
      "canonical exact-SHA record checks source, target, hash, rows and exported flag; "
      "holdout equivalents; two-commit push caught; writer and refresh write nothing)")
