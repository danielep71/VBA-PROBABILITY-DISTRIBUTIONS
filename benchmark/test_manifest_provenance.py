"""
Fixtures for check_manifest_provenance.py.

Covers the cases required when the guard was commissioned:

  * 9fba175 fails - the real rebind;
  * 228337e passes through grid co-change - a real export;
  * c496f1b passes only as an exact restoration;
  * byte-identical grid plus a valid export record passes;
  * export record changed but missing or mismatched hash fails;
  * holdout equivalents of each;
  * a two-commit push where different commits change the manifest and the
    grid fails - the aggregate diff would hide it;
  * bare write_manifest.py and ordinary refresh_evidence.py write nothing.

Synthetic cases are built in a throwaway git repository under a temporary
directory. Nothing here touches the real repository's history, and no
manifest, grid, contract or observation in this repository is modified.

Run: python3 test_manifest_provenance.py   (exit 0 = pass, nonzero = fail)
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import check_manifest_provenance as G                                # noqa: E402

fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


# --- the three real commits -------------------------------------------------
check(G.check_commit("9fba175") != [], "9fba175 (the real rebind) must fail")
check(G.check_commit("228337e") == [], "228337e (a real export) must pass")
check(G.check_commit("c496f1b") == [], "c496f1b (the repair) must pass")

# c496f1b must pass ONLY as a restoration, not by any other route.
_out = G.git("show", "--name-only", "--format=", "c496f1b").stdout
_changed = {ln.strip() for ln in _out.splitlines() if ln.strip()}
check(_changed == {"benchmark/observation_manifest.json"},
      "c496f1b must be a manifest-only commit for the restoration path to apply")
_ok, _why = G.is_exact_restoration("c496f1b", "benchmark/observation_manifest.json",
                                   _changed)
check(_ok and "restoration" in _why,
      f"c496f1b must classify as an exact restoration, got: {_why}")

# 9fba175 must NOT be reachable by the restoration path.
_out = G.git("show", "--name-only", "--format=", "9fba175").stdout
_ch = {ln.strip() for ln in _out.splitlines() if ln.strip()}
_ok, _ = G.is_exact_restoration("9fba175", "benchmark/observation_manifest.json", _ch)
check(not _ok, "9fba175 must not qualify as a restoration")


# --- synthetic repository ---------------------------------------------------
def run(cwd, *args):
    return subprocess.run(list(args), cwd=cwd, capture_output=True, text=True,
                          check=False)


def build_repo(tmp):
    run(tmp, "git", "init", "-q")
    run(tmp, "git", "config", "user.email", "t@t")
    run(tmp, "git", "config", "user.name", "t")
    for rel in ("benchmark/holdout",):
        os.makedirs(os.path.join(tmp, rel), exist_ok=True)
    grid = "benchmark/probability_accuracy_grid.csv"
    hgrid = "benchmark/holdout/holdout_grid.csv"
    write(tmp, grid, "function,arg1\nA,1\nB,2\n")
    write(tmp, hgrid, "function,arg1\nC,3\n")
    write(tmp, "benchmark/observation_manifest.json", json.dumps({"v": 1}) + "\n")
    write(tmp, "benchmark/holdout/holdout_manifest.json", json.dumps({"v": 1}) + "\n")
    run(tmp, "git", "add", "-A")
    run(tmp, "git", "commit", "-qm", "base")


def write(tmp, rel, text):
    path = os.path.join(tmp, rel)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="\n") as f:
        f.write(text)


def commit(tmp, msg):
    run(tmp, "git", "add", "-A")
    run(tmp, "git", "commit", "-qm", msg)
    return run(tmp, "git", "rev-parse", "HEAD").stdout.strip()


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


def sha_and_rows(tmp, rel):
    path = os.path.join(tmp, rel)
    return G.sha256_of(path), G.row_count_of(path)


def export_record(tmp, targets, break_hash=False, break_rows=False,
                  drop_field=None, mark_not_exported=False):
    grids = {}
    for rel in targets:
        sha, rows = sha_and_rows(tmp, rel)
        if break_hash:
            sha = "sha256:" + "0" * 64
        if break_rows:
            rows = rows + 7
        entry = {"exported": not mark_not_exported, "sha256": sha,
                 "row_count": rows}
        grids[rel] = entry
    rec = {"export_session_utc": "2026-09-02T12:00:00Z",
           "source_commit_sha": "deadbee", "excel_version": "16.0",
           "excel_build": "20131", "office_bitness": "64",
           "regression_total": 909, "regression_passed": 909,
           "regression_failed": 0, "grids": grids}
    if drop_field:
        rec.pop(drop_field, None)
    return json.dumps(rec, indent=2) + "\n"


MAIN_M = "benchmark/observation_manifest.json"
MAIN_G = "benchmark/probability_accuracy_grid.csv"
HOLD_M = "benchmark/holdout/holdout_manifest.json"
HOLD_G = "benchmark/holdout/holdout_grid.csv"


def case_grid_cochange(tmp):
    write(tmp, MAIN_G, "function,arg1\nA,1\nB,2\nD,4\n")
    write(tmp, MAIN_M, json.dumps({"v": 2}) + "\n")
    return G.check_commit(commit(tmp, "export"))


def case_record_valid(tmp):
    # Byte-identical grid, valid export record -> must pass.
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
    write(tmp, G.EXPORT_RECORD,
          export_record(tmp, [MAIN_G], drop_field="excel_build"))
    return G.check_commit(commit(tmp, "missing field"))


def case_record_wrong_target(tmp):
    # Record identifies only the holdout grid; main manifest changed.
    write(tmp, MAIN_M, json.dumps({"v": 7}) + "\n")
    write(tmp, G.EXPORT_RECORD, export_record(tmp, [HOLD_G]))
    return G.check_commit(commit(tmp, "wrong target"))


def case_record_not_exported(tmp):
    write(tmp, MAIN_M, json.dumps({"v": 8}) + "\n")
    write(tmp, G.EXPORT_RECORD,
          export_record(tmp, [MAIN_G], mark_not_exported=True))
    return G.check_commit(commit(tmp, "exported=false"))


def case_holdout_cochange(tmp):
    write(tmp, HOLD_G, "function,arg1\nC,3\nE,5\n")
    write(tmp, HOLD_M, json.dumps({"v": 2}) + "\n")
    return G.check_commit(commit(tmp, "holdout export"))


def case_holdout_rebind(tmp):
    write(tmp, HOLD_M, json.dumps({"v": 9}) + "\n")
    write(tmp, "benchmark/notes.md", "unrelated\n")
    return G.check_commit(commit(tmp, "holdout rebind"))


def case_holdout_record_bad_hash(tmp):
    write(tmp, HOLD_M, json.dumps({"v": 10}) + "\n")
    write(tmp, G.EXPORT_RECORD, export_record(tmp, [HOLD_G], break_hash=True))
    return G.check_commit(commit(tmp, "holdout bad hash"))


def case_restoration(tmp):
    original = open(os.path.join(tmp, MAIN_M)).read()
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
    # THE CASE AN AGGREGATE DIFF HIDES: one commit changes the grid, another
    # rebinds the manifest. Per-commit evaluation must fail the second.
    write(tmp, MAIN_G, "function,arg1\nA,1\nB,2\nZ,9\n")
    c1 = commit(tmp, "grid only")
    write(tmp, MAIN_M, json.dumps({"v": 77}) + "\n")
    c2 = commit(tmp, "manifest only")
    return G.check_commit(c1), G.check_commit(c2)


check(with_repo(case_grid_cochange) == [], "grid co-change must pass")
check(with_repo(case_record_valid) == [],
      "byte-identical grid with a valid export record must pass")
check(with_repo(case_record_bad_hash) != [],
      "export record with a mismatched grid hash must fail")
check(with_repo(case_record_bad_rows) != [],
      "export record with a mismatched row count must fail")
check(with_repo(case_record_missing_field) != [],
      "export record missing a required field must fail")
check(with_repo(case_record_wrong_target) != [],
      "export record not identifying the manifest's grid must fail")
check(with_repo(case_record_not_exported) != [],
      "export record marking the grid as not exported must fail")
check(with_repo(case_holdout_cochange) == [],
      "holdout grid co-change must pass")
check(with_repo(case_holdout_rebind) != [],
      "holdout manifest rebind must fail")
check(with_repo(case_holdout_record_bad_hash) != [],
      "holdout export record with a bad hash must fail")
_bad, _good = with_repo(case_restoration)
check(_bad != [], "synthetic rebind must fail")
check(_good == [], "exact restoration must pass")
check(with_repo(case_manifest_only_novel) != [],
      "manifest-only commit with never-before-seen content must fail")
_c1, _c2 = with_repo(case_two_commit_push)
check(_c1 == [], "grid-only commit is fine on its own")
check(_c2 != [],
      "two-commit push: the manifest-only commit must fail even though the "
      "push's aggregate diff contains the grid")

# --- the writer and the refresh path must not write -------------------------
_bare = subprocess.run([sys.executable, "write_manifest.py"], cwd=HERE,
                       capture_output=True, text=True)
check(_bare.returncode != 0, "bare write_manifest.py must exit non-zero")
check("from-fresh-export" in (_bare.stderr + _bare.stdout),
      "bare write_manifest.py must name the required flag")
_dry = subprocess.run([sys.executable, "write_manifest.py", "--dry-run"],
                      cwd=HERE, capture_output=True, text=True)
check(_dry.returncode == 0 and "nothing was written" in _dry.stdout,
      "--dry-run must preview and exit 0")

import refresh_evidence as R                                          # noqa: E402
check(all("write_manifest.py" not in argv for _, _, argv in R.REGENERATE),
      "ordinary regeneration must not invoke write_manifest.py")
check("--from-fresh-export" in R.MAIN_BINDING[2],
      "main binding must pass --from-fresh-export")
check("--from-fresh-export" in R.HOLDOUT_BINDING[2],
      "holdout binding must pass --from-fresh-export")

if fails:
    print("FAIL: manifest provenance guard")
    for f_ in fails:
        print("  - " + f_)
    raise SystemExit(1)
print("PASS: manifest provenance guard (9fba175 fails, 228337e passes by grid "
      "co-change, c496f1b passes only as restoration; export-record exception "
      "validated for hash, rows, fields, target and exported flag; holdout "
      "equivalents; two-commit push caught; writer and refresh write nothing)")
