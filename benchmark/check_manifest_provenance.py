"""
Per-commit guard on the provenance manifests.

WHY

A manifest asserts "these committed observations were produced by this
source". Only a fresh Excel export makes that true. In 9fba175 a bare
write_manifest.py rebound benchmark/observation_manifest.json to source three
commits newer than the observations: the seven-line signature looked
unchanged, but the strict gate's failure had moved from STALE EVIDENCE (2
mismatches, truthful) to STALE HOLDOUT EVIDENCE, and the main binding read
clean. One `write_manifest.py --holdout` away from a green gate on stale
evidence. Nothing caught it - the source-binding verifier cannot, because it
hashes source and observations, both present either way, and the manifest's
own environment block comes from a committed file with no timestamp.

WHAT IS CHECKED

Each commit is examined SEPARATELY, never the aggregate diff of a push. An
aggregate check is defeated by a two-commit push in which one commit touches
the grid and another rebinds the manifest.

A commit that modifies a manifest is legal only if, in that SAME commit, one
of the following holds:

  A. grid co-change - the manifest's own grid is also modified. This is what
     a real export looks like: 228337e changed the grid, the summary and the
     manifest together.

  B. validated export record - benchmark/excel_regression_record.json is also
     modified AND validates under _export_record.validate(). Present so that
     a re-export producing a byte-identical grid (a .bas change that alters no
     observation value) is still legal. The record must IDENTIFY the exported
     target and bind the fields in EXPORT_RECORD_REQUIRED below, including the
     SHA-256 and row count of each grid it claims to have exported, matching
     the committed grid. "The file changed" is not enough.

  C. exact restoration - the commit modifies the manifest and nothing else,
     and the resulting manifest is byte-identical to some earlier commit's
     version of it. This is the repair path: c496f1b restored the manifest
     that 9fba175 corrupted while the grid was untouched. It cannot launder an
     ordinary rebind, because a rebind produces a manifest that has never
     existed before.

Run: python3 check_manifest_provenance.py [--since <rev>]
"""
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

MANIFESTS = {
    "benchmark/observation_manifest.json": "benchmark/probability_accuracy_grid.csv",
    "benchmark/holdout/holdout_manifest.json": "benchmark/holdout/holdout_grid.csv",
}
EXPORT_RECORD = "benchmark/excel_regression_record.json"

# Fields the export record must bind. Refinement 2: the exception is validated,
# not merely "a file changed".
EXPORT_RECORD_REQUIRED = (
    "export_session_utc",       # when the export ran
    "source_commit_sha",        # source identity
    "excel_version",
    "excel_build",
    "office_bitness",
    "regression_total",         # regression totals
    "regression_passed",
    "regression_failed",
    "grids",                    # per-grid: exported, sha256, row_count
)
GRID_RECORD_REQUIRED = ("exported", "sha256", "row_count")


def git(*args):
    return subprocess.run(["git"] + list(args), cwd=ROOT, capture_output=True,
                          text=True, check=False)


def sha256_of(path):
    import hashlib
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(65536), b""):
            h.update(block)
    return "sha256:" + h.hexdigest()


def row_count_of(path):
    with open(path, encoding="utf-8", newline="") as f:
        return max(sum(1 for _ in f) - 1, 0)


def validate_export_record(text, manifest_path):
    """Return (ok, reason). The record must claim THIS manifest's grid as
    exported, and its hash and row count must match the committed grid."""
    try:
        rec = json.loads(text)
    except ValueError as exc:
        return False, f"export record is not valid JSON: {exc}"
    missing = [k for k in EXPORT_RECORD_REQUIRED if k not in rec]
    if missing:
        return False, f"export record missing field(s): {', '.join(missing)}"
    grids = rec.get("grids")
    if not isinstance(grids, dict) or not grids:
        return False, "export record 'grids' must be a non-empty object"

    grid_rel = MANIFESTS[manifest_path]
    entry = grids.get(grid_rel)
    if entry is None:
        return False, (f"export record does not identify {grid_rel} as an "
                       "exported target")
    missing = [k for k in GRID_RECORD_REQUIRED if k not in entry]
    if missing:
        return False, (f"export record entry for {grid_rel} missing: "
                       f"{', '.join(missing)}")
    if entry["exported"] is not True:
        return False, (f"export record marks {grid_rel} as not exported; a "
                       "manifest may not be bound to a grid that was not "
                       "re-exported")
    grid_abs = os.path.join(ROOT, grid_rel)
    if not os.path.exists(grid_abs):
        return False, f"{grid_rel} is missing from the working tree"
    actual_sha = sha256_of(grid_abs)
    if entry["sha256"] != actual_sha:
        return False, (f"export record hash for {grid_rel} does not match the "
                       f"committed grid: recorded {entry['sha256'][:23]}..., "
                       f"actual {actual_sha[:23]}...")
    actual_rows = row_count_of(grid_abs)
    if int(entry["row_count"]) != actual_rows:
        return False, (f"export record row count for {grid_rel} is "
                       f"{entry['row_count']}, committed grid has {actual_rows}")
    return True, ""


def is_exact_restoration(commit, manifest_path, changed):
    """Refinement 3: manifest-only commit whose result equals an earlier
    committed version of that manifest byte for byte."""
    if changed != {manifest_path}:
        return False, "not a manifest-only commit"
    blob = git("rev-parse", f"{commit}:{manifest_path}").stdout.strip()
    if not blob:
        return False, "manifest missing at this commit"
    history = git("rev-list", f"{commit}~1", "--", manifest_path).stdout.split()
    for earlier in history:
        prior = git("rev-parse", f"{earlier}:{manifest_path}").stdout.strip()
        if prior == blob:
            return True, f"exact restoration of the manifest as at {earlier[:7]}"
    return False, ("manifest-only commit whose content has never existed "
                   "before: this is a rebind, not a restoration")


def check_commit(commit):
    """Return a list of problems for one commit."""
    out = git("show", "--name-only", "--format=", commit).stdout
    changed = {ln.strip() for ln in out.splitlines() if ln.strip()}
    problems = []
    for manifest_path, grid_path in MANIFESTS.items():
        if manifest_path not in changed:
            continue
        if grid_path in changed:
            continue                                   # A: grid co-change
        if EXPORT_RECORD in changed:                   # B: validated record
            text = git("show", f"{commit}:{EXPORT_RECORD}").stdout
            ok, why = validate_export_record(text, manifest_path)
            if ok:
                continue
            problems.append(f"{commit[:7]}: {manifest_path} changed with an "
                            f"export record that does not validate - {why}")
            continue
        ok, why = is_exact_restoration(commit, manifest_path, changed)  # C
        if ok:
            continue
        problems.append(
            f"{commit[:7]}: {manifest_path} was modified without {grid_path}, "
            f"without a validated {EXPORT_RECORD}, and not as an exact "
            f"restoration ({why}). A manifest asserts the observations were "
            "produced by the checked-out source; only a fresh export makes "
            "that true.")
    return problems


def main():
    since = None
    if "--since" in sys.argv:
        since = sys.argv[sys.argv.index("--since") + 1]
    if since:
        rev = f"{since}..HEAD"
    else:
        # Default: the commits this push adds. On a hosted runner
        # GITHUB_EVENT_BEFORE marks the previous tip; fall back to HEAD alone,
        # which still checks each commit separately rather than a diff range.
        before = os.environ.get("GITHUB_EVENT_BEFORE", "")
        rev = (f"{before}..HEAD" if before and not set(before) == {"0"}
               else "HEAD~1..HEAD")
    res = git("rev-list", rev)
    if res.returncode != 0:
        commits = [git("rev-parse", "HEAD").stdout.strip()]
    else:
        commits = [c for c in res.stdout.split() if c]
    if not commits:
        print("PASS: manifest provenance (no commits in range)")
        return 0

    problems = []
    for c in commits:
        problems.extend(check_commit(c))
    if problems:
        print("FAIL: manifest provenance")
        for p in problems:
            print("  - " + p)
        return 1
    print(f"PASS: manifest provenance ({len(commits)} commit(s) checked "
          "separately; a manifest changes only with its grid, a validated "
          "export record, or as an exact restoration)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
