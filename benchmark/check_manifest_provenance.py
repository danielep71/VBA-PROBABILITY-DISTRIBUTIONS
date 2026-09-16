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
     a real export normally looks like.

  B. canonical Excel certification record - benchmark/excel_regression_record.json
     is also modified and validates under excel_certification.py. The record
     must bind a full candidate SHA, the exact policy-selected VBA source bytes,
     a green import/compile/regression/cleanup run, and explicitly claim THIS
     grid as freshly exported with matching SHA-256 and row count. A green
     regression-only record is not sufficient.

  C. exact restoration - the commit modifies the manifest and nothing else,
     and the resulting manifest is byte-identical to some earlier commit's
     version of it. This is the repair path used by c496f1b.

Run: python3 check_manifest_provenance.py [--since <rev>]
"""
import os
import subprocess
import sys

from excel_certification import (CertificationError, row_count_file,
                                 sha256_file, validate_record_text)

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

MANIFESTS = {
    "benchmark/observation_manifest.json": "benchmark/probability_accuracy_grid.csv",
    "benchmark/holdout/holdout_manifest.json": "benchmark/holdout/holdout_grid.csv",
}
EXPORT_RECORD = "benchmark/excel_regression_record.json"


def git(*args):
    return subprocess.run(["git"] + list(args), cwd=ROOT, capture_output=True,
                          text=True, check=False)


def sha256_of(path):
    return sha256_file(path)


def row_count_of(path):
    return row_count_file(path)


def validate_export_record(text, manifest_path, commit):
    """Return (ok, reason) for the canonical record at one evidence commit."""
    grid_rel = MANIFESTS[manifest_path]
    try:
        validate_record_text(
            text,
            ROOT,
            policy_ref=commit,
            evidence_ref=commit,
            require_pass=True,
            require_grid=grid_rel,
        )
    except CertificationError as exc:
        return False, str(exc)
    return True, ""


def is_exact_restoration(commit, manifest_path, changed):
    """Manifest-only commit whose result equals an earlier committed version."""
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
        if EXPORT_RECORD in changed:                   # B: canonical record
            text = git("show", f"{commit}:{EXPORT_RECORD}").stdout
            ok, why = validate_export_record(text, manifest_path, commit)
            if ok:
                continue
            problems.append(f"{commit[:7]}: {manifest_path} changed with an "
                            f"Excel certification record that does not validate - {why}")
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
          "separately; a manifest changes only with its grid, a canonical "
          "fresh-export Excel certification record, or as an exact restoration)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
