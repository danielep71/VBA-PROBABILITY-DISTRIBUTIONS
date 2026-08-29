"""
Regenerate every committed evidence artifact, then run every check the hosted
Accuracy Gate runs, in the same order.

WHY THIS EXISTS
    The regeneration steps have a mandatory order and an easily-forgotten member.
    Two real failures came from running only part of the sequence by hand:

      * accuracy_summary.md generated BEFORE write_manifest.py, so it embedded
        the previous commit SHA and CI's freshness diff failed;
      * numerical_limitations.csv edited without re-rendering benchmark/README.md,
        which is generated from it, so CI's README diff failed.

    Both were caught by CI rather than locally, because a hand-run sequence
    covered some steps and not others. This script runs the whole set, so
    "green here" means "green there".

WHAT IT DOES NOT DO
    It cannot re-export observations - that requires Excel. It warns when the
    source appears to have changed without a fresh export, but the warning is a
    heuristic: only you know whether the export ran. Rebinding the manifest
    without re-exporting is worse than leaving it stale, because it certifies
    old measurements as belonging to new code.

USAGE
    python refresh_evidence.py            regenerate, then verify
    python refresh_evidence.py --check    verify only, change nothing
    python refresh_evidence.py --bind-exported-holdout
                                          bind a holdout just exported by Excel,
                                          then regenerate and verify
"""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
HOLDOUT = os.path.join(HERE, "holdout")

# (label, working directory, argv). Order matters: the manifest must be written
# before the summary, because the summary embeds the manifest's commit SHA and
# timestamp.
REGENERATE = [
    ("bind observations to source", HERE, ["write_manifest.py"]),
    ("accuracy summary", HERE, ["compute_errors.py", "--out", "accuracy_summary.md"]),
    ("benchmark README contract table", HERE, ["render_contract_table.py", "--write"]),
    ("independent holdout summary", HOLDOUT, ["analyze_holdout.py"]),
]

HOLDOUT_BINDING = (
    "bind independent holdout to source",
    HERE,
    ["write_manifest.py", "--holdout"],
)

# Mirrors .github/workflows/accuracy-gate.yml, in the same order. The evaluator
# tests run first on purpose: a broken evaluator makes every verdict below
# meaningless rather than merely wrong.
VERIFY = [
    ("evaluator unit tests", HERE, ["test_contract_eval.py"]),
    ("manifest unit tests", HERE, ["test_manifest.py"]),
    ("strict accuracy gate", HERE, ["compute_errors.py"]),
    ("gate blocks without references", HERE, ["test_gate_degradation.py"]),
    ("source claims vs registries", HERE, ["check_source_thresholds.py"]),
    ("holdout analyzer tests", HOLDOUT, ["test_analyze_holdout.py"]),
    # analyze_holdout.py WRITES holdout_summary.md as a side effect, so it is
    # not read-only. Under --check the summary is restored afterwards, keeping
    # that mode honest about changing nothing.
    ("independent holdout", HOLDOUT, ["analyze_holdout.py"]),
]

HOLDOUT_SUMMARY = "benchmark/holdout/holdout_summary.md"

ARTIFACTS = [
    "benchmark/accuracy_summary.md",
    "benchmark/README.md",
    "benchmark/observation_manifest.json",
    "benchmark/holdout/holdout_summary.md",
]


def run(cwd, argv):
    """Run a benchmark script with the interpreter that started this one."""
    proc = subprocess.run([sys.executable] + argv, cwd=cwd,
                          capture_output=True, text=True)
    return proc.returncode, (proc.stdout + proc.stderr).strip()


def git(*args):
    """
    Run git, or report politely that it is unavailable.

    git is NOT required by this project: GitHub Desktop bundles its own copy and
    does not put it on PATH, which is the supported setup here. write_manifest.py
    reads .git directly for exactly this reason. Every git call in this script is
    therefore a convenience - it names which artifacts changed - and the
    regeneration and all seven checks work without it.
    """
    try:
        proc = subprocess.run(["git"] + list(args), cwd=ROOT,
                              capture_output=True, text=True)
    except (FileNotFoundError, OSError):
        return None, ""
    return proc.returncode, proc.stdout.strip()


def warn_if_export_looks_missing():
    """
    Heuristic: a modified .bas alongside an unmodified observation grid usually
    means the Excel export has not been re-run. Warn, never block - the check
    cannot be certain, and a false block would be worse than a false warning.
    """
    code, out = git("status", "--porcelain")
    if code is None or code != 0:
        return
    # "XY path" - split rather than slice a fixed width, which mis-handles
    # staged entries and produced a truncated path in an earlier version.
    changed = {line.split(maxsplit=1)[1].strip()
               for line in out.splitlines() if line.strip() and " " in line}
    src_changed = any(p.endswith(".bas") and p.startswith("src/") for p in changed)
    grid_changed = "benchmark/probability_accuracy_grid.csv" in changed
    if src_changed and not grid_changed:
        print("  ! WARNING: a src/*.bas file is modified but the observation grid")
        print("             is not. If the change affects any computed value, the")
        print("             Excel export must be re-run BEFORE this script, or the")
        print("             manifest will certify old measurements as belonging to")
        print("             new code. Comment-only edits are fine.")
        print()


def main():
    check_only = "--check" in sys.argv
    bind_exported_holdout = "--bind-exported-holdout" in sys.argv
    unknown = set(sys.argv[1:]) - {"--check", "--bind-exported-holdout"}
    if unknown or (check_only and bind_exported_holdout):
        raise SystemExit("usage: refresh_evidence.py [--check | --bind-exported-holdout]")
    saved = None
    if check_only:
        path = os.path.join(ROOT, HOLDOUT_SUMMARY)
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                saved = f.read()

    if not check_only:
        warn_if_export_looks_missing()
        print("Regenerating committed artifacts")
        regenerate = list(REGENERATE)
        if bind_exported_holdout:
            # The flag is intentionally explicit.  The committed 559-row holdout
            # predates current source and must not be rebound merely because a
            # developer regenerated summaries.
            regenerate.insert(1, HOLDOUT_BINDING)
        for label, cwd, argv in regenerate:
            code, out = run(cwd, argv)
            # compute_errors exits non-zero on a failing gate; that is reported
            # by the verify pass below, so regeneration continues either way.
            print(f"  {'ok ' if code == 0 else '(!)'} {label}")
        print()

    print("Verifying, in the order the hosted Accuracy Gate runs")
    failures = []
    for label, cwd, argv in VERIFY:
        code, out = run(cwd, argv)
        print(f"  {'PASS' if code == 0 else 'FAIL'}  {label}")
        if code != 0:
            failures.append((label, out))

    if check_only and saved is not None:
        # Restore what analyze_holdout.py rewrote, so --check is genuinely
        # read-only rather than merely usually harmless.
        with open(os.path.join(ROOT, HOLDOUT_SUMMARY), "w",
                  encoding="utf-8", newline="\n") as f:
            f.write(saved)

    print()
    code, out = git("status", "--porcelain", *ARTIFACTS)
    if code is None:
        print("Artifacts regenerated. (git is not on PATH, so the changed-file")
        print("list is unavailable - commit whatever GitHub Desktop shows under")
        print("benchmark/.)")
        touched = []
    else:
        touched = [l.split(maxsplit=1)[1].strip()
                   for l in out.splitlines() if l.strip() and " " in l]
    if touched:
        print("Artifacts changed - commit these with the source:")
        for t in touched:
            print(f"  {t}")
        # A manifest whose only change is the recorded commit SHA means the
        # PREVIOUS commit shipped evidence bound to the commit before it. Not a
        # correctness problem - the source hashes still matched - but it is drift
        # worth naming rather than leaving the maintainer to infer.
        code, diff = git("diff", "--", "benchmark/observation_manifest.json")
        diff = diff or ""
        old_sha = new_sha = None
        for line in diff.splitlines():
            if '"commit_sha"' in line and line.startswith("-"):
                old_sha = line.split('"')[3]
            elif '"commit_sha"' in line and line.startswith("+"):
                new_sha = line.split('"')[3]
        if old_sha and new_sha and old_sha != new_sha:
            print()
            print(f"  note: the committed manifest recorded {old_sha}, but HEAD is")
            print(f"        {new_sha}. The last commit shipped evidence bound to the")
            print("        commit before it. Running this script as part of every")
            print("        evidence commit keeps the two in step.")
    elif code is not None:
        print("Artifacts already current - nothing to commit here.")

    if failures:
        print()
        print(f"{len(failures)} check(s) failed. CI will fail too:")
        for label, out in failures:
            first = out.splitlines()[0] if out else "(no output)"
            print(f"  {label}: {first}")
        sys.exit(1)

    print()
    print("All checks pass. CI should agree.")


if __name__ == "__main__":
    main()
