"""
Write benchmark/observation_manifest.json, binding the current committed
observations to the source that produced them.

Run this AT EXPORT TIME - immediately after re-exporting observations from the
current source in Excel - and commit the manifest together with the grid. The
strict accuracy gate then refuses to certify the summary if the checked-out
.bas files no longer match the recorded hashes (i.e. source drifted from the
evidence).

Provenance is recorded automatically where it can be read without guessing:

  * commit SHA - resolved from .git (HEAD, packed-refs) with no git executable,
    so it works from GitHub Desktop or a plain checkout.
  * Excel version / build / bitness - read from benchmark/excel_environment.json,
    written by the Export_ExcelEnvironment macro at export time.

Explicit flags still override both, and anything genuinely unavailable is
recorded as "unrecorded" rather than guessed:

    python write_manifest.py --excel-version 2408 --excel-build 16.0.17928.20216 \
                             --office-bitness 64 --commit-sha <sha>
"""
import argparse
import datetime
import json
import os

from _manifest import (build_manifest, repo_root, MANIFEST_NAME)

ENV_FILE = "excel_environment.json"


def _read_commit_sha(root):
    """Resolve HEAD to a short SHA by reading .git directly (no git executable)."""
    git = os.path.join(root, ".git")
    try:
        if os.path.isfile(git):                      # worktree/submodule pointer
            with open(git, encoding="utf-8") as f:
                git = os.path.join(root, f.read().split(":", 1)[1].strip())
        with open(os.path.join(git, "HEAD"), encoding="utf-8") as f:
            head = f.read().strip()
        if not head.startswith("ref:"):
            return head[:7]                          # detached HEAD
        ref = head.split(":", 1)[1].strip()
        loose = os.path.join(git, *ref.split("/"))
        if os.path.isfile(loose):
            with open(loose, encoding="utf-8") as f:
                return f.read().strip()[:7]
        packed = os.path.join(git, "packed-refs")    # refs may be packed
        if os.path.isfile(packed):
            with open(packed, encoding="utf-8") as f:
                for line in f:
                    if line.startswith(("#", "^")):
                        continue
                    parts = line.split()
                    if len(parts) == 2 and parts[1] == ref:
                        return parts[0][:7]
    except (OSError, IndexError, UnicodeDecodeError):
        pass
    return None


def _read_excel_environment(root):
    """Read the Excel environment recorded by Export_ExcelEnvironment, if present."""
    path = os.path.join(root, "benchmark", ENV_FILE)
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return {k: str(data[k]).strip() for k in
                ("excel_version", "excel_build", "office_bitness") if data.get(k)}
    except (OSError, ValueError):
        return {}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel-version", default="unrecorded")
    ap.add_argument("--excel-build", default="unrecorded")
    ap.add_argument("--office-bitness", default="unrecorded")
    ap.add_argument("--commit-sha", default="unrecorded")
    ap.add_argument("--notes", default="")
    a = ap.parse_args()

    root = repo_root()
    grid = os.path.join(root, "benchmark", "probability_accuracy_grid.csv")
    contracts = os.path.join(root, "benchmark", "accuracy_contracts.csv")
    now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    # Explicit flags win; otherwise read what can be read, never guess.
    commit_sha = a.commit_sha
    if commit_sha == "unrecorded":
        commit_sha = _read_commit_sha(root) or "unrecorded"

    environment = {"excel_version": a.excel_version, "excel_build": a.excel_build,
                   "office_bitness": a.office_bitness}
    recorded = _read_excel_environment(root)
    for key, value in recorded.items():
        if environment[key] == "unrecorded":
            environment[key] = value

    missing = [k for k, v in environment.items() if v == "unrecorded"]
    if missing:
        print("  warning: unrecorded provenance: " + ", ".join(missing)
              + " (run Export_ExcelEnvironment, or pass the flags)")

    manifest = build_manifest(
        root, grid, contracts, generated_utc=now, commit_sha=commit_sha,
        environment=environment, notes=a.notes)

    out = os.path.join(root, "benchmark", MANIFEST_NAME)
    with open(out, "w", encoding="utf-8", newline="\n") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)
        f.write("\n")
    n = len(manifest["source_binding"])
    print(f"wrote {out}: {n} .bas modules bound, schema {manifest['grid_schema_version']}")


if __name__ == "__main__":
    main()
