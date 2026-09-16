"""
Write a source-provenance manifest for freshly exported observations.

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

The independent holdout has a dedicated binding because its exporter and grid
are separate.  Use this only after a real-Excel holdout export:

    python write_manifest.py --holdout --from-fresh-export

WRITING REQUIRES --from-fresh-export.  A bare invocation refuses and exits
non-zero: a manifest asserts that the committed observations were produced by
the checked-out source, and only a real export makes that true.  Use --dry-run
to see what would be written without touching anything.
"""
import argparse
import datetime
import json
import os
import subprocess

from _manifest import (build_manifest, build_holdout_manifest, repo_root,
                       MANIFEST_NAME, HOLDOUT_MANIFEST_NAME)
from excel_certification import CertificationError, validate_record

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


RECORD_FILE = "benchmark/excel_regression_record.json"


def _git_show(root, spec):
    """Return the bytes of `spec` (e.g. HEAD:path) or None if unavailable."""
    try:
        r = subprocess.run(["git", "show", spec], cwd=root, capture_output=True,
                           check=False)
        return r.stdout if r.returncode == 0 else None
    except (OSError, ValueError):
        return None


def _grid_unchanged_since_head(root, grid_path):
    """(unchanged, reason). Unchanged means: no export can have produced this.

    A validated export record naming this grid as exported clears the check -
    the documented case of a real export yielding a byte-identical grid.
    Anything git cannot answer is treated as changed: this interlock must never
    block a legitimate export because git was unavailable. The provenance guard
    in CI is the backstop for that.
    """
    rel = os.path.relpath(grid_path, root).replace(os.sep, "/")
    # Ask git whether the path differs from HEAD rather than comparing bytes.
    # A byte comparison against `git show HEAD:<path>` is wrong wherever the
    # working tree and the blob differ only in line endings: this repository
    # declares *.csv as eol=lf, Excel writes CRLF, so on Windows the grid
    # always looked modified and this interlock never fired. `git diff` honours
    # .gitattributes and core.autocrlf, so it answers the question actually
    # being asked - did anything change - on every platform.
    try:
        r = subprocess.run(["git", "diff", "--quiet", "HEAD", "--", rel],
                           cwd=root, capture_output=True, check=False)
    except (OSError, ValueError):
        return False, ""
    if r.returncode == 1:
        return False, ""                     # differs from HEAD: a real export
    if r.returncode != 0:
        return False, ""                     # git unavailable or path unknown
    record = os.path.join(root, RECORD_FILE)
    if os.path.exists(record):
        try:
            with open(record, encoding="utf-8") as f:
                rec = json.load(f)
            validate_record(
                rec, root, policy_ref="HEAD", evidence_ref="HEAD",
                require_pass=True, require_grid=rel,
            )
            return False, ""
        except (OSError, ValueError, CertificationError):
            pass
    return True, f"{rel} is unchanged from HEAD"


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
    # INTERLOCK. Writing a manifest asserts "these observations were produced
    # by this source". Only a fresh export makes that true. 9fba175 carried a
    # bare invocation that rebound the main manifest to source three commits
    # newer than the observations, turning a truthful STALE EVIDENCE boundary
    # into a false clean binding - one holdout write away from a green gate on
    # stale evidence. A bare call now refuses.
    ap.add_argument("--from-fresh-export", action="store_true",
                    help="assert that observations were just re-exported from "
                         "the checked-out source in Excel; required to write")
    ap.add_argument("--dry-run", action="store_true",
                    help="print what would be written and exit 0 without "
                         "touching any manifest")
    ap.add_argument("--holdout", action="store_true",
                    help="bind benchmark/holdout/holdout_grid.csv instead of the main grid")
    a = ap.parse_args()

    root = repo_root()
    if a.holdout:
        grid = os.path.join(root, "benchmark", "holdout", "holdout_grid.csv")
    else:
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
    if a.holdout and (missing or commit_sha == "unrecorded"):
        missing_fields = list(missing)
        if commit_sha == "unrecorded":
            missing_fields.append("commit_sha")
        raise SystemExit("refusing to bind holdout with unrecorded provenance: "
                         + ", ".join(missing_fields))

    # INTERLOCK, evaluated before anything is written.
    if not a.from_fresh_export and not a.dry_run:
        which = "holdout" if a.holdout else "main"
        raise SystemExit(
            f"refusing to rewrite the {which} manifest.\n"
            "  Writing a manifest asserts that the committed observations were\n"
            "  produced by the checked-out source. Only a fresh Excel export\n"
            "  makes that true; rebinding without one silently converts stale\n"
            "  evidence into a false clean binding (see 9fba175).\n"
            "  If you have just re-exported, pass --from-fresh-export.\n"
            "  To preview without writing, pass --dry-run.")

    # SECOND INTERLOCK: verify the assertion instead of trusting it.
    #
    # --from-fresh-export is a promise, and a promise is cheaper to make than an
    # export. It was passed once without an export behind it, rebinding the main
    # manifest to a newer commit while the observations were unchanged - the same
    # false binding as 9fba175, reached through the flag meant to prevent it.
    #
    # A real export rewrites the grid. If the grid is byte-identical to HEAD,
    # nothing was exported and the assertion is provably false, so refuse. The
    # genuine exception - a .bas change that alters no observation value, giving
    # a byte-identical grid after a real export - is the same one
    # check_manifest_provenance.py already recognises: a validated
    # excel_regression_record.json naming this grid as exported.
    if a.from_fresh_export and not a.dry_run:
        stale, why = _grid_unchanged_since_head(root, grid)
        if stale:
            which = "holdout" if a.holdout else "main"
            raise SystemExit(
                f"refusing to rewrite the {which} manifest: {why}\n"
                "  --from-fresh-export asserts that the observations were just\n"
                "  re-exported from the checked-out source, but the grid is\n"
                "  unchanged from HEAD, so no export produced it.\n"
                "  Export from Excel first (Export_Accuracy_Observations), then\n"
                "  re-run this command.\n"
                "  If a real export legitimately reproduced a byte-identical\n"
                f"  grid, commit a validated {RECORD_FILE} naming this grid as\n"
                "  exported, as check_manifest_provenance.py requires.")

    builder = build_holdout_manifest if a.holdout else build_manifest
    manifest = builder(root, grid, contracts, generated_utc=now,
                       commit_sha=commit_sha, environment=environment,
                       notes=a.notes)

    if a.holdout:
        out = os.path.join(root, "benchmark", "holdout", HOLDOUT_MANIFEST_NAME)
    else:
        out = os.path.join(root, "benchmark", MANIFEST_NAME)
    if a.dry_run:
        n = len(manifest["source_binding"])
        rows = manifest.get("observation_row_count")
        print(f"dry run: would write {out}")
        print(f"  commit_sha {manifest['commit_sha']}, {n} .bas modules bound, "
              f"{rows} observation rows, schema {manifest['grid_schema_version']}")
        print("  nothing was written")
        return

    with open(out, "w", encoding="utf-8", newline="\n") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)
        f.write("\n")
    n = len(manifest["source_binding"])
    rows = manifest.get("observation_row_count")
    row_note = f", {rows} observation rows" if rows is not None else ""
    print(f"wrote {out}: {n} .bas modules bound{row_note}, "
          f"schema {manifest['grid_schema_version']}")


if __name__ == "__main__":
    main()
