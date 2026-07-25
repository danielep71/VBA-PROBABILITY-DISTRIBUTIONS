"""
Observation-provenance manifest: bind the committed accuracy evidence to the
exact source that produced it.

The observations in probability_accuracy_grid.csv are produced by VBA (the
library kernels in src/, plus the export/study macros) running inside Excel.
Nothing in the grid records WHICH source produced them, so an algorithm can
change in src/** while the committed observations - and therefore the green
accuracy summary - stay put. This module records a content hash of every .bas
file (algorithms, exporters, tests) plus the grid schema and environment, and
the gate refuses to certify a summary whose recorded source no longer matches
the checked-out source.

Hashes are taken over LF-normalized bytes so a CRLF working tree (the .bas
files are eol=crlf) and an LF checkout hash identically - the binding is to
content, not to line-ending presentation. No git or network is required.
"""
import glob
import hashlib
import json
import os

MANIFEST_NAME = "observation_manifest.json"
MANIFEST_VERSION = 1

# The observations depend on every piece of VBA that can change a returned
# value: the kernels (src), the exporters and study macros (benchmark), and the
# test harness (tests). Any .bas change must invalidate the binding until the
# observations are re-exported.
_BAS_GLOBS = ("src/**/*.bas", "tests/**/*.bas", "benchmark/**/*.bas")

# Grid columns the analyzer relies on. If these change, old observations may be
# misread, so the schema is part of the binding.
GRID_SCHEMA_VERSION = "arg4+expected_error/v1"


def repo_root(start=None):
    """Repo root = the directory that contains benchmark/."""
    here = os.path.dirname(os.path.abspath(start or __file__))
    return os.path.dirname(here)


def normalized_hash(path):
    """SHA-256 over LF-normalized file bytes (CRLF/CR/LF all hash the same)."""
    data = open(path, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")
    return "sha256:" + hashlib.sha256(data).hexdigest()


def _rel(root, path):
    return os.path.relpath(path, root).replace(os.sep, "/")


def source_hashes(root):
    """Map every tracked .bas path (repo-relative) to its normalized hash."""
    out = {}
    for pattern in _BAS_GLOBS:
        for p in glob.glob(os.path.join(root, pattern), recursive=True):
            out[_rel(root, p)] = normalized_hash(p)
    return dict(sorted(out.items()))


def grid_columns(grid_path):
    with open(grid_path, newline="") as f:
        return f.readline().strip().split(",")


def build_manifest(root, grid_path, contracts_path, generated_utc,
                   commit_sha="unrecorded", environment=None, notes=""):
    """Assemble a manifest dict for the current source and evidence."""
    return {
        "manifest_version": MANIFEST_VERSION,
        "generated_utc": generated_utc,
        "commit_sha": commit_sha,
        "grid_schema_version": GRID_SCHEMA_VERSION,
        "grid_schema_columns": grid_columns(grid_path),
        "source_binding": source_hashes(root),
        "contract_registry_sha256": normalized_hash(contracts_path),
        "environment": environment or {
            "excel_version": "unrecorded",
            "excel_build": "unrecorded",
            "office_bitness": "unrecorded",
        },
        "notes": notes,
    }


def load_manifest(root):
    path = os.path.join(root, "benchmark", MANIFEST_NAME)
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def verify_source_binding(root, manifest, grid_path):
    """
    Return a list of human-readable mismatches between the manifest and the
    checked-out source/grid. Empty list means the evidence is bound to this
    exact source. Enforces module hashes and grid schema; environment,
    commit_sha and contract sha are provenance metadata, not enforced here.
    """
    problems = []
    recorded = manifest.get("source_binding", {})
    current = source_hashes(root)

    for path, h in recorded.items():
        if path not in current:
            problems.append(f"{path}: recorded in manifest but missing from the tree")
        elif current[path] != h:
            problems.append(f"{path}: source changed since observations were generated")
    for path in current:
        if path not in recorded:
            problems.append(f"{path}: present in the tree but absent from the manifest "
                            "(source added without re-exporting observations)")

    if manifest.get("grid_schema_version") != GRID_SCHEMA_VERSION:
        problems.append(f"grid schema version {manifest.get('grid_schema_version')!r} "
                        f"!= current {GRID_SCHEMA_VERSION!r}")
    else:
        cols = grid_columns(grid_path)
        if manifest.get("grid_schema_columns") != cols:
            problems.append("grid schema columns changed since the manifest was written")

    return problems
