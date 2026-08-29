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
HOLDOUT_MANIFEST_NAME = "holdout_manifest.json"
MANIFEST_VERSION = 1

# The observations depend on every piece of VBA that can change a returned
# value: the kernels (src), the exporters and study macros (benchmark), and the
# test harness (tests). Any .bas change must invalidate the binding until the
# observations are re-exported.
_BAS_GLOBS = ("src/**/*.bas", "tests/**/*.bas", "benchmark/**/*.bas")

# The independent holdout is produced by the six public source modules and its
# dedicated exporter.  Study macros and the regression harness do not produce
# holdout observations, so binding them would invalidate this evidence for a
# change that cannot affect it.
_HOLDOUT_BAS_GLOBS = (
    "src/**/*.bas",
    # Recursive, not a literal exporter path. A literal path binds only the file
    # named here, so a SECOND module added beside it would produce holdout
    # observations from source no manifest records - the exact fail-open class
    # this binding exists to prevent. src/**/*.bas already detects additions;
    # this makes the holdout directory symmetric with it.
    "benchmark/holdout/**/*.bas",
)

# Grid columns the analyzer relies on. If these change, old observations may be
# misread, so the schema is part of the binding.
GRID_SCHEMA_VERSION = "arg4+expected_error/v1"
HOLDOUT_GRID_SCHEMA_VERSION = "holdout/arg4+expected_error/v1"


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


def source_hashes(root, patterns=_BAS_GLOBS):
    """Map every tracked .bas path (repo-relative) to its normalized hash."""
    out = {}
    for pattern in patterns:
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
        "observation_grid_sha256": normalized_hash(grid_path),
        "source_binding": source_hashes(root),
        "contract_registry_sha256": normalized_hash(contracts_path),
        "environment": environment or {
            "excel_version": "unrecorded",
            "excel_build": "unrecorded",
            "office_bitness": "unrecorded",
        },
        "notes": notes,
    }


def build_holdout_manifest(root, grid_path, contracts_path, generated_utc,
                           commit_sha="unrecorded", environment=None, notes=""):
    """Assemble the independent-holdout provenance record."""
    with open(grid_path, newline="") as f:
        row_count = max(sum(1 for _ in f) - 1, 0)
    return {
        "manifest_version": MANIFEST_VERSION,
        "evidence_kind": "independent_holdout",
        "generated_utc": generated_utc,
        "commit_sha": commit_sha,
        "grid_schema_version": HOLDOUT_GRID_SCHEMA_VERSION,
        "grid_schema_columns": grid_columns(grid_path),
        "observation_row_count": row_count,
        "observation_grid_sha256": normalized_hash(grid_path),
        "source_binding": source_hashes(root, _HOLDOUT_BAS_GLOBS),
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


def load_holdout_manifest(root):
    """Load the dedicated holdout binding, or return None while it is unbound."""
    path = os.path.join(root, "benchmark", "holdout", HOLDOUT_MANIFEST_NAME)
    if not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError) as exc:
        return {"_load_error": f"{type(exc).__name__}: {exc}"}


def verify_source_binding(root, manifest, grid_path, contracts_path):
    """
    Return a list of human-readable mismatches between the manifest and the
    checked-out source, grid, and contract registry. Empty list means these
    exact observation bytes are bound to this exact source. Enforces:
      * every .bas module hash (source structure and behaviour),
      * the grid schema version and columns,
      * the full observation-grid content hash (the observation bytes),
      * the contract-registry content hash (the thresholds that produced the verdicts).
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

    # Observation bytes: an edited observed_vba (or any grid content) must invalidate
    # the binding even when the columns are unchanged.
    grid_recorded = manifest.get("observation_grid_sha256")
    if grid_recorded is None:
        problems.append("manifest predates observation-content binding "
                        "(no observation_grid_sha256); regenerate with write_manifest.py")
    elif normalized_hash(grid_path) != grid_recorded:
        problems.append("observation grid contents changed since the manifest was written "
                        "(an observation was edited or re-exported without re-binding)")

    # Contract registry: the verdicts are only valid for the thresholds they were
    # produced against, so the registry is bound too.
    reg_recorded = manifest.get("contract_registry_sha256")
    if reg_recorded is None:
        problems.append("manifest has no contract_registry_sha256; regenerate with write_manifest.py")
    elif normalized_hash(contracts_path) != reg_recorded:
        problems.append("contract registry changed since the manifest was written "
                        "(thresholds/contracts edited without re-binding)")

    return problems


def verify_holdout_binding(root, manifest, grid_path, contracts_path):
    """
    Verify that independent-holdout observations came from the checked-out
    production modules and holdout exporter under the recorded schema and
    contract registry.  Missing provenance is a mismatch, never a legacy PASS.
    """
    if manifest is None:
        return [f"no {HOLDOUT_MANIFEST_NAME}: independent holdout observations are unbound"]
    if manifest.get("_load_error"):
        return [f"malformed {HOLDOUT_MANIFEST_NAME}: {manifest['_load_error']}"]

    problems = []
    if manifest.get("manifest_version") != MANIFEST_VERSION:
        problems.append(f"holdout manifest version {manifest.get('manifest_version')!r} "
                        f"!= current {MANIFEST_VERSION!r}")
    if manifest.get("evidence_kind") != "independent_holdout":
        problems.append("manifest evidence_kind is not 'independent_holdout'")

    recorded = manifest.get("source_binding", {})
    current = source_hashes(root, _HOLDOUT_BAS_GLOBS)
    for path, h in recorded.items():
        if path not in current:
            problems.append(f"{path}: recorded in holdout manifest but missing from the tree")
        elif current[path] != h:
            problems.append(f"{path}: source changed since holdout observations were generated")
    for path in current:
        if path not in recorded:
            problems.append(f"{path}: present in the tree but absent from the holdout manifest "
                            "(source added without re-exporting holdout observations)")

    if manifest.get("grid_schema_version") != HOLDOUT_GRID_SCHEMA_VERSION:
        problems.append(f"holdout grid schema version {manifest.get('grid_schema_version')!r} "
                        f"!= current {HOLDOUT_GRID_SCHEMA_VERSION!r}")
    elif manifest.get("grid_schema_columns") != grid_columns(grid_path):
        problems.append("holdout grid schema columns changed since the manifest was written")

    with open(grid_path, newline="") as f:
        row_count = max(sum(1 for _ in f) - 1, 0)
    if manifest.get("observation_row_count") != row_count:
        problems.append("holdout observation row count changed since the manifest was written")

    grid_recorded = manifest.get("observation_grid_sha256")
    if grid_recorded is None:
        problems.append("holdout manifest has no observation_grid_sha256; re-export and re-bind")
    elif normalized_hash(grid_path) != grid_recorded:
        problems.append("holdout observation grid contents changed since the manifest was written "
                        "(an observation was edited or re-exported without re-binding)")

    reg_recorded = manifest.get("contract_registry_sha256")
    if reg_recorded is None:
        problems.append("holdout manifest has no contract_registry_sha256; re-export and re-bind")
    elif normalized_hash(contracts_path) != reg_recorded:
        problems.append("contract registry changed since holdout observations were generated")

    return problems
