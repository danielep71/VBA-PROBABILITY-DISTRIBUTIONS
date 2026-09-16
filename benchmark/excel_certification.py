"""Canonical Excel-host certification record validation.

The record binds runtime evidence to one exact Git candidate SHA and exact VBA
source bytes. A green regression record does not, by itself, prove that either
accuracy grid was freshly exported. Grid claims are separate and require an
explicit exported=true entry with matching committed bytes and the exporter
module included in the certified source inventory.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import subprocess
from datetime import datetime

POLICY_PATH = ".github/excel-evidence-policy.json"
RETAINED_RECORD = "benchmark/excel_regression_record.json"
SHA40 = re.compile(r"^[0-9a-f]{40}$")
SHA256 = re.compile(r"^sha256:[0-9a-f]{64}$")
STAGES = ("import", "compile", "regression", "cleanup")
STAGE_STATUSES = {"PASS", "FAIL", "NOT_RUN", "TIMEOUT"}


class CertificationError(ValueError):
    pass


def _run_git(root, *args):
    return subprocess.run(["git"] + list(args), cwd=root, capture_output=True,
                          check=False)


def git_text(root, *args):
    proc = _run_git(root, *args)
    if proc.returncode:
        msg = proc.stderr.decode("utf-8", "replace").strip()
        raise CertificationError(msg or "git command failed")
    return proc.stdout.decode("utf-8", "strict")


def git_bytes(root, ref, path):
    proc = _run_git(root, "show", f"{ref}:{path}")
    if proc.returncode:
        raise CertificationError(f"{path} is not present at {ref}")
    return proc.stdout


def git_head(root):
    return git_text(root, "rev-parse", "HEAD").strip()


def sha256_bytes(raw):
    return "sha256:" + hashlib.sha256(raw).hexdigest()


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(65536), b""):
            h.update(block)
    return "sha256:" + h.hexdigest()


def row_count_bytes(raw):
    text = raw.decode("utf-8-sig")
    return max(len(text.splitlines()) - 1, 0)


def row_count_file(path):
    with open(path, "rb") as f:
        return row_count_bytes(f.read())


def _load_json_bytes(raw, label):
    try:
        value = json.loads(raw.decode("utf-8-sig"))
    except (UnicodeDecodeError, ValueError) as exc:
        raise CertificationError(f"{label} is not valid UTF-8 JSON: {exc}")
    if not isinstance(value, dict):
        raise CertificationError(f"{label} must be a JSON object")
    return value


def load_policy(root, ref="HEAD"):
    policy = _load_json_bytes(git_bytes(root, ref, POLICY_PATH), "Excel evidence policy")
    expected = {"schema_version", "repository", "workflow", "entry_point",
                "expected_assertions", "regression_sources", "grid_exporters"}
    if set(policy) != expected:
        missing = sorted(expected - set(policy))
        extra = sorted(set(policy) - expected)
        raise CertificationError(
            "Excel evidence policy keys differ from schema"
            + (f"; missing: {', '.join(missing)}" if missing else "")
            + (f"; extra: {', '.join(extra)}" if extra else ""))
    if policy["schema_version"] != 1:
        raise CertificationError("unsupported Excel evidence policy schema")
    if not isinstance(policy["repository"], str) or "/" not in policy["repository"]:
        raise CertificationError("policy repository must be owner/name")
    if not isinstance(policy["workflow"], str) or not policy["workflow"].startswith(".github/workflows/"):
        raise CertificationError("policy workflow must be a repository workflow path")
    if not isinstance(policy["entry_point"], str) or not policy["entry_point"].strip():
        raise CertificationError("policy entry_point must be non-empty")
    if type(policy["expected_assertions"]) is not int or policy["expected_assertions"] <= 0:
        raise CertificationError("policy expected_assertions must be a positive integer")
    sources = policy["regression_sources"]
    if not isinstance(sources, list) or not sources or not all(isinstance(p, str) and p for p in sources):
        raise CertificationError("policy regression_sources must be a non-empty string array")
    if len(sources) != len(set(sources)):
        raise CertificationError("policy regression_sources must be unique")
    exporters = policy["grid_exporters"]
    if not isinstance(exporters, dict) or not exporters:
        raise CertificationError("policy grid_exporters must be a non-empty object")
    for grid, exporter in exporters.items():
        if not isinstance(grid, str) or not grid.endswith(".csv") or not isinstance(exporter, str) or not exporter.endswith(".bas"):
            raise CertificationError("policy grid_exporters entries must map CSV paths to BAS paths")
    return policy


def _timestamp(value, field):
    if not isinstance(value, str) or not value:
        raise CertificationError(f"{field} must be a timestamp string")
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as exc:
        raise CertificationError(f"{field} is not ISO-8601: {exc}")
    if parsed.utcoffset() is None:
        raise CertificationError(f"{field} must include a timezone")
    return parsed


def _require_keys(obj, expected, label):
    if not isinstance(obj, dict):
        raise CertificationError(f"{label} must be an object")
    keys = set(obj)
    missing = sorted(expected - keys)
    extra = sorted(keys - expected)
    if missing or extra:
        raise CertificationError(
            f"{label} keys differ from schema"
            + (f"; missing: {', '.join(missing)}" if missing else "")
            + (f"; extra: {', '.join(extra)}" if extra else ""))


def _source_map(record):
    sources = record.get("sources")
    if not isinstance(sources, list) or not sources:
        raise CertificationError("sources must be a non-empty array")
    result = {}
    previous = None
    for item in sources:
        _require_keys(item, {"path", "sha256"}, "source entry")
        path = item["path"]
        digest = item["sha256"]
        if not isinstance(path, str) or not path or path.startswith(("/", "\\")) or ".." in path.split("/"):
            raise CertificationError("source path must be repository-relative")
        if not isinstance(digest, str) or not SHA256.fullmatch(digest):
            raise CertificationError(f"source digest for {path!r} is not canonical SHA-256")
        if path in result:
            raise CertificationError(f"duplicate source entry: {path}")
        if previous is not None and path <= previous:
            raise CertificationError("sources must be sorted by path")
        previous = path
        result[path] = digest
    return result


def _required_sources(policy, record):
    required = set(policy["regression_sources"])
    grids = record.get("grids")
    if isinstance(grids, dict):
        for grid, value in grids.items():
            if isinstance(value, dict) and value.get("exported") is True:
                exporter = policy["grid_exporters"].get(grid)
                if exporter is None:
                    raise CertificationError(f"record claims unsupported exported grid: {grid}")
                required.add(exporter)
    return sorted(required)


def _validate_runner(record, policy):
    runner = record["runner"]
    _require_keys(runner, {"class", "identity", "workflow"}, "runner")
    if runner["class"] not in ("self-hosted-excel", "manual-interactive"):
        raise CertificationError("runner class is not supported")
    if not isinstance(runner["identity"], str) or not runner["identity"].strip():
        raise CertificationError("runner identity is required")
    workflow = runner["workflow"]
    if runner["class"] == "self-hosted-excel":
        _require_keys(workflow, {"repository", "path", "sha", "run_id", "run_attempt"}, "runner workflow")
        if workflow["repository"] != policy["repository"] or workflow["path"] != policy["workflow"]:
            raise CertificationError("runner workflow identity differs from policy")
        if not isinstance(workflow["sha"], str) or not SHA40.fullmatch(workflow["sha"]):
            raise CertificationError("runner workflow sha must be a full commit SHA")
        if type(workflow["run_id"]) is not int or workflow["run_id"] <= 0 or type(workflow["run_attempt"]) is not int or workflow["run_attempt"] <= 0:
            raise CertificationError("runner workflow run_id/run_attempt must be positive integers")
    elif workflow is not None:
        raise CertificationError("manual-interactive evidence must not claim a hosted workflow")


def _validate_environment(environment, require_runtime):
    expected = {"excel_version", "excel_build", "office_bitness", "windows",
                "os_architecture", "automation_security", "vba_project_access"}
    _require_keys(environment, expected, "environment")
    if require_runtime:
        for key in expected:
            if not isinstance(environment[key], str) or not environment[key].strip():
                raise CertificationError(f"environment.{key} must be non-empty for executed evidence")
        if environment["office_bitness"] not in ("32-bit", "64-bit"):
            raise CertificationError("environment.office_bitness must be 32-bit or 64-bit")


def _validate_stages(record, require_pass):
    stages = record["stages"]
    _require_keys(stages, set(STAGES), "stages")
    for name in STAGES:
        stage = stages[name]
        _require_keys(stage, {"status", "detail"}, f"stage {name}")
        if stage["status"] not in STAGE_STATUSES:
            raise CertificationError(f"invalid {name} stage status")
        if not isinstance(stage["detail"], str) or not stage["detail"].strip():
            raise CertificationError(f"stage {name} requires detail")
    if stages["compile"]["status"] == "PASS" and stages["import"]["status"] != "PASS":
        raise CertificationError("compile PASS requires import PASS")
    if stages["regression"]["status"] == "PASS" and stages["compile"]["status"] != "PASS":
        raise CertificationError("regression PASS requires compile PASS")
    if stages["import"]["status"] != "PASS" and stages["compile"]["status"] not in ("NOT_RUN", "FAIL"):
        raise CertificationError("compile cannot pass/run after failed import")
    if stages["compile"]["status"] != "PASS" and stages["regression"]["status"] not in ("NOT_RUN", "FAIL"):
        raise CertificationError("regression cannot pass/run after failed compile")
    if require_pass and any(stages[name]["status"] != "PASS" for name in STAGES):
        bad = ", ".join(f"{name}={stages[name]['status']}" for name in STAGES if stages[name]["status"] != "PASS")
        raise CertificationError("green certification required; " + bad)


def _validate_harness(record, policy):
    harness = record["harness"]
    _require_keys(harness, {"entry_point", "assertions", "passed", "failed"}, "harness")
    if harness["entry_point"] != policy["entry_point"]:
        raise CertificationError("harness entry point differs from policy")
    for field in ("assertions", "passed", "failed"):
        if type(harness[field]) is not int or harness[field] < 0:
            raise CertificationError(f"harness.{field} must be a non-negative integer")
    if harness["passed"] + harness["failed"] != harness["assertions"]:
        raise CertificationError("harness counters are inconsistent")
    if record["stages"]["regression"]["status"] == "PASS":
        if harness["assertions"] != policy["expected_assertions"]:
            raise CertificationError(
                f"regression assertion count {harness['assertions']} differs from policy {policy['expected_assertions']}")
        if harness["failed"] != 0 or harness["passed"] != harness["assertions"]:
            raise CertificationError("regression PASS contradicts harness counters")


def _validate_log(record, log_directory):
    log = record["log"]
    _require_keys(log, {"path", "sha256"}, "log")
    if not isinstance(log["path"], str) or not log["path"] or os.path.isabs(log["path"]) or ".." in log["path"].replace("\\", "/").split("/"):
        raise CertificationError("log path must be relative")
    if not isinstance(log["sha256"], str) or not SHA256.fullmatch(log["sha256"]):
        raise CertificationError("log sha256 is not canonical")
    if log_directory is not None:
        path = os.path.join(log_directory, log["path"])
        if not os.path.isfile(path):
            raise CertificationError(f"retained log is missing: {path}")
        if sha256_file(path) != log["sha256"]:
            raise CertificationError("retained log digest does not match record")


def _validate_grids(record, policy, root, evidence_ref, require_grid):
    grids = record["grids"]
    if not isinstance(grids, dict):
        raise CertificationError("grids must be an object")
    if set(grids) != set(policy["grid_exporters"]):
        raise CertificationError("grids keys must exactly match policy grid targets")
    for grid, entry in grids.items():
        _require_keys(entry, {"exported", "sha256", "row_count", "exported_utc"}, f"grid {grid}")
        if type(entry["exported"]) is not bool:
            raise CertificationError(f"grid {grid} exported must be boolean")
        if not entry["exported"]:
            if any(entry[k] is not None for k in ("sha256", "row_count", "exported_utc")):
                raise CertificationError(f"grid {grid} not exported but carries export claims")
            continue
        if not isinstance(entry["sha256"], str) or not SHA256.fullmatch(entry["sha256"]):
            raise CertificationError(f"grid {grid} has invalid sha256")
        if type(entry["row_count"]) is not int or entry["row_count"] < 0:
            raise CertificationError(f"grid {grid} has invalid row_count")
        _timestamp(entry["exported_utc"], f"grid {grid} exported_utc")
        if evidence_ref is not None:
            raw = git_bytes(root, evidence_ref, grid)
            if sha256_bytes(raw) != entry["sha256"]:
                raise CertificationError(f"grid {grid} digest differs from {evidence_ref}")
            if row_count_bytes(raw) != entry["row_count"]:
                raise CertificationError(f"grid {grid} row count differs from {evidence_ref}")
    if require_grid is not None:
        if require_grid not in grids or grids[require_grid]["exported"] is not True:
            raise CertificationError(f"record does not claim a fresh export of {require_grid}")


def validate_record(record, root, *, policy_ref="HEAD", evidence_ref=None,
                    expected_candidate_sha=None, require_pass=False,
                    require_grid=None, log_directory=None):
    policy = load_policy(root, policy_ref)
    expected = {"schema_version", "repository", "candidate_sha", "execution",
                "started_at", "finished_at", "runner", "environment", "sources",
                "stages", "harness", "log", "grids"}
    _require_keys(record, expected, "Excel certification record")
    if record["schema_version"] != 1:
        raise CertificationError("unsupported Excel certification schema")
    if record["repository"] != policy["repository"]:
        raise CertificationError("record repository differs from policy")
    candidate = record["candidate_sha"]
    if not isinstance(candidate, str) or not SHA40.fullmatch(candidate):
        raise CertificationError("candidate_sha must be a full lowercase Git SHA")
    if expected_candidate_sha is not None and candidate != expected_candidate_sha:
        raise CertificationError(
            f"record candidate {candidate} differs from required candidate {expected_candidate_sha}")
    git_text(root, "cat-file", "-e", candidate + "^{commit}")
    if record["execution"] not in ("automated", "manual"):
        raise CertificationError("execution must be automated or manual")
    started = _timestamp(record["started_at"], "started_at")
    finished = _timestamp(record["finished_at"], "finished_at")
    if finished < started:
        raise CertificationError("finished_at precedes started_at")
    _validate_runner(record, policy)
    _validate_environment(record["environment"], require_runtime=True)
    _validate_stages(record, require_pass)
    _validate_harness(record, policy)
    _validate_log(record, log_directory)
    _validate_grids(record, policy, root, evidence_ref, require_grid)

    source_map = _source_map(record)
    required = _required_sources(policy, record)
    if sorted(source_map) != required:
        missing = sorted(set(required) - set(source_map))
        extra = sorted(set(source_map) - set(required))
        raise CertificationError(
            "certified source inventory differs from policy"
            + (f"; missing: {', '.join(missing)}" if missing else "")
            + (f"; extra: {', '.join(extra)}" if extra else ""))
    for path, digest in source_map.items():
        candidate_raw = git_bytes(root, candidate, path)
        if sha256_bytes(candidate_raw) != digest:
            raise CertificationError(f"source digest for {path} differs from candidate {candidate}")
        if evidence_ref is not None:
            evidence_raw = git_bytes(root, evidence_ref, path)
            if sha256_bytes(evidence_raw) != digest:
                raise CertificationError(
                    f"source {path} at evidence commit {evidence_ref} differs from certified candidate {candidate}")
    return policy


def validate_record_text(text, root, **kwargs):
    try:
        record = json.loads(text)
    except ValueError as exc:
        raise CertificationError(f"Excel certification record is not valid JSON: {exc}")
    if not isinstance(record, dict):
        raise CertificationError("Excel certification record must be a JSON object")
    validate_record(record, root, **kwargs)
    return record
