"""Apply an Excel grid export made unattended on the self-hosted runner (#181).

The Excel VBA Regression workflow, dispatched with export inputs, runs the
regression and then ci/Export-ExcelGrids.ps1 in the same job. It uploads the
grids exactly as Excel wrote them, the Excel environment file completed with
the commit SHA, the green regression record and log, and export-manifest.json,
as the artifact excel-grid-export-<run id>. This tool turns that artifact into
the evidence change a person commits. It never commits, and the runner never
writes to the repository.

USAGE
    gh run download <run id> --name excel-grid-export-<run id> --dir <dir>
    python benchmark/apply_excel_export.py --artifact <dir> [--grid main ...]

Run it from the checkout of the exported commit, with a clean working tree:
the evidence commit must be the export's direct child for the records to say
what Excel ran.

WHAT IT REFUSES
    Nothing is written unless every check below passes:

    * the export manifest reports PASS, names the full SHA of HEAD, and lists
      exactly the grids passed with --grid (when --grid is given);
    * no tracked file differs from HEAD;
    * the regression record validates as green for HEAD, with its retained
      log, through excel_certification.validate_record;
    * the environment file names HEAD, was completed by the adapter, and agrees
      with the bytes Excel wrote and with the regression record;
    * every requested grid is in the artifact with the raw SHA-256 the manifest
      records, and, compared with the grid at HEAD, it changes only the
      observed_vba cells its exporter owns, fills every one of them, and holds
      exactly the bytes VBA's Print # writes (CRLF lines). A study probe CSV,
      which its probe writes whole (#225), must instead keep the committed
      header, line count and each row's key fields. The manifest's counts must
      agree with this recount;
    * the imported modules the manifest lists hash to their blobs at HEAD;
    * a missing-only main export keeps the observations already in the grid,
      so it is accepted only while no .bas module differs from the source
      binding of the committed observation_manifest.json. After any VBA
      change only a full export can make those observations current.

WHAT IT DOES
    1. Copies each grid and excel_environment.json into place, byte for byte.
    2. Runs finalize_excel_certification.py --from-fresh-export with --main
       and/or --holdout, which records raw_sha256 from those bytes.
       holdout_older, survival_boundary, tchi_large_df and the study grids
       have no exporter in the Excel evidence policy, so the record makes no
       claim for them.
    3. Rewrites each grid with the LF bytes Git stores under `*.csv text
       eol=lf`, as a fresh checkout of the evidence commit would hold them.
    4. Runs refresh_evidence.py with --bind-exported-main and/or
       --bind-exported-holdout. Before the commit exists, its "retained Excel
       record" step compares the record with HEAD and must fail when a grid's
       bytes changed; that one failure is expected, and step 5 replaces it.
       Any other failure is an open item: the tool reports it as NOT READY and
       leaves the files in place. An export that satisfies a release gate is
       the usual case: release_readiness.json must then resolve that gate's
       blocker by hand.
    5. Validates the finalized record with check_excel_certification.py
       against a snapshot commit of the working tree (git stash create, which
       changes no ref and no stash entry), so the exact bytes to be committed
       are checked.
    6. Prints the files to commit, the commit message body and the commands
       that validate the evidence commit.

Exit code 0 when the evidence is ready to commit; 1 on a refusal, a failed
step, or an open item.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from dataclasses import dataclass

from excel_certification import (CertificationError, canonical_grid_bytes, git_bytes,
                                 git_text, sha256_bytes, validate_record)
import _manifest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

ADAPTER = "ci/Export-ExcelGrids.ps1"
MANIFEST = "export-manifest.json"
RECORD = "excel-certification.json"
LOG = "test-result.txt"
ENVIRONMENT = "excel_environment.json"
EXCEL_WRITTEN_ENVIRONMENT = "excel_environment.excel.json"
ENV_PATH = "benchmark/excel_environment.json"
RETAINED_RECORD = "benchmark/excel_regression_record.json"
SHA40 = re.compile(r"^[0-9a-f]{40}$")
# refresh_evidence.py's VERIFY label for the retained-record check, which
# compares the record with HEAD and so fails before the evidence commit exists
# whenever an exported grid's bytes changed.
PRE_COMMIT_STEP = "retained Excel record"
README_STEP = "root README assurance block"
MODES = ("full", "missing-only")


@dataclass(frozen=True)
class GridSpec:
    name: str
    path: str
    observed: int            # index of observed_vba
    min_columns: int         # the exporter skips rows with fewer fields
    scope: tuple[int, bytes] | None   # (column, value) a row must carry
    macros: dict[str, str]   # mode -> exporter entry point
    exporter: str            # the module that holds the entry points
    finalize_flag: str | None
    bind_flag: str | None
    kind: str = "observed"   # "generated": the probe writes the whole CSV (#225)
    key_columns: int = 0     # generated: leading fields that identify a row


def _study(name: str, file_name: str, module: str, macro: str, columns: int,
           key_columns: int, folder: str = "positive_ratio_subnormal") -> GridSpec:
    folder = f"benchmark/{folder}/"
    return GridSpec(name, folder + file_name, -1, columns, None, {"full": macro},
                    folder + module, None, None, "generated", key_columns)


# Mirrors the exporter modules and ci/ExcelGridExport.psm1.
GRIDS = {
    "main": GridSpec("main", "benchmark/probability_accuracy_grid.csv", 9, 12,
                     (11, b"main grid"),
                     {"full": "Export_Accuracy_Observations",
                      "missing-only": "Export_Accuracy_MissingOnly"},
                     "benchmark/M_STATS_PROBDIST_ACCURACYEXPORT.bas",
                     "--main", "--bind-exported-main"),
    "holdout": GridSpec("holdout", "benchmark/holdout/holdout_grid.csv", 9, 12, None,
                        {"full": "Export_Holdout"},
                        "benchmark/holdout/M_STATS_PROBDIST_HOLDOUT.bas",
                        "--holdout", "--bind-exported-holdout"),
    "holdout_older": GridSpec("holdout_older",
                              "benchmark/holdout_older/holdout_older_grid.csv", 8, 11,
                              None, {"full": "Export_HoldoutOlder"},
                              "benchmark/holdout_older/M_STATS_PROBDIST_HOLDOUTOLDER.bas",
                              None, None),
    "survival_boundary": GridSpec("survival_boundary",
                                  "benchmark/survival_boundary/survival_boundary_grid.csv", 8, 11,
                                  None, {"full": "Export_SurvivalBoundary"},
                                  "benchmark/survival_boundary/M_STATS_PROBDIST_SURVBOUND.bas",
                                  None, None),
    "tchi_large_df": GridSpec("tchi_large_df",
                              "benchmark/tchi_large_df/tchi_large_df_grid.csv", 8, 9,
                              None, {"full": "Export_TChiLargeDf"},
                              "benchmark/tchi_large_df/M_STATS_PROBDIST_TCHI.bas",
                              None, None),
    # Study grids, selected through the workflow's export_studies input.
    "prs_gamma_probe": _study("prs_gamma_probe", "gamma_probe.csv",
                              "M_STATS_PROBDIST_GAMMAPROBE.bas",
                              "Export_GammaPositiveRatio", 15, 9),
    "prs_chisquare_probe": _study("prs_chisquare_probe", "chisquare_probe.csv",
                                  "M_STATS_PROBDIST_CHISQPROBE.bas",
                                  "Export_ChiSquarePositiveRatio", 14, 8),
    "prs_gamma_holdout": _study("prs_gamma_holdout", "gamma_holdout.csv",
                                "M_STATS_PROBDIST_GAMMAHOLD.bas", "Export_GammaHoldout", 15, 9),
    "prs_chisquare_holdout": _study("prs_chisquare_holdout", "chisquare_holdout.csv",
                                    "M_STATS_PROBDIST_CHISQHOLD.bas",
                                    "Export_ChiSquareHoldout", 14, 8),
    "prs_gamma_q_complement": _study("prs_gamma_q_complement", "gamma_q_complement.csv",
                                     "M_STATS_PROBDIST_GQPROBE.bas",
                                     "Export_GammaQComplement", 12, 3),
    "gss_small_shape": _study("gss_small_shape", "gamma_series_small_shape_grid.csv",
                              "M_STATS_PROBDIST_GSSS.bas", "Export_GammaSeriesSmallShape",
                              13, 10, folder="gamma_series_small_shape"),
    "glq_lower_q": _study("glq_lower_q", "gamma_lower_q_grid.csv",
                          "M_STATS_PROBDIST_GLQ.bas", "Export_GammaLowerQ",
                          13, 10, folder="gamma_lower_q"),
    "gbr_binary64_map": _study("gbr_binary64_map", "binary64_map_grid.csv",
                               "M_STATS_PROBDIST_GBR.bas", "Export_GammaBinary64Map",
                               13, 5, folder="gamma_route_study"),
}


class ApplyError(Exception):
    pass


# --------------------------------------------------------------------------
# Grid check (the same rules as Test-ExportedGrid in ci/ExcelGridExport.psm1)
# --------------------------------------------------------------------------
def grid_lines(raw: bytes) -> list[bytes]:
    """The lines an exporter works on: CRLF and CR become LF, the text is split
    on LF, and a final empty element is dropped because it is never Printed."""
    lines = raw.replace(b"\r\n", b"\n").replace(b"\r", b"\n").split(b"\n")
    if lines and lines[-1] == b"":
        lines.pop()
    return lines


def printed(lines: list[bytes]) -> bytes:
    """The bytes VBA's Print # writes for these lines."""
    return b"".join(line + b"\r\n" for line in lines)


def row_in_scope(spec: GridSpec, mode: str, line: bytes) -> bool:
    if not line.strip(b" "):
        return False
    cols = line.split(b",")
    if len(cols) < spec.min_columns:
        return False
    if spec.scope is not None and cols[spec.scope[0]].strip(b" ") != spec.scope[1]:
        return False
    if mode == "missing-only" and cols[spec.observed].strip(b" "):
        return False
    return True


def verify_generated(spec: GridSpec, baseline: bytes, exported: bytes,
                     max_listed: int = 20) -> dict[str, object]:
    """A probe CSV written whole, against the committed file it replaces (the
    same rules as Test-GeneratedGrid): Print # bytes, the same header and line
    count, every row with the header's field count, and each row's first
    key_columns fields unchanged. Only the measurement after them may differ."""
    problems: list[str] = []
    before, after = grid_lines(baseline), grid_lines(exported)
    counts = {"rows": max(len(after) - 1, 0), "in_scope_rows": max(len(before) - 1, 0),
              "filled_rows": 0, "error_rows": 0, "out_of_scope_rows": 0,
              "changed_observations": 0}
    if exported != printed(after):
        problems.append("the export is not CRLF-terminated lines as VBA Print # writes them")
    if not before:
        problems.append("the grid at HEAD is empty")
    if len(before) != len(after):
        problems.append(f"line count changed from {len(before)} to {len(after)}")
        return {**counts, "problems": problems}
    if not before:
        return {**counts, "problems": problems}
    if before[0] != after[0]:
        problems.append("the header line changed")
    if len(before[0].split(b",")) != spec.min_columns:
        problems.append(f"the committed header does not have {spec.min_columns} fields")
    rows: list[str] = []
    for index in range(1, len(before)):
        old_cols, new_cols, number = before[index].split(b","), after[index].split(b","), index + 1
        if len(new_cols) != spec.min_columns:
            rows.append(f"line {number}: {len(new_cols)} fields, not {spec.min_columns}")
            continue
        if old_cols[:spec.key_columns] != new_cols[:spec.key_columns]:
            rows.append(f"line {number}: the point's key fields changed")
            continue
        counts["filled_rows"] += 1
        measured = new_cols[spec.key_columns:]
        if b"ERROR" in measured or b"CVERR" in measured:
            counts["error_rows"] += 1
        if before[index] != after[index]:
            counts["changed_observations"] += 1
    problems.extend(rows[:max_listed])
    if len(rows) > max_listed:
        problems.append(f"... and {len(rows) - max_listed} more row problem(s)")
    if len(before) < 2:
        problems.append("no row is in the exporter's scope")
    return {**counts, "problems": problems}


def verify_grid(spec: GridSpec, mode: str, baseline: bytes, exported: bytes,
                max_listed: int = 20) -> dict[str, object]:
    """Counts and problems of an exported grid against the grid it came from."""
    if spec.kind == "generated":
        return verify_generated(spec, baseline, exported, max_listed)
    problems: list[str] = []
    before, after = grid_lines(baseline), grid_lines(exported)
    counts = {"rows": max(len(after) - 1, 0), "in_scope_rows": 0, "filled_rows": 0,
              "error_rows": 0, "out_of_scope_rows": 0, "changed_observations": 0}
    if exported != printed(after):
        problems.append("the export is not CRLF-terminated lines as VBA Print # writes them")
    if not before:
        problems.append("the grid at HEAD is empty")
    if len(before) != len(after):
        problems.append(f"line count changed from {len(before)} to {len(after)}")
        return {**counts, "problems": problems}
    if before and before[0] != after[0]:
        problems.append("the header line changed")
    rows: list[str] = []
    for index in range(1, len(before)):
        old, new, number = before[index], after[index], index + 1
        if not row_in_scope(spec, mode, old):
            counts["out_of_scope_rows"] += 1
            if old != new:
                rows.append(f"line {number}: out-of-scope row changed")
            continue
        counts["in_scope_rows"] += 1
        old_cols, new_cols = old.split(b","), new.split(b",")
        if len(old_cols) != len(new_cols):
            rows.append(f"line {number}: field count changed")
            continue
        for column, (a, b) in enumerate(zip(old_cols, new_cols)):
            if column != spec.observed and a != b:
                rows.append(f"line {number}: column {column} changed outside observed_vba")
        observed = new_cols[spec.observed].strip(b" ")
        if not observed:
            rows.append(f"line {number}: observed_vba left blank")
        else:
            counts["filled_rows"] += 1
            if observed == b"ERROR":
                counts["error_rows"] += 1
        if old_cols[spec.observed] != new_cols[spec.observed]:
            counts["changed_observations"] += 1
    problems.extend(rows[:max_listed])
    if len(rows) > max_listed:
        problems.append(f"... and {len(rows) - max_listed} more row problem(s)")
    if counts["in_scope_rows"] == 0:
        problems.append("missing-only found no blank observed_vba cell in a main-grid row"
                        if mode == "missing-only" else "no row is in the exporter's scope")
    return {**counts, "problems": problems}


# --------------------------------------------------------------------------
# Artifact checks
# --------------------------------------------------------------------------
def git(root: str, *args: str) -> str:
    try:
        return git_text(root, *args)
    except CertificationError as exc:
        raise ApplyError(f"git {' '.join(args)}: {exc}") from exc


def read_bytes(path: str, what: str) -> bytes:
    if not os.path.isfile(path):
        raise ApplyError(f"{what} is missing from the artifact: {path}")
    with open(path, "rb") as f:
        return f.read()


def read_json(path: str, what: str) -> dict[str, object]:
    raw = read_bytes(path, what)
    try:
        value = json.loads(raw.decode("utf-8-sig"))
    except (UnicodeDecodeError, ValueError) as exc:
        raise ApplyError(f"{what} is not valid JSON: {exc}") from exc
    if not isinstance(value, dict):
        raise ApplyError(f"{what} must be a JSON object")
    return value


def normalize_bitness(value: object) -> str:
    text = str(value).strip().lower().replace(" ", "")
    if text in ("64", "64bit", "64-bit", "x64"):
        return "64-bit"
    if text in ("32", "32bit", "32-bit", "x86"):
        return "32-bit"
    return str(value).strip()


@dataclass
class Plan:
    head: str
    artifact: str
    manifest: dict[str, object]
    record: dict[str, object]
    grids: list[tuple[GridSpec, str, bytes, dict[str, object]]]  # spec, mode, raw, counts
    environment: bytes


def check_artifact(root: str, artifact: str, expected: list[str] | None) -> Plan:
    """Every refusal, before anything is written."""
    if not os.path.isdir(artifact):
        raise ApplyError(f"artifact directory not found: {artifact}")
    manifest = read_json(os.path.join(artifact, MANIFEST), MANIFEST)
    if manifest.get("schema_version") != 1 or manifest.get("kind") != "excel-grid-export":
        raise ApplyError(f"{MANIFEST} is not a schema-1 excel-grid-export manifest")
    if manifest.get("status") != "PASS":
        raise ApplyError(f"the export did not pass on the runner: {manifest.get('failure')}")

    head = git(root, "rev-parse", "HEAD").strip()
    candidate = manifest.get("candidate_sha")
    if not isinstance(candidate, str) or not SHA40.fullmatch(candidate):
        raise ApplyError(f"{MANIFEST} candidate_sha is not a full SHA: {candidate!r}")
    if candidate != head:
        raise ApplyError(f"the artifact was exported at {candidate}, but HEAD is {head}; "
                         f"check out {candidate} first")

    requested = manifest.get("requested")
    if (not isinstance(requested, list) or not requested
            or any(name not in GRIDS for name in requested)
            or len(set(requested)) != len(requested)):
        raise ApplyError(f"{MANIFEST} requested must list known grids once each: {requested!r}")
    if expected is not None and sorted(expected) != sorted(requested):
        raise ApplyError(f"the artifact exported {', '.join(requested)}, "
                         f"not {', '.join(expected)}")

    dirty = git(root, "status", "--porcelain", "--untracked-files=no").strip()
    if dirty:
        raise ApplyError("tracked files differ from HEAD; apply to a clean checkout:\n" + dirty)

    record = read_json(os.path.join(artifact, RECORD), RECORD)
    try:
        validate_record(record, root, policy_ref=head, evidence_ref=head,
                        expected_candidate_sha=head, require_pass=True,
                        log_directory=artifact)
    except (CertificationError, KeyError, TypeError, ValueError) as exc:
        raise ApplyError(f"the regression record is not green for HEAD: {exc}") from exc
    regression = manifest.get("regression")
    record_digest = sha256_bytes(read_bytes(os.path.join(artifact, RECORD), RECORD))
    if not isinstance(regression, dict) or regression.get("record_sha256") != record_digest:
        raise ApplyError(f"{RECORD} is not the regression record the export manifest names")
    workflow = manifest.get("workflow")
    runner = record.get("runner")
    if isinstance(workflow, dict) and isinstance(runner, dict) and isinstance(runner.get("workflow"), dict):
        if workflow.get("run_id") != runner["workflow"].get("run_id"):
            raise ApplyError("the regression record comes from another workflow run than the export")

    environment = check_environment(artifact, manifest, record, head)

    entries = manifest.get("grids")
    if not isinstance(entries, list) or not all(isinstance(e, dict) for e in entries):
        raise ApplyError(f"{MANIFEST} grids must be a list of objects")
    by_name = {str(e.get("name")): e for e in entries}
    if sorted(by_name) != sorted(requested) or len(by_name) != len(entries):
        raise ApplyError(f"{MANIFEST} grids {sorted(by_name)} differ from requested {sorted(requested)}")
    grids = []
    for name in requested:
        spec, entry = GRIDS[name], by_name[name]
        grids.append((spec,) + check_grid(root, artifact, head, spec, entry))
    for spec, mode, _, _ in grids:
        if mode == "missing-only":
            check_kept_observations(root, spec)

    sources = manifest.get("sources")
    if not isinstance(sources, list) or not sources:
        raise ApplyError(f"{MANIFEST} lists no imported modules")
    for item in sources:
        if not isinstance(item, dict):
            raise ApplyError(f"{MANIFEST} sources must be objects")
        path = str(item.get("path"))
        try:
            digest = sha256_bytes(git_bytes(root, head, path))
        except CertificationError as exc:
            raise ApplyError(str(exc)) from exc
        if item.get("sha256") != digest:
            raise ApplyError(f"imported module {path} differs from its blob at HEAD")
    imported = {str(item.get("path")) for item in sources if isinstance(item, dict)}
    for spec, _, _, _ in grids:
        if by_name[spec.name].get("exporter") != spec.exporter or spec.exporter not in imported:
            raise ApplyError(f"{MANIFEST} does not list {spec.exporter} as the imported "
                             f"exporter of {spec.path}")
    return Plan(head, artifact, manifest, record, grids, environment)


def check_kept_observations(root: str, spec: GridSpec) -> None:
    """A missing-only export leaves every filled observed_vba cell as it was,
    and the finalizer and the binding then certify the whole grid for HEAD.
    That is true only if the kept observations came from the same VBA, so the
    committed main binding must still match every .bas module (#181)."""
    manifest = _manifest.load_manifest(root)
    binding = manifest.get("source_binding") if isinstance(manifest, dict) else None
    if not isinstance(binding, dict) or not binding:
        raise ApplyError(f"{spec.path}: a missing-only export needs the committed "
                         f"{_manifest.MANIFEST_NAME} to show which source produced the "
                         "observations it keeps; run a full export")
    current = _manifest.source_hashes(root)
    drift = sorted(path for path in set(binding) | set(current)
                   if binding.get(path) != current.get(path))
    if drift:
        raise ApplyError(f"{spec.path}: a missing-only export keeps the existing "
                         "observations, but these modules differ from the source they "
                         f"were bound to: {', '.join(drift)}; run a full export")


def check_environment(artifact: str, manifest: dict[str, object],
                      record: dict[str, object], head: str) -> bytes:
    path = os.path.join(artifact, ENVIRONMENT)
    raw = read_bytes(path, ENVIRONMENT)
    env = read_json(path, ENVIRONMENT)
    expected_keys = {"excel_version", "excel_build", "office_bitness", "commit_sha",
                     "commit_sha_recorded_by"}
    if set(env) != expected_keys:
        raise ApplyError(f"{ENVIRONMENT} keys {sorted(env)} differ from {sorted(expected_keys)}")
    if env["commit_sha"] != head:
        raise ApplyError(f"{ENVIRONMENT} records commit {env['commit_sha']}, not HEAD {head}")
    if env["commit_sha_recorded_by"] != ADAPTER:
        raise ApplyError(f"{ENVIRONMENT} commit_sha was not recorded by {ADAPTER}")
    written = read_json(os.path.join(artifact, EXCEL_WRITTEN_ENVIRONMENT),
                        EXCEL_WRITTEN_ENVIRONMENT)
    for key in ("excel_version", "excel_build", "office_bitness"):
        if written.get(key) != env[key]:
            raise ApplyError(f"{ENVIRONMENT} {key} differs from what Excel wrote")
    described = manifest.get("environment")
    if not isinstance(described, dict) or described.get("sha256") != sha256_bytes(raw):
        raise ApplyError(f"{ENVIRONMENT} is not the file the export manifest names")
    recorded = record["environment"]
    assert isinstance(recorded, dict)
    if (str(env["excel_version"]) != recorded["excel_version"]
            or str(env["excel_build"]) != recorded["excel_build"]
            or normalize_bitness(env["office_bitness"]) != normalize_bitness(recorded["office_bitness"])):
        raise ApplyError(f"{ENVIRONMENT} names another Excel than the regression record")
    return raw


def check_grid(root: str, artifact: str, head: str, spec: GridSpec,
               entry: dict[str, object]) -> tuple[str, bytes, dict[str, object]]:
    mode = entry.get("mode")
    if mode not in spec.macros:
        raise ApplyError(f"{spec.path}: mode {mode!r} has no exporter")
    assert isinstance(mode, str)
    if entry.get("path") != spec.path or entry.get("macro") != spec.macros[mode]:
        raise ApplyError(f"{spec.path}: the manifest names {entry.get('path')} "
                         f"exported by {entry.get('macro')}")
    if entry.get("artifact_path") != "grids/" + spec.path:
        raise ApplyError(f"{spec.path}: unexpected artifact path {entry.get('artifact_path')!r}")
    raw = read_bytes(os.path.join(artifact, "grids", *spec.path.split("/")), f"grid {spec.path}")
    if sha256_bytes(raw) != entry.get("raw_sha256"):
        raise ApplyError(f"{spec.path}: the artifact's bytes ({sha256_bytes(raw)}) differ from "
                         f"the raw SHA-256 the manifest records ({entry.get('raw_sha256')})")
    if entry.get("verified") is not True:
        raise ApplyError(f"{spec.path}: the runner did not verify this grid")
    try:
        baseline = git_bytes(root, head, spec.path)
    except CertificationError as exc:
        raise ApplyError(str(exc)) from exc
    if entry.get("baseline_sha256") != sha256_bytes(baseline):
        raise ApplyError(f"{spec.path}: the export did not start from the grid at HEAD")
    counts = verify_grid(spec, mode, baseline, raw)
    problems = counts["problems"]
    assert isinstance(problems, list)
    if problems:
        raise ApplyError(f"{spec.path} fails the export check:\n  " + "\n  ".join(problems))
    for key in ("rows", "in_scope_rows", "filled_rows", "error_rows", "out_of_scope_rows",
                "changed_observations"):
        if entry.get(key) != counts[key]:
            raise ApplyError(f"{spec.path}: the manifest reports {key}={entry.get(key)}, "
                             f"the artifact holds {counts[key]}")
    return mode, raw, counts


# --------------------------------------------------------------------------
# Apply
# --------------------------------------------------------------------------
def run_tool(root: str, argv: list[str]) -> subprocess.CompletedProcess[str]:
    print("$ python " + " ".join(argv), flush=True)
    proc = subprocess.run([sys.executable] + argv, cwd=root, capture_output=True, text=True)
    output = (proc.stdout + proc.stderr).rstrip()
    if output:
        print("    " + output.replace("\n", "\n    "), flush=True)
    return proc


def readme_detail(root: str) -> str:
    """The reasons render_readme_assurance.py gives, which refresh_evidence.py
    shortens to one line. A gate an export satisfies while
    release_readiness.json still lists its blocker open is the usual one."""
    script = os.path.join(root, "benchmark", "render_readme_assurance.py")
    if not os.path.isfile(script):
        return ""
    proc = subprocess.run([sys.executable, script, "--check"], cwd=root,
                          capture_output=True, text=True)
    lines = [line.strip() for line in (proc.stdout + proc.stderr).splitlines()
             if line.strip().startswith("- ")]
    return "".join("\n      " + line for line in lines)


def apply(root: str, plan: Plan) -> tuple[list[str], list[str]]:
    """Write the evidence and run the tools.

    Returns notes for the summary and the open items that keep the evidence
    from being ready to commit. Raises ApplyError when a step fails in a way
    that a hand edit cannot fix.
    """
    notes: list[str] = []
    open_items: list[str] = []
    # The finalizer records raw_sha256 from the working-tree bytes, so it must
    # see exactly what Excel wrote.
    for spec, _, raw, _ in plan.grids:
        with open(os.path.join(root, *spec.path.split("/")), "wb") as f:
            f.write(raw)
    with open(os.path.join(root, *ENV_PATH.split("/")), "wb") as f:
        f.write(plan.environment)

    record_path = os.path.join(plan.artifact, RECORD)
    finalize = [spec.finalize_flag for spec, _, _, _ in plan.grids if spec.finalize_flag]
    if finalize:
        proc = run_tool(root, ["benchmark/finalize_excel_certification.py", "--record",
                               record_path, *finalize, "--from-fresh-export"])
        if proc.returncode:
            raise ApplyError("finalize_excel_certification.py failed; the copied files are "
                             "left in place for inspection (git checkout -- . discards them)")

    # Then give the working tree the LF bytes Git stores, as a fresh checkout
    # of the evidence commit would, so local tools that hash a grid file (the
    # #72 result binding, for one) see the committed bytes, not CRLF.
    for spec, _, raw, _ in plan.grids:
        with open(os.path.join(root, *spec.path.split("/")), "wb") as f:
            f.write(canonical_grid_bytes(raw))

    binds = [spec.bind_flag for spec, _, _, _ in plan.grids if spec.bind_flag]
    proc = run_tool(root, ["benchmark/refresh_evidence.py", *binds])
    if proc.returncode:
        failed = re.findall(r"^\s+FAIL\s+(.+?)\s*$", proc.stdout, flags=re.MULTILINE)
        changed = [spec.path for spec, _, raw, _ in plan.grids if spec.finalize_flag
                   and canonical_grid_bytes(raw) != git_bytes(root, plan.head, spec.path)]
        if PRE_COMMIT_STEP in failed and changed:
            notes.append(f"refresh_evidence.py: '{PRE_COMMIT_STEP}' compares the record with "
                         "HEAD and fails until the evidence is committed; the snapshot check "
                         "below validates the record against the bytes to be committed")
            failed = [label for label in failed if label != PRE_COMMIT_STEP]
        elif not failed:
            failed = [f"refresh_evidence.py exited {proc.returncode}; see its output above"]
        for label in failed:
            open_items.append(label + (readme_detail(root) if label == README_STEP else ""))

    claimed = [spec.path for spec, _, _, _ in plan.grids if spec.finalize_flag]
    if claimed:
        snapshot = git(root, "stash", "create").strip() or plan.head
        for path in claimed:
            proc = run_tool(root, ["benchmark/check_excel_certification.py", "--record",
                                   RETAINED_RECORD, "--candidate-sha", plan.head,
                                   "--evidence-ref", snapshot, "--require-pass",
                                   "--require-grid", path, "--log-directory", plan.artifact])
            if proc.returncode:
                raise ApplyError(f"the finalized record does not validate {path} against the "
                                 "working tree; the files are left in place for inspection")
    return notes, open_items


def commit_message(plan: Plan) -> str:
    manifest = plan.manifest
    names = ", ".join(spec.path for spec, _, _, _ in plan.grids)
    workflow = manifest.get("workflow")
    excel = manifest.get("excel")
    harness = plan.record["harness"]
    assert isinstance(harness, dict) and isinstance(excel, dict)
    if isinstance(workflow, dict):
        origin = (f"the Excel VBA Regression workflow, run {workflow.get('run_id')} attempt "
                  f"{workflow.get('run_attempt')} on runner {workflow.get('runner')}")
    else:
        origin = f"{ADAPTER}, run locally"
    lines = [f"evidence: Excel export of {', '.join(spec.name for spec, _, _, _ in plan.grids)} "
             f"at {plan.head[:7]}",
             "",
             f"Exported unattended by {origin}, from candidate {plan.head} "
             f"(Excel {excel.get('version')} build {excel.get('build')}, "
             f"{excel.get('office_bitness')}). The regression in the same job passed "
             f"{harness['passed']}/{harness['assertions']} assertions.",
             ""]
    for spec, mode, raw, counts in plan.grids:
        if spec.kind == "generated":
            lines.append(f"- {spec.path}: written whole by {spec.macros[mode]}; "
                         f"{counts['rows']} rows with their committed key fields, "
                         f"{counts['error_rows']} with an ERROR status, "
                         f"{counts['changed_observations']} row(s) changed; "
                         f"raw {sha256_bytes(raw)}")
            continue
        lines.append(f"- {spec.path}: {mode} export by {spec.macros[mode]}; "
                     f"{counts['rows']} rows, {counts['in_scope_rows']} in scope and filled, "
                     f"{counts['error_rows']} ERROR, {counts['changed_observations']} "
                     f"observation(s) changed, {counts['out_of_scope_rows']} out-of-scope rows "
                     f"unchanged; raw {sha256_bytes(raw)}")
    lines.append("")
    steps = [f"finalize_excel_certification.py "
             f"{' '.join(s.finalize_flag for s, _, _, _ in plan.grids if s.finalize_flag)} "
             "--from-fresh-export"] if any(s.finalize_flag for s, _, _, _ in plan.grids) else []
    binds = " ".join(s.bind_flag for s, _, _, _ in plan.grids if s.bind_flag)
    steps.append(f"refresh_evidence.py {binds}".rstrip())
    lines.append(f"Applied with benchmark/apply_excel_export.py ({names}): it checked the "
                 f"artifact against HEAD, then ran {' and '.join(steps)}, and validated the "
                 "finalized record against the staged tree with check_excel_certification.py.")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0] if __doc__ else None)
    ap.add_argument("--artifact", required=True,
                    help="directory holding the downloaded excel-grid-export-<run id> artifact")
    ap.add_argument("--grid", action="append", choices=sorted(GRIDS),
                    help="a grid the artifact must hold; repeat for each. Default: the "
                         "grids the export manifest lists")
    a = ap.parse_args(argv)
    artifact = os.path.abspath(a.artifact)
    try:
        plan = check_artifact(ROOT, artifact, a.grid)
        print(f"Artifact accepted: candidate {plan.head}; grids "
              + ", ".join(spec.path for spec, _, _, _ in plan.grids))
        notes, open_items = apply(ROOT, plan)
    except ApplyError as exc:
        print(f"REFUSED: {exc}")
        return 1
    except CertificationError as exc:
        print(f"FAILED: {exc}")
        return 1

    changed = git(ROOT, "status", "--porcelain", "--untracked-files=no").splitlines()
    print()
    for note in notes:
        print("note: " + note)
    if open_items:
        print("NOT READY: the export is applied, but refresh_evidence.py still fails:")
        for item in open_items:
            print("  - " + item)
        print("Fix these in the working tree (for example, resolve in "
              "benchmark/release_readiness.json a blocker the export satisfies), then run "
              "`python benchmark/refresh_evidence.py` (no bind flags: the manifests are "
              "already bound) until only the pre-commit record check fails, and commit.")
    else:
        print("Evidence ready to commit.")
    print("Changed files:")
    for line in changed:
        print("  " + line)
    print()
    print("Commit message:")
    print("-" * 72)
    print(commit_message(plan))
    print("-" * 72)
    print()
    print("After committing, validate the evidence commit:")
    for spec, _, _, _ in plan.grids:
        if spec.finalize_flag:
            print(f"  python benchmark/check_excel_certification.py --record {RETAINED_RECORD} "
                  f"--candidate-sha {plan.head} --evidence-ref HEAD --require-pass "
                  f"--require-grid {spec.path} --log-directory {artifact}")
    print(f"  python benchmark/check_manifest_provenance.py --since {plan.head}")
    print("  python benchmark/refresh_evidence.py --check")
    return 1 if open_items else 0


if __name__ == "__main__":
    sys.exit(main())
