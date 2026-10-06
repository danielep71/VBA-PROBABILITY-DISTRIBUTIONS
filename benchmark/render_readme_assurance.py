"""Render the root README assurance block from committed evidence (#28).

Single source of truth: every figure in the three generated regions of the root
README - the evidence badges, the "Assurance at a glance" table and the
evidence-state list - comes from one data model built here from committed,
machine-readable authorities. Nothing is scraped from GitHub and no value falls
back to a previous figure or to zero. Local Git history is required to validate
the retained Excel candidate and its exact source bytes; rendering stays offline.

    python render_readme_assurance.py --write   regenerate the three regions
    python render_readme_assurance.py --check   fail if README.md differs from
                                                a fresh render (hosted gate)

Exit status:
    0  rendered (or README current) and no blocking state;
    1  a required input is missing, malformed or contradicts another input
       (nothing is written), README.md differs from the render (--check), or
       the evidence is stale: the Excel regression sources, the main-grid
       binding or the holdout binding no longer match the checked-out source.
       Stale evidence is still rendered, visibly, by --write.

Incomplete certification (a grid the retained Excel record did not export, or
whose digest differs) is not staleness: it is rendered as incomplete and must be
owned by an open blocker in release_readiness.json. The authorities are listed
in PROVENANCE.md, "README assurance block".
"""
import argparse
import csv
import difflib
import json
import os
import re
import sys
from collections import Counter
from datetime import date

from _manifest import normalized_hash, verify_holdout_binding, verify_source_binding
from check_grid_coverage import _load_json, evaluate_paths, validate_strict, validate_transition
from excel_certification import (CertificationError, _source_map, _timestamp,
                                 row_count_file, validate_record)

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
REPO_URL = "https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS"

README = "README.md"
PUBLIC_API = "docs/PUBLIC_API.txt"
POLICY = ".github/excel-evidence-policy.json"
EXCEL_RECORD = "benchmark/excel_regression_record.json"
CONTRACTS = "benchmark/accuracy_contracts.csv"
GRID = "benchmark/probability_accuracy_grid.csv"
MAIN_MANIFEST = "benchmark/observation_manifest.json"
SUMMARY = "benchmark/accuracy_summary.md"
EXEMPTIONS = "benchmark/accuracy_row_exemptions.json"
FINGERPRINT = "benchmark/coverage_debt_v1_0_0.json"
HOLDOUT_GRID = "benchmark/holdout/holdout_grid.csv"
HOLDOUT_MANIFEST = "benchmark/holdout/holdout_manifest.json"
HOLDOUT_SUMMARY = "benchmark/holdout/holdout_summary.md"
READINESS = "benchmark/release_readiness.json"

# Paths every render reads, for documentation and for test fixtures. The .bas
# sources are read too, through the manifest bindings and the Excel record.
INPUTS = (README, PUBLIC_API, POLICY, EXCEL_RECORD, CONTRACTS, GRID, MAIN_MANIFEST,
          SUMMARY, EXEMPTIONS, FINGERPRINT, HOLDOUT_GRID, HOLDOUT_MANIFEST,
          HOLDOUT_SUMMARY, READINESS)

GENERATOR = "benchmark/render_readme_assurance.py"
REGIONS = ("assurance badges", "assurance table", "evidence state")

CONTRACT_STATUSES = ("active", "characterization_only")
SUMMARY_HEADER = "| Contract | Measure | Metric | Threshold | Worst error | Points | Verdict |"
HOLDOUT_HEADER = ("| Contract | Metric | Threshold | Holdout worst | Points | Margin | "
                  "Provenance | Verdict |")
# Checked in this order: "PENDING - ..." cells never contain PASS/FAIL, but the
# explicit order keeps the classification independent of that accident.
MAIN_VERDICTS = ("CHARACTERIZATION ONLY", "KNOWN LIMITATION", "PENDING", "FAIL", "PASS")
HOLDOUT_VERDICTS = ("PASS", "FAIL", "INCOMPLETE", "NO OBS")

READINESS_KEYS = {"schema_version", "release", "tracker_issue", "reviewed", "certified",
                  "blockers"}
BLOCKER_KEYS = {"issue", "class", "gate", "status", "summary"}
BLOCKER_CLASSES = {
    "silent_wrong": "silent wrong result",
    "numerical_uncertainty": "unresolved numerical degradation",
    "numerical_defect": "numerical defects",
    "assurance": "assurance gates",
}
GATES = ("main_grid_coverage", "main_grid_export", "holdout_export")
# A blocker that the committed evidence cannot score must not disappear by
# being deleted from the registry: it may only move to "resolved" with the
# evidence that resolved it. Removing an entry here is a reviewed code change.
REQUIRED_BLOCKERS = {34: "silent_wrong", 35: "numerical_uncertainty"}

EXCEL_COLOR = "d97706"
CONTRACTS_COLOR = "0f766e"
HOLDOUT_COLOR = "4c1d95"
ALERT_COLOR = "b91c1c"


class AssuranceError(ValueError):
    """A required input is missing, malformed or contradicts another input."""


# --- inputs ------------------------------------------------------------------

def _path(root, rel):
    return os.path.join(root, *rel.split("/"))


def _read_text(root, rel):
    try:
        with open(_path(root, rel), encoding="utf-8") as f:
            return f.read().replace("\r\n", "\n")
    except FileNotFoundError:
        raise AssuranceError(f"required input missing: {rel}")
    except (OSError, UnicodeDecodeError) as exc:
        raise AssuranceError(f"cannot read {rel}: {exc}")


def _read_json(root, rel):
    try:
        value = json.loads(_read_text(root, rel))
    except ValueError as exc:
        raise AssuranceError(f"{rel} is not valid JSON: {exc}")
    if not isinstance(value, dict):
        raise AssuranceError(f"{rel} must be a JSON object")
    return value


def _read_csv(root, rel):
    text = _read_text(root, rel)
    rows = list(csv.DictReader(text.splitlines()))
    if not rows:
        raise AssuranceError(f"{rel} has no data rows")
    return rows


def api_facts(root):
    names = []
    for number, line in enumerate(_read_text(root, PUBLIC_API).splitlines(), start=1):
        if not line.strip() or line.startswith("#"):
            continue
        fields = line.split("\t")
        if len(fields) != 4 or not fields[2].startswith("K_STATS_"):
            raise AssuranceError(f"{PUBLIC_API} line {number} is not a K_STATS_ declaration")
        names.append(fields[2])
    if not names:
        raise AssuranceError(f"{PUBLIC_API} declares no public functions")
    if len(names) != len(set(names)):
        raise AssuranceError(f"{PUBLIC_API} declares a function twice")
    surfaces = {name.split("_")[2] for name in names}
    return {"functions": len(names), "surfaces": len(surfaces)}


def contract_facts(root):
    rows = _read_csv(root, CONTRACTS)
    ids = [row.get("contract_id", "").strip() for row in rows]
    if not all(ids) or len(ids) != len(set(ids)):
        raise AssuranceError(f"{CONTRACTS} has a blank or duplicate contract_id")
    statuses = Counter(row.get("status", "").strip() for row in rows)
    unknown = sorted(set(statuses) - set(CONTRACT_STATUSES))
    if unknown:
        raise AssuranceError(f"{CONTRACTS} has unknown status(es): {', '.join(unknown)}")
    return {"ids": set(ids), "rows": len(rows), "active": statuses["active"],
            "characterization": statuses["characterization_only"]}


def _table_rows(text, header, rel):
    lines = text.splitlines()
    if lines.count(header) != 1:
        raise AssuranceError(f"{rel}: expected exactly one verdict table")
    rows = []
    for line in lines[lines.index(header) + 2:]:
        if not line.startswith("|"):
            break
        cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
        rows.append((cells[0], cells[-1]))
    if not rows:
        raise AssuranceError(f"{rel}: verdict table is empty")
    ids = [contract for contract, _ in rows]
    if len(ids) != len(set(ids)):
        raise AssuranceError(f"{rel}: a contract appears twice in the verdict table")
    return rows


def _classify(cell, verdicts, rel, contract):
    for verdict in verdicts:
        if verdict in cell:
            return verdict
    raise AssuranceError(f"{rel}: unrecognized verdict {cell!r} for {contract}")


def _one(pattern, text, rel, what):
    found = re.findall(pattern, text, flags=re.MULTILINE)
    if len(found) != 1:
        raise AssuranceError(f"{rel}: expected exactly one {what}")
    return found[0]


def summary_facts(root, contracts):
    text = _read_text(root, SUMMARY)
    rows = _table_rows(text, SUMMARY_HEADER, SUMMARY)
    ids = {contract for contract, _ in rows}
    if ids != contracts["ids"]:
        missing = sorted(contracts["ids"] - ids)[:3]
        extra = sorted(ids - contracts["ids"])[:3]
        raise AssuranceError(
            f"{SUMMARY} verdicts do not cover {CONTRACTS} exactly"
            + (f"; unscored: {', '.join(missing)}" if missing else "")
            + (f"; unregistered: {', '.join(extra)}" if extra else ""))
    verdicts = Counter(_classify(cell, MAIN_VERDICTS, SUMMARY, contract)
                       for contract, cell in rows)
    tally = _one(r"^> \*\*Verdict tally\*\* — FAIL: (\d+), KNOWN LIMITATION: (\d+), "
                 r"CHARACTERIZATION ONLY: (\d+), PENDING: (\d+)\.$", text, SUMMARY,
                 "verdict tally")
    stated = dict(zip(("FAIL", "KNOWN LIMITATION", "CHARACTERIZATION ONLY", "PENDING"),
                      (int(n) for n in tally)))
    for verdict, count in stated.items():
        if verdicts[verdict] != count:
            raise AssuranceError(f"{SUMMARY}: tally states {verdict}: {count} but the "
                                 f"table holds {verdicts[verdict]}")
    if verdicts["CHARACTERIZATION ONLY"] != contracts["characterization"]:
        raise AssuranceError(f"{SUMMARY} and {CONTRACTS} disagree on characterization-only "
                             "contracts")
    return {
        "verdicts": {verdict: verdicts[verdict] for verdict in MAIN_VERDICTS},
        "source_commit": _one(r"^- source commit: `([0-9a-f]+)`", text, SUMMARY,
                              "source commit"),
        "grid_rows": int(_one(r"^- total grid rows: ([0-9,]+)$", text, SUMMARY,
                              "total grid rows").replace(",", "")),
    }


def grid_rows(root, rel):
    return len(_read_csv(root, rel))


def main_binding_facts(root, summary):
    manifest = _read_json(root, MAIN_MANIFEST)
    commit = manifest.get("commit_sha")
    if not isinstance(commit, str) or not re.fullmatch(r"[0-9a-f]{7,40}", commit):
        raise AssuranceError(f"{MAIN_MANIFEST}: commit_sha is missing or malformed")
    if summary["source_commit"] != commit:
        raise AssuranceError(f"{SUMMARY} reports source commit {summary['source_commit']} "
                             f"but {MAIN_MANIFEST} records {commit}")
    problems = verify_source_binding(root, manifest, _path(root, GRID), _path(root, CONTRACTS))
    return {"commit": commit, "problems": problems}


def coverage_facts(root):
    result = evaluate_paths(_path(root, GRID), _path(root, CONTRACTS), _path(root, EXEMPTIONS))
    if result["errors"]:
        raise AssuranceError("main-grid coverage is invalid: " + result["errors"][0])
    if os.path.exists(_path(root, FINGERPRINT)):
        mode = "transition"
        fingerprint, errors = _load_json(_path(root, FINGERPRINT), "coverage-debt fingerprint")
        if not errors:
            errors = validate_transition(fingerprint, result)
    else:
        mode = "strict"
        errors = validate_strict(result)
    if errors:
        raise AssuranceError(f"main-grid coverage fails its {mode} guard: {errors[0]}")
    return {"mode": mode, "total": result["total_rows"], "main": result["main_rows"],
            "claimed": result["claimed_rows"], "exempt": result["exempt_rows"],
            "missing": result["missing_rows"]}


def holdout_facts(root, contracts):
    text = _read_text(root, HOLDOUT_SUMMARY)
    rows = _table_rows(text, HOLDOUT_HEADER, HOLDOUT_SUMMARY)
    unregistered = sorted({contract for contract, _ in rows} - contracts["ids"])
    if unregistered:
        raise AssuranceError(f"{HOLDOUT_SUMMARY} scores unregistered contract(s): "
                             + ", ".join(unregistered[:3]))
    verdicts = Counter()
    for contract, cell in rows:
        if cell not in HOLDOUT_VERDICTS:
            raise AssuranceError(f"{HOLDOUT_SUMMARY}: unrecognized verdict {cell!r} for "
                                 f"{contract}")
        verdicts[cell] += 1
    npass, nfail, nincomplete, total = _one(
        r"^> (\d+) pass, (\d+) fail(?:, (\d+) incomplete)? across (\d+) contract\(s\) ",
        text, HOLDOUT_SUMMARY, "holdout tally")
    stated = (int(npass), int(nfail), int(nincomplete or 0), int(total))
    if stated != (verdicts["PASS"], verdicts["FAIL"], verdicts["INCOMPLETE"], len(rows)):
        raise AssuranceError(f"{HOLDOUT_SUMMARY}: tally {stated} contradicts its table")
    manifest = _read_json(root, HOLDOUT_MANIFEST)
    observations = manifest.get("observation_row_count")
    if type(observations) is not int or observations <= 0:
        raise AssuranceError(f"{HOLDOUT_MANIFEST}: observation_row_count is missing or invalid")
    problems = verify_holdout_binding(root, manifest, _path(root, HOLDOUT_GRID),
                                      _path(root, CONTRACTS))
    return {"contracts": len(rows), "pass": verdicts["PASS"], "fail": verdicts["FAIL"],
            "incomplete": verdicts["INCOMPLETE"], "observations": observations,
            "problems": problems}


def excel_facts(root):
    """Validate the retained record against its actual candidate, then assess
    working-tree freshness separately. A changed checkout can be rendered STALE;
    an invalid candidate/source claim must never be rendered as evidence."""
    policy = _read_json(root, POLICY)
    record = _read_json(root, EXCEL_RECORD)
    try:
        # Canonical validation checks candidate existence and raw source blobs,
        # not just the shape of the SHA/digests or their working-tree matches.
        committed_policy = validate_record(record, root)
        if policy != committed_policy:
            raise CertificationError("Excel evidence policy differs from committed HEAD")
        green = all(stage["status"] == "PASS" for stage in record["stages"].values())
        if green:
            # PASS stages alone cannot certify an unobserved Excel environment.
            validate_record(record, root, require_pass=True)
        finished = _timestamp(record["finished_at"], "finished_at")
        sources = _source_map(record)
    except CertificationError as exc:
        raise AssuranceError(f"{EXCEL_RECORD}: {exc}")

    stale = []
    for rel, digest in sources.items():
        if not os.path.isfile(_path(root, rel)):
            stale.append(f"{rel} (missing)")
        elif normalized_hash(_path(root, rel)) != digest:
            stale.append(rel)
    grids = {}
    for rel, entry in record["grids"].items():
        if not entry["exported"]:
            grids[rel] = "not exported"
        elif (os.path.isfile(_path(root, rel))
              and normalized_hash(_path(root, rel)) == entry["sha256"]
              and row_count_file(_path(root, rel)) == entry["row_count"]):
            grids[rel] = "certified"
        else:
            grids[rel] = "digest differs"
    harness = record["harness"]
    environment = record["environment"]
    return {
        "candidate": record["candidate_sha"][:7],
        "passed": harness["passed"],
        "assertions": harness["assertions"],
        "green": green,
        "excel": f"Excel {environment['excel_version']} build {environment['excel_build']}, "
                 f"{environment['office_bitness']}",
        "finished": finished.date().isoformat(),
        "stale": stale,
        "grids": grids,
    }


def readiness_facts(root):
    data = _read_json(root, READINESS)
    if set(data) != READINESS_KEYS:
        raise AssuranceError(f"{READINESS}: keys differ from schema")
    if data["schema_version"] != "release-readiness/v1":
        raise AssuranceError(f"{READINESS}: unsupported schema_version")
    if not isinstance(data["release"], str) or not re.fullmatch(r"v\d+\.\d+\.\d+", data["release"]):
        raise AssuranceError(f"{READINESS}: release must look like vX.Y.Z")
    if type(data["tracker_issue"]) is not int or data["tracker_issue"] <= 0:
        raise AssuranceError(f"{READINESS}: tracker_issue must be a positive integer")
    try:
        date.fromisoformat(data["reviewed"])
    except (TypeError, ValueError):
        raise AssuranceError(f"{READINESS}: reviewed must be a YYYY-MM-DD date")
    if type(data["certified"]) is not bool:
        raise AssuranceError(f"{READINESS}: certified must be true or false")
    blockers = data["blockers"]
    if not isinstance(blockers, list):
        raise AssuranceError(f"{READINESS}: blockers must be a list")
    previous = 0
    for blocker in blockers:
        if not isinstance(blocker, dict):
            raise AssuranceError(f"{READINESS}: every blocker must be an object")
        resolved = blocker.get("status") == "resolved"
        if set(blocker) != BLOCKER_KEYS | ({"resolution"} if resolved else set()):
            raise AssuranceError(f"{READINESS}: blocker keys differ from schema "
                                 f"(issue {blocker.get('issue')!r})")
        issue = blocker["issue"]
        if type(issue) is not int or issue <= previous:
            raise AssuranceError(f"{READINESS}: blocker issues must be unique positive "
                                 "integers in ascending order")
        previous = issue
        if blocker["class"] not in BLOCKER_CLASSES:
            raise AssuranceError(f"{READINESS}: #{issue} has unknown class {blocker['class']!r}")
        if blocker["gate"] is not None and blocker["gate"] not in GATES:
            raise AssuranceError(f"{READINESS}: #{issue} has unknown gate {blocker['gate']!r}")
        if blocker["status"] not in ("open", "resolved"):
            raise AssuranceError(f"{READINESS}: #{issue} status must be open or resolved")
        summary = blocker["summary"]
        if not isinstance(summary, str) or not summary.strip() or "\n" in summary or "|" in summary:
            raise AssuranceError(f"{READINESS}: #{issue} needs a single-line summary")
        if resolved:
            resolution = blocker["resolution"]
            if (not isinstance(resolution, dict) or set(resolution) != {"evidence"}
                    or not isinstance(resolution["evidence"], str)
                    or not resolution["evidence"].strip()):
                raise AssuranceError(f"{READINESS}: resolved #{issue} must name its evidence")
    by_issue = {blocker["issue"]: blocker for blocker in blockers}
    for issue, klass in REQUIRED_BLOCKERS.items():
        if issue not in by_issue or by_issue[issue]["class"] != klass:
            raise AssuranceError(f"{READINESS}: #{issue} must stay registered as {klass} "
                                 "(open, or resolved with evidence)")
    open_gates = Counter(b["gate"] for b in blockers if b["status"] == "open" and b["gate"])
    doubled = sorted(gate for gate, count in open_gates.items() if count > 1)
    if doubled:
        raise AssuranceError(f"{READINESS}: more than one open blocker owns {doubled[0]}")
    return data


# --- model ---------------------------------------------------------------------

def build_model(root=ROOT):
    """Read every authority, cross-check them, and return the render model."""
    api = api_facts(root)
    contracts = contract_facts(root)
    summary = summary_facts(root, contracts)
    rows = grid_rows(root, GRID)
    if summary["grid_rows"] != rows:
        raise AssuranceError(f"{SUMMARY} reports {summary['grid_rows']:,} grid rows but "
                             f"{GRID} holds {rows:,}")
    binding = main_binding_facts(root, summary)
    coverage = coverage_facts(root)
    if coverage["total"] != rows:
        raise AssuranceError("coverage evaluation and grid disagree on the row count")
    holdout = holdout_facts(root, contracts)
    excel = excel_facts(root)
    readiness = readiness_facts(root)

    gates = {
        "main_grid_coverage": coverage["missing"] == 0,
        "main_grid_export": excel["grids"].get(GRID) == "certified",
        "holdout_export": excel["grids"].get(HOLDOUT_GRID) == "certified",
    }
    owners = {}
    for gate, satisfied in gates.items():
        owner = [b["issue"] for b in readiness["blockers"]
                 if b["gate"] == gate and b["status"] == "open"]
        if not satisfied and not owner:
            raise AssuranceError(f"{gate} is incomplete but {READINESS} has no open blocker "
                                 "for it")
        if satisfied and owner:
            raise AssuranceError(f"{gate} is satisfied but {READINESS} still lists "
                                 f"#{owner[0]} open; resolve that entry with its evidence")
        owners[gate] = owner[0] if owner else None

    stale = []
    if excel["stale"]:
        stale.append("Excel regression sources changed since the certified run: "
                     + ", ".join(excel["stale"]))
    if binding["problems"]:
        stale.append("main-grid observations are not bound to the current source: "
                     + binding["problems"][0])
    if holdout["problems"]:
        stale.append("holdout observations are not bound to the current source: "
                     + holdout["problems"][0])

    open_blockers = [b for b in readiness["blockers"] if b["status"] == "open"]
    verdicts = summary["verdicts"]
    clean = (not stale and not open_blockers and all(gates.values()) and excel["green"]
             and verdicts["FAIL"] == verdicts["PENDING"] == verdicts["KNOWN LIMITATION"] == 0
             and holdout["fail"] == holdout["incomplete"] == 0)
    if readiness["certified"] and not clean:
        raise AssuranceError(f"{READINESS} claims {readiness['release']} is certified, but "
                             "blockers, incomplete gates or non-passing evidence remain")
    return {"api": api, "contracts": contracts, "verdicts": verdicts, "grid_rows": rows,
            "binding": binding, "coverage": coverage, "holdout": holdout, "excel": excel,
            "readiness": readiness, "open_blockers": open_blockers, "owners": owners,
            "stale": stale}


# --- rendering -----------------------------------------------------------------

def _issue(number):
    return f"[#{number}]({REPO_URL}/issues/{number})"


def _shield(text):
    return (text.replace("-", "--").replace("_", "__").replace(" ", "_")
            .replace("/", "%2F"))


def _badge(alt, label, message, color, link, logo=""):
    return (f"[![{alt}](https://img.shields.io/badge/{_shield(label)}-{_shield(message)}-"
            f"{color}?style=for-the-badge{logo})]({link})")


def render_badges(model):
    excel, verdicts, holdout = model["excel"], model["verdicts"], model["holdout"]
    if excel["stale"]:
        excel_message, excel_color = "STALE", ALERT_COLOR
    else:
        excel_message = f"{excel['passed']} of {excel['assertions']} passed"
        excel_color = EXCEL_COLOR if excel["green"] else ALERT_COLOR
    blocking = verdicts["FAIL"] + verdicts["PENDING"] + verdicts["KNOWN LIMITATION"]
    if model["binding"]["problems"]:
        contracts_message, contracts_color = "STALE", ALERT_COLOR
    else:
        contracts_message = f"{verdicts['PASS']} of {model['contracts']['active']} PASS"
        contracts_color = CONTRACTS_COLOR if not blocking else ALERT_COLOR
    if holdout["problems"]:
        holdout_message, holdout_color = "STALE", ALERT_COLOR
    else:
        holdout_message = f"{holdout['pass']} of {holdout['contracts']} PASS"
        holdout_color = (HOLDOUT_COLOR if not (holdout["fail"] or holdout["incomplete"])
                         else ALERT_COLOR)
    return "\n".join([
        _badge("Excel CI", "Excel Regression", excel_message, excel_color,
               "docs/EXCEL_VBA_CI.md", "&logo=githubactions&logoColor=white"),
        _badge("Accuracy Contracts", "Accuracy Contracts", contracts_message, contracts_color,
               "benchmark/accuracy_summary.md"),
        _badge("Independent Holdout", "Holdout", holdout_message, holdout_color,
               "benchmark/holdout/holdout_summary.md"),
    ])


def render_table(model):
    api, excel, contracts = model["api"], model["excel"], model["contracts"]
    verdicts, coverage, holdout = model["verdicts"], model["coverage"], model["holdout"]
    readiness = model["readiness"]
    if excel["stale"]:
        excel_row = (f"| **STALE** | Excel regression evidence: the regression sources changed "
                     f"after the certified run on `{excel['candidate']}` |")
    else:
        excel_row = (f"| **{excel['passed']:,}** | of {excel['assertions']:,} deterministic VBA "
                     f"assertions passed in the retained Excel run on `{excel['candidate']}`; "
                     "regression sources unchanged since |")
    contract_row = (f"| **{contracts['active']:,}** | active accuracy contracts "
                    f"({contracts['rows']:,} registry rows; {contracts['characterization']:,} "
                    f"characterization-only): {verdicts['PASS']:,} PASS, {verdicts['FAIL']:,} "
                    f"FAIL, {verdicts['PENDING']:,} PENDING |")
    if model["binding"]["problems"]:
        contract_row = contract_row.replace("| active", "| **STALE:** active", 1)
    grid_row = (f"| **{model['grid_rows']:,}** | observation rows; main grid "
                f"{coverage['main']:,}: {coverage['claimed']:,} claimed, "
                f"{coverage['exempt']:,} exempt, {coverage['missing']:,} without a disposition"
                + (f" ({_issue(model['owners']['main_grid_coverage'])})"
                   if coverage["missing"] else "") + " |")
    if holdout["problems"]:
        holdout_row = (f"| **STALE** | independent holdout: {holdout['observations']:,} "
                       "observations no longer bound to the current source |")
    else:
        holdout_row = (f"| **{holdout['pass']:,}** | of {holdout['contracts']:,} contracts "
                       f"PASS on an **independent holdout** of {holdout['observations']:,} "
                       "observations |")
    if readiness["certified"]:
        release_row = f"| **Certified** | {readiness['release']} release |"
    else:
        headline = [b for b in model["open_blockers"]
                    if b["class"] in ("silent_wrong", "numerical_uncertainty")]
        release_row = (f"| **Not certified** | {readiness['release']} release: "
                       f"{len(model['open_blockers']):,} registered open blockers"
                       + (", including " + ", ".join(_issue(b["issue"]) for b in headline)
                          if headline else "")
                       + f"; tracker {_issue(readiness['tracker_issue'])} |")
    return "\n".join([
        "| | |",
        "|---:|:---|",
        f"| **{api['functions']:,}** | worksheet-facing `K_STATS_` functions across "
        f"**{api['surfaces']:,}** distribution surfaces |",
        excel_row,
        contract_row,
        grid_row,
        holdout_row,
        release_row,
        "",
        f"*Generated by `{GENERATOR}` from committed evidence; the Accuracy Gate fails if "
        "this block differs from a fresh render. Sources: "
        "[benchmark/PROVENANCE.md](benchmark/PROVENANCE.md#readme-assurance-block).*",
    ])


def render_state(model):
    excel, verdicts, coverage = model["excel"], model["verdicts"], model["coverage"]
    holdout, readiness, owners = model["holdout"], model["readiness"], model["owners"]

    def export(rel, gate, what):
        state = excel["grids"].get(rel, "not exported")
        if state == "certified":
            return f"- **{what} export certification** — certified by the retained Excel record."
        reason = ("the retained Excel record did not export it" if state == "not exported"
                  else "the retained Excel record's digest differs from the committed grid")
        return (f"- **{what} export certification** — **incomplete**: {reason} "
                f"({_issue(owners[gate])}).")

    lines = ["**Evidence state.** Generated from the committed records listed in "
             "[benchmark/PROVENANCE.md](benchmark/PROVENANCE.md#readme-assurance-block). "
             "Documentation and tooling commits do not re-run Excel; the Excel evidence stays "
             "current only while the regression source content is unchanged.", ""]
    if excel["stale"]:
        lines.append(f"- **Excel regression** — **STALE**: "
                     f"{', '.join(f'`{p}`' for p in excel['stale'])} changed after the "
                     f"certified run on `{excel['candidate']}`. Re-run the Excel regression.")
    else:
        lines.append(f"- **Excel regression** — {excel['passed']:,}/{excel['assertions']:,} "
                     f"assertions {'passed' if excel['green'] else 'recorded, not all stages passed'}"
                     f" on `{excel['candidate']}` ({excel['excel']}; finished "
                     f"{excel['finished']}). Regression sources are unchanged since that run.")
    if model["binding"]["problems"]:
        lines.append(f"- **Main-grid observations** — **STALE**: "
                     f"{model['binding']['problems'][0]}.")
    else:
        lines.append(f"- **Main-grid observations** — {model['grid_rows']:,} rows bound to the "
                     f"current source by `observation_manifest.json` (exported from "
                     f"`{model['binding']['commit']}`).")
    lines.append(export(GRID, "main_grid_export", "Main-grid"))
    if coverage["missing"]:
        lines.append(f"- **Main-grid coverage** — `{coverage['mode']}` mode: "
                     f"{coverage['claimed']:,} of {coverage['main']:,} rows claimed, "
                     f"{coverage['exempt']:,} exempt, **{coverage['missing']:,} without a "
                     f"disposition** ({_issue(owners['main_grid_coverage'])}).")
    else:
        lines.append(f"- **Main-grid coverage** — complete: every one of {coverage['main']:,} "
                     "main-grid rows is claimed or exempt.")
    lines.append(f"- **Contract verdicts** — {verdicts['PASS']:,} PASS, {verdicts['FAIL']:,} "
                 f"FAIL, {verdicts['KNOWN LIMITATION']:,} KNOWN LIMITATION, "
                 f"{verdicts['PENDING']:,} PENDING; {verdicts['CHARACTERIZATION ONLY']:,} "
                 "characterization-only.")
    if holdout["problems"]:
        lines.append(f"- **Independent holdout** — **STALE**: {holdout['problems'][0]}.")
    else:
        tail = (f", {holdout['incomplete']:,} incomplete" if holdout["incomplete"] else "")
        lines.append(f"- **Independent holdout** — {holdout['pass']:,} PASS, "
                     f"{holdout['fail']:,} FAIL{tail} of {holdout['contracts']:,} contracts on "
                     f"{holdout['observations']:,} observations bound to the current source.")
    lines.append(export(HOLDOUT_GRID, "holdout_export", "Holdout"))
    if readiness["certified"]:
        lines.append(f"- **Release {readiness['release']}** — certified.")
    else:
        lines.append(f"- **Release {readiness['release']}** — **not certified**. Open "
                     f"blockers registered in `release_readiness.json` (reviewed "
                     f"{readiness['reviewed']}); {_issue(readiness['tracker_issue'])} tracks "
                     "the full release:")
        for klass, label in BLOCKER_CLASSES.items():
            group = [b for b in model["open_blockers"] if b["class"] == klass]
            if not group:
                continue
            if klass in ("silent_wrong", "numerical_uncertainty"):
                for blocker in group:
                    lines.append(f"  - {label}: {_issue(blocker['issue'])} — "
                                 f"{blocker['summary']}")
            else:
                lines.append(f"  - {label}: " + ", ".join(_issue(b["issue"]) for b in group))
    return "\n".join(lines)


def render(model):
    return {"assurance badges": render_badges(model),
            "assurance table": render_table(model),
            "evidence state": render_state(model)}


# --- README splicing -------------------------------------------------------------

def _markers(region):
    return (f"<!-- BEGIN generated: {region} via {GENERATOR}. Do not hand-edit. -->",
            f"<!-- END generated: {region} -->")


def splice(text, rendered):
    """Replace every generated region; fail unless each appears once, in order."""
    position = 0
    for region in REGIONS:
        begin, end = _markers(region)
        if text.count(begin) != 1 or text.count(end) != 1:
            raise AssuranceError(f"{README}: expected exactly one '{region}' region")
        i = text.index(begin)
        j = text.index(end)
        if i < position or j < i:
            raise AssuranceError(f"{README}: generated regions are missing or out of order")
        text = text[:i] + begin + "\n" + rendered[region] + "\n" + text[j:]
        position = text.index(end)
    return text


def main(argv=None, root=ROOT):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    mode = ap.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true", help="regenerate the README regions")
    mode.add_argument("--check", action="store_true", help="fail if README.md is not current")
    args = ap.parse_args(argv)
    try:
        model = build_model(root)
        current = _read_text(root, README)
        expected = splice(current, render(model))
    except AssuranceError as exc:
        print(f"FAIL: README assurance block\n  - {exc}")
        return 1

    status = 0
    if args.write:
        if expected != current:
            with open(_path(root, README), "w", encoding="utf-8", newline="\n") as f:
                f.write(expected)
            print(f"updated the generated assurance regions in {README}")
        else:
            print(f"{README} assurance regions already current")
    elif expected != current:
        print(f"FAIL: {README} assurance regions differ from a fresh render; run "
              "`python render_readme_assurance.py --write` and commit the result")
        diff = difflib.unified_diff(current.splitlines(), expected.splitlines(),
                                    "README.md (committed)", "README.md (rendered)", lineterm="")
        for line in list(diff)[:60]:
            print("  " + line)
        status = 1
    for problem in model["stale"]:
        print(f"BLOCKING: stale evidence - {problem}")
        status = 1
    if status == 0:
        print("PASS: README assurance block matches committed evidence")
    return status


if __name__ == "__main__":
    sys.exit(main())
