#!/usr/bin/env python3
"""Validate strict SemVer, changelog, and canonical release-history semantics."""

from __future__ import annotations

import json
import re
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

from _gatelib import git_text, run_gate
from _gatelib import parse_report_args as parse_args

CONFIG_PATH = ".github/repository-profile.json"
HISTORY_POLICY_PATH = ".github/release-history-policy.json"
SEMVER_RE = re.compile(
    r"^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)"
    r"(?:-([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?"
    r"(?:\+([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?$"
)
HEADING_RE = re.compile(r"^## \[([^\]]+)\] - (\d{4}-\d{2}-\d{2})\s*$")
LINK_RE = re.compile(r"^\[([^\]]+)\]:\s*(\S+)\s*$")
SHA_RE = re.compile(r"[0-9a-f]{40}")
TAG_RE = re.compile(r"v(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:[-+][0-9A-Za-z.-]+)?")
REVIEW_RE = re.compile(
    r"https://github\.com/[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+/(?:issues|pull)/[1-9]\d*"
)
SYNTHETIC_PR_MERGE_RE = re.compile(r"Merge ([0-9a-f]{40}) into ([0-9a-f]{40})")
TOOL_NAME = "Release semantics"
CHANGELOG_DATE_SEMANTICS = "release-section-cut-freeze-date"
HISTORY_FINDINGS = frozenset({"merge-commit", "duplicate-subject"})


@dataclass(frozen=True)
class SemVer:
    text: str
    major: int
    minor: int
    patch: int
    prerelease: tuple[str, ...]
    build: tuple[str, ...]


def parse_semver(value: str) -> SemVer:
    match = SEMVER_RE.fullmatch(value)
    if not match:
        raise ValueError(f"invalid SemVer: {value!r}")
    prerelease = tuple(match.group(4).split(".")) if match.group(4) else ()
    build = tuple(match.group(5).split(".")) if match.group(5) else ()
    for identifier in prerelease:
        if identifier.isdigit() and len(identifier) > 1 and identifier.startswith("0"):
            raise ValueError(
                f"numeric pre-release identifier must not contain leading zeros: {identifier!r}"
            )
    return SemVer(
        value,
        int(match.group(1)),
        int(match.group(2)),
        int(match.group(3)),
        prerelease,
        build,
    )


def compare(left: SemVer, right: SemVer) -> int:
    core_left = (left.major, left.minor, left.patch)
    core_right = (right.major, right.minor, right.patch)
    if core_left != core_right:
        return 1 if core_left > core_right else -1
    if not left.prerelease and not right.prerelease:
        return 0
    if not left.prerelease:
        return 1
    if not right.prerelease:
        return -1
    for left_item, right_item in zip(left.prerelease, right.prerelease):
        if left_item == right_item:
            continue
        left_numeric = left_item.isdigit()
        right_numeric = right_item.isdigit()
        if left_numeric and right_numeric:
            return 1 if int(left_item) > int(right_item) else -1
        if left_numeric != right_numeric:
            return -1 if left_numeric else 1
        return 1 if left_item > right_item else -1
    if len(left.prerelease) == len(right.prerelease):
        return 0
    return 1 if len(left.prerelease) > len(right.prerelease) else -1


def valid_date(value: str) -> bool:
    try:
        date.fromisoformat(value)
    except ValueError:
        return False
    return True


def _parse_releases(
    changelog: str, findings: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    releases: list[dict[str, Any]] = []
    seen: dict[str, int] = {}
    for number, raw in enumerate(changelog.splitlines(), start=1):
        match = HEADING_RE.match(raw)
        if not match:
            continue
        value, date_text = match.groups()
        try:
            parsed = parse_semver(value)
        except ValueError as error:
            findings.append(
                {"path": "CHANGELOG.md", "line": number, "message": str(error)}
            )
            continue
        if not valid_date(date_text):
            findings.append(
                {
                    "path": "CHANGELOG.md",
                    "line": number,
                    "message": f"release date is not a real Gregorian date: {date_text}",
                }
            )
        if value in seen:
            findings.append(
                {
                    "path": "CHANGELOG.md",
                    "line": number,
                    "message": (
                        f"duplicate release version {value!r}; first declared at line {seen[value]}"
                    ),
                }
            )
        else:
            seen[value] = number
        releases.append(
            {"version": value, "semver": parsed, "date": date_text, "line": number}
        )
    return releases


def _validate_release_order(
    releases: list[dict[str, Any]],
    version: str,
    version_semver: SemVer | None,
    findings: list[dict[str, Any]],
) -> None:
    for current, older in zip(releases, releases[1:]):
        if compare(current["semver"], older["semver"]) > 0:
            continue
        findings.append(
            {
                "path": "CHANGELOG.md",
                "line": current["line"],
                "message": (
                    "released versions must be strictly descending by SemVer precedence; "
                    f"{current['version']} is not newer than {older['version']}"
                ),
            }
        )
    if (
        releases
        and version_semver is not None
        and version != "0.0.0"
        and releases[0]["version"] != version
    ):
        findings.append(
            {
                "path": "VERSION",
                "message": (
                    f"VERSION {version!r} must match the newest dated changelog release "
                    f"{releases[0]['version']!r}."
                ),
            }
        )


def _validate_release_dates(
    releases: list[dict[str, Any]], findings: list[dict[str, Any]]
) -> None:
    for current, older in zip(releases, releases[1:]):
        if not valid_date(current["date"]) or not valid_date(older["date"]):
            continue
        current_date = date.fromisoformat(current["date"])
        older_date = date.fromisoformat(older["date"])
        if current_date >= older_date:
            continue
        findings.append(
            {
                "path": "CHANGELOG.md",
                "line": current["line"],
                "message": (
                    "release-section cut/freeze dates must not move backward across "
                    "newest-to-oldest release headings; "
                    f"{current['version']} uses {current['date']} but older "
                    f"{older['version']} uses {older['date']}"
                ),
            }
        )


def _parse_links(
    changelog: str, findings: list[dict[str, Any]]
) -> dict[str, tuple[str, int]]:
    links: dict[str, tuple[str, int]] = {}
    for number, raw in enumerate(changelog.splitlines(), start=1):
        match = LINK_RE.match(raw)
        if not match:
            continue
        name, url = match.groups()
        if name in links:
            findings.append(
                {
                    "path": "CHANGELOG.md",
                    "line": number,
                    "message": f"duplicate comparison-link definition for [{name}]",
                }
            )
        else:
            links[name] = (url, number)
    return links


def _validate_expected_link(
    links: dict[str, tuple[str, int]],
    name: str,
    expected: str,
    missing_message: str,
    findings: list[dict[str, Any]],
) -> None:
    actual = links.get(name)
    if actual is None:
        findings.append({"path": "CHANGELOG.md", "message": missing_message})
    elif actual[0] != expected:
        findings.append(
            {
                "path": "CHANGELOG.md",
                "line": actual[1],
                "message": f"[{name}] link must be {expected}; observed {actual[0]}",
            }
        )


def _validate_links(
    releases: list[dict[str, Any]],
    links: dict[str, tuple[str, int]],
    repository: str,
    findings: list[dict[str, Any]],
) -> None:
    if not releases:
        return
    latest = releases[0]["version"]
    expected_unreleased = f"https://github.com/{repository}/compare/v{latest}...HEAD"
    _validate_expected_link(
        links,
        "Unreleased",
        expected_unreleased,
        f"missing [Unreleased] comparison link; expected {expected_unreleased}",
        findings,
    )
    for index, release in enumerate(releases):
        value = release["version"]
        if index + 1 < len(releases):
            older = releases[index + 1]["version"]
            expected = f"https://github.com/{repository}/compare/v{older}...v{value}"
        else:
            expected = f"https://github.com/{repository}/releases/tag/v{value}"
        _validate_expected_link(
            links,
            value,
            expected,
            f"missing [{value}] release comparison link; expected {expected}",
            findings,
        )


def _git_output(root: Path, *arguments: str) -> str:
    completed = git_text(root, *arguments)
    if completed.returncode:
        detail = completed.stderr.strip() or "git command failed"
        raise ValueError(detail)
    return completed.stdout.strip()


def _candidate_json(root: Path, candidate: str, path: str) -> dict[str, Any]:
    raw = _git_output(root, "show", f"{candidate}:{path}")
    value = json.loads(raw)
    if not isinstance(value, dict):
        raise ValueError(f"{path}: root must be an object")
    return value


def _history_exception(value: object) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ValueError("release-history exception must be an object")
    expected = {"base_tag", "commit", "findings", "review_ref", "reason"}
    if set(value) != expected:
        raise ValueError("release-history exception has invalid fields")
    base_tag = value["base_tag"]
    commit = value["commit"]
    kinds = value["findings"]
    review_ref = value["review_ref"]
    reason = value["reason"]
    if not isinstance(base_tag, str) or TAG_RE.fullmatch(base_tag) is None:
        raise ValueError("release-history exception base_tag is invalid")
    if not isinstance(commit, str) or SHA_RE.fullmatch(commit) is None:
        raise ValueError("release-history exception commit is invalid")
    if (
        not isinstance(kinds, list)
        or not kinds
        or any(item not in HISTORY_FINDINGS for item in kinds)
        or len(kinds) != len(set(kinds))
    ):
        raise ValueError("release-history exception findings are invalid")
    if not isinstance(review_ref, str) or REVIEW_RE.fullmatch(review_ref) is None:
        raise ValueError("release-history exception review_ref must be a GitHub issue or PR URL")
    if not isinstance(reason, str) or not reason.strip():
        raise ValueError("release-history exception reason must be non-empty")
    return value


def _historical_record(value: object) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ValueError("historical release-history record must be an object")
    expected = {"release", "commit", "pull_request", "reason"}
    if set(value) != expected:
        raise ValueError("historical release-history record has invalid fields")
    release = value["release"]
    commit = value["commit"]
    pull_request = value["pull_request"]
    reason = value["reason"]
    if not isinstance(release, str) or TAG_RE.fullmatch(release) is None:
        raise ValueError("historical release tag is invalid")
    if not isinstance(commit, str) or SHA_RE.fullmatch(commit) is None:
        raise ValueError("historical release commit is invalid")
    if not isinstance(pull_request, int) or isinstance(pull_request, bool) or pull_request <= 0:
        raise ValueError("historical pull_request must be a positive integer")
    if not isinstance(reason, str) or not reason.strip():
        raise ValueError("historical release reason must be non-empty")
    return value


def _load_history_policy(root: Path, candidate: str) -> dict[str, Any]:
    policy = _candidate_json(root, candidate, HISTORY_POLICY_PATH)
    expected = {"schema_version", "rules", "exceptions", "historical_records"}
    if set(policy) != expected:
        raise ValueError(f"{HISTORY_POLICY_PATH}: invalid top-level fields")
    if policy.get("schema_version") != 1:
        raise ValueError(f"{HISTORY_POLICY_PATH}: unsupported schema version")
    rules = policy.get("rules")
    expected_rules = {
        "merge_commits": "block-unless-excepted",
        "duplicate_subjects": "block-unless-excepted",
    }
    if rules != expected_rules:
        raise ValueError(f"{HISTORY_POLICY_PATH}: unsupported rule set")
    exceptions = policy.get("exceptions")
    historical = policy.get("historical_records")
    if not isinstance(exceptions, list):
        raise ValueError(f"{HISTORY_POLICY_PATH}: exceptions must be an array")
    if not isinstance(historical, list):
        raise ValueError(f"{HISTORY_POLICY_PATH}: historical_records must be an array")
    policy["exceptions"] = [_history_exception(item) for item in exceptions]
    policy["historical_records"] = [_historical_record(item) for item in historical]
    return policy


def _effective_history_candidate(root: Path, checkout_sha: str) -> str:
    parents = _git_output(root, "show", "-s", "--format=%P", checkout_sha).split()
    if len(parents) != 2:
        return checkout_sha
    subject = _git_output(root, "show", "-s", "--format=%s", checkout_sha)
    match = SYNTHETIC_PR_MERGE_RE.fullmatch(subject)
    if match is None:
        return checkout_sha
    head_sha, base_sha = match.groups()
    if parents != [base_sha, head_sha]:
        return checkout_sha
    return head_sha


def _previous_release_tag(root: Path, candidate: str) -> str:
    parent = _git_output(root, "rev-parse", f"{candidate}^")
    tag = _git_output(
        root, "describe", "--tags", "--abbrev=0", "--match", "v[0-9]*", parent
    )
    if TAG_RE.fullmatch(tag) is None:
        raise ValueError(f"previous release tag is not canonical SemVer: {tag!r}")
    return tag


def _history_rows(root: Path, base_tag: str, candidate: str) -> list[dict[str, Any]]:
    raw = _git_output(
        root,
        "log",
        "--reverse",
        "--format=%H%x1f%P%x1f%s%x1e",
        f"{base_tag}..{candidate}",
    )
    rows: list[dict[str, Any]] = []
    for record in raw.split("\x1e"):
        if not record.strip():
            continue
        fields = record.strip().split("\x1f")
        if len(fields) != 3:
            raise ValueError("could not parse Git history record")
        sha, parents, subject = fields
        rows.append(
            {
                "commit": sha,
                "parents": parents.split() if parents else [],
                "subject": subject.strip(),
            }
        )
    return rows


def _history_conditions(rows: list[dict[str, Any]]) -> dict[str, set[str]]:
    observed: dict[str, set[str]] = {}
    subjects: dict[str, str] = {}
    for row in rows:
        commit = row["commit"]
        if len(row["parents"]) > 1:
            observed.setdefault(commit, set()).add("merge-commit")
        normalized = row["subject"].strip().casefold()
        if not normalized:
            continue
        if normalized in subjects:
            observed.setdefault(commit, set()).add("duplicate-subject")
        else:
            subjects[normalized] = commit
    return observed


def _history_findings(
    rows: list[dict[str, Any]], policy: dict[str, Any], base_tag: str
) -> tuple[list[dict[str, Any]], list[dict[str, str]]]:
    findings: list[dict[str, Any]] = []
    active: dict[str, set[str]] = {}
    for item in policy["exceptions"]:
        if item["base_tag"] != base_tag:
            findings.append(
                {
                    "code": "superseded-history-exception-base",
                    "path": item["commit"],
                    "message": (
                        f"exception records base tag {item['base_tag']!r} but current "
                        f"previous-release base tag is {base_tag!r}; remove the active "
                        "exception or preserve durable historical evidence under "
                        "historical_records"
                    ),
                }
            )
            continue
        active.setdefault(item["commit"], set()).update(item["findings"])
    observed = _history_conditions(rows)
    by_sha = {row["commit"]: row for row in rows}
    used: list[dict[str, str]] = []
    for commit, conditions in sorted(observed.items()):
        allowed = active.get(commit, set())
        for condition in sorted(conditions):
            if condition in allowed:
                used.append({"commit": commit, "finding": condition})
                continue
            row = by_sha[commit]
            if condition == "merge-commit":
                message = "multi-parent commit is not covered by a reviewed history exception"
                code = "unapproved-merge-commit"
            else:
                message = f"duplicate commit subject is not excepted: {row['subject']!r}"
                code = "duplicate-commit-subject"
            findings.append({"code": code, "path": commit, "message": message})
    for commit, allowed in sorted(active.items()):
        actual = observed.get(commit, set())
        if commit not in by_sha:
            findings.append(
                {
                    "code": "stale-history-exception",
                    "path": commit,
                    "message": "active exception commit is not in the inspected release range",
                }
            )
            continue
        for condition in sorted(allowed - actual):
            findings.append(
                {
                    "code": "overbroad-history-exception",
                    "path": commit,
                    "message": f"exception permits {condition!r} but that condition is not present",
                }
            )
    return findings, used


def _release_history_report(root: Path, config: dict[str, Any]) -> dict[str, Any]:
    checkout_sha = _git_output(root, "rev-parse", "HEAD")
    if config.get("mode") != "template" or config.get("profile") is not None:
        return {
            "applicable": False,
            "checkout_sha": checkout_sha,
            "candidate_sha": checkout_sha,
            "previous_tag": None,
            "commits": [],
            "exceptions_used": [],
            "historical_records": [],
            "findings": [],
        }
    candidate = _effective_history_candidate(root, checkout_sha)
    policy = _load_history_policy(root, candidate)
    base_tag = _previous_release_tag(root, candidate)
    rows = _history_rows(root, base_tag, candidate)
    findings, used = _history_findings(rows, policy, base_tag)
    return {
        "applicable": True,
        "checkout_sha": checkout_sha,
        "candidate_sha": candidate,
        "previous_tag": base_tag,
        "commits": rows,
        "exceptions_used": used,
        "historical_records": policy["historical_records"],
        "findings": findings,
    }


def analyze(version: str, changelog: str, repository: str) -> dict[str, Any]:
    findings: list[dict[str, Any]] = []
    try:
        version_semver = parse_semver(version)
    except ValueError as error:
        version_semver = None
        findings.append({"path": "VERSION", "message": str(error)})

    unreleased_count = sum(
        raw.strip() == "## [Unreleased]" for raw in changelog.splitlines()
    )
    if unreleased_count != 1:
        findings.append(
            {
                "path": "CHANGELOG.md",
                "message": (
                    "Changelog requires exactly one [Unreleased] heading; "
                    f"observed {unreleased_count}."
                ),
            }
        )

    releases = _parse_releases(changelog, findings)
    _validate_release_order(releases, version, version_semver, findings)
    _validate_release_dates(releases, findings)
    links = _parse_links(changelog, findings)
    _validate_links(releases, links, repository, findings)

    release_evidence = [
        {"version": item["version"], "date": item["date"], "line": item["line"]}
        for item in releases
    ]
    return {
        "schema_version": 1,
        "tool": TOOL_NAME,
        "status": "pass" if not findings else "fail",
        "version": version,
        "repository": repository,
        "releases": release_evidence,
        "date_policy": {
            "semantic": CHANGELOG_DATE_SEMANTICS,
            "ordering": "newer-release-cut-date>=older-release-cut-date",
            "tag_publication_relation": "independent",
        },
        "link_policy": {
            "unreleased": "latest-tag...HEAD",
            "initial_release": "release-tag",
            "later_release": "preceding-tag...release-tag",
        },
        "findings": findings,
    }


def run_check(root: Path) -> dict[str, Any]:
    root = root.resolve()
    version = (root / "VERSION").read_text(encoding="utf-8").strip()
    changelog = (root / "CHANGELOG.md").read_text(encoding="utf-8")
    config = json.loads((root / CONFIG_PATH).read_text(encoding="utf-8"))
    if not isinstance(config, dict):
        raise ValueError(f"{CONFIG_PATH}: root must be an object")
    report = analyze(version, changelog, config["repository"])
    history = _release_history_report(root, config)
    report["history_policy"] = history
    report["findings"].extend(history["findings"])
    report["status"] = "pass" if not report["findings"] else "fail"
    return report


def markdown_report(report: dict[str, Any]) -> str:
    history = report.get("history_policy")
    lines = [
        "## Release semantics",
        "",
        f"- **Status:** {str(report['status']).upper()}",
        f"- **VERSION:** `{report['version']}`",
        "- **Changelog date semantic:** release-section cut/freeze date",
        f"- **Released headings:** {len(report['releases'])}",
        f"- **Findings:** {len(report['findings'])}",
    ]
    if history is not None:
        lines.extend(
            [
                f"- **Template history policy:** {'applicable' if history['applicable'] else 'not applicable'}",
                f"- **History range base:** `{history['previous_tag'] or 'not applicable'}`",
                f"- **History candidate:** `{history['candidate_sha']}`",
                f"- **History commits inspected:** {len(history['commits'])}",
                f"- **History exceptions used:** {len(history['exceptions_used'])}",
            ]
        )
        if history["checkout_sha"] != history["candidate_sha"]:
            lines.append(
                f"- **Synthetic PR merge checkout:** `{history['checkout_sha']}` resolved to PR head"
            )
    if report["releases"]:
        lines.extend(["", "| Version | Cut/freeze date | Line |", "| --- | --- | ---: |"])
        for item in report["releases"]:
            lines.append(f"| `{item['version']}` | {item['date']} | {item['line']} |")
    if report["findings"]:
        lines.extend(["", "### Findings", ""])
        for item in report["findings"]:
            location = item.get("path", ".")
            if item.get("line"):
                location += f":{item['line']}"
            prefix = f"`{item['code']}` — " if item.get("code") else ""
            lines.append(f"- `{location}` — {prefix}{item['message']}")
    return "\n".join(lines) + "\n"


def fixture(
    version: str, headings: list[tuple[str, str]], links: dict[str, str]
) -> dict[str, Any]:
    repository = "example/repo"
    lines = ["# Changelog", "", "## [Unreleased]", "", "No unreleased changes.", ""]
    for value, date_text in headings:
        lines.extend([f"## [{value}] - {date_text}", "", "- Change.", ""])
    for name, url in links.items():
        lines.append(f"[{name}]: {url}")
    return analyze(version, "\n".join(lines) + "\n", repository)


def canonical_links(versions: list[str]) -> dict[str, str]:
    repository = "example/repo"
    result = {
        "Unreleased": f"https://github.com/{repository}/compare/v{versions[0]}...HEAD"
    }
    for index, value in enumerate(versions):
        if index + 1 < len(versions):
            result[value] = (
                f"https://github.com/{repository}/compare/v{versions[index + 1]}...v{value}"
            )
        else:
            result[value] = f"https://github.com/{repository}/releases/tag/v{value}"
    return result


def _fixture_git(root: Path, *arguments: str) -> str:
    completed = subprocess.run(
        [
            "git",
            "-C",
            str(root),
            "-c",
            "user.name=Release History Fixture",
            "-c",
            "user.email=history@example.invalid",
            *arguments,
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    return completed.stdout.strip()


def _history_base_fixture(root: Path) -> None:
    (root / ".github").mkdir(parents=True)
    (root / CONFIG_PATH).write_text(
        json.dumps({"mode": "template", "profile": None, "repository": "example/repo"}) + "\n",
        encoding="utf-8",
    )
    policy = {
        "schema_version": 1,
        "rules": {
            "merge_commits": "block-unless-excepted",
            "duplicate_subjects": "block-unless-excepted",
        },
        "exceptions": [],
        "historical_records": [],
    }
    (root / HISTORY_POLICY_PATH).write_text(json.dumps(policy, indent=2) + "\n", encoding="utf-8")
    (root / "fixture.txt").write_text("base\n", encoding="utf-8")
    _fixture_git(root, "init", "-b", "main")
    _fixture_git(root, "add", "--all")
    _fixture_git(root, "commit", "-m", "Base release")
    _fixture_git(root, "tag", "-a", "v1.0.0", "-m", "v1.0.0")


def _history_commit(root: Path, filename: str, text: str, subject: str) -> str:
    (root / filename).write_text(text + "\n", encoding="utf-8")
    _fixture_git(root, "add", filename)
    _fixture_git(root, "commit", "-m", subject)
    return _fixture_git(root, "rev-parse", "HEAD")


def _history_merge_case(
    root: Path, approved: bool, exception_base_tag: str = "v1.0.0"
) -> dict[str, Any]:
    _history_base_fixture(root)
    _fixture_git(root, "switch", "-c", "feature")
    _history_commit(root, "feature.txt", "feature", "Feature change")
    _fixture_git(root, "switch", "main")
    _history_commit(root, "main.txt", "main", "Mainline preparation")
    _fixture_git(root, "merge", "--no-ff", "feature", "-m", "Reviewed feature merge")
    merge_sha = _fixture_git(root, "rev-parse", "HEAD")
    if approved:
        policy = json.loads((root / HISTORY_POLICY_PATH).read_text(encoding="utf-8"))
        policy["exceptions"] = [
            {
                "base_tag": exception_base_tag,
                "commit": merge_sha,
                "findings": ["merge-commit"],
                "review_ref": "https://github.com/example/repo/pull/1",
                "reason": "Fixture-reviewed ancestry-preserving merge.",
            }
        ]
        (root / HISTORY_POLICY_PATH).write_text(
            json.dumps(policy, indent=2) + "\n", encoding="utf-8"
        )
        _fixture_git(root, "add", HISTORY_POLICY_PATH)
        _fixture_git(root, "commit", "-m", "Register reviewed merge exception")
    config = json.loads((root / CONFIG_PATH).read_text(encoding="utf-8"))
    return _release_history_report(root, config)


def _history_self_test_cases() -> list[tuple[str, bool]]:
    results: list[tuple[str, bool]] = []
    with tempfile.TemporaryDirectory(prefix="release-history-") as temporary:
        area = Path(temporary)

        compliant = area / "compliant"
        compliant.mkdir()
        _history_base_fixture(compliant)
        _history_commit(compliant, "focused.txt", "one", "Focused squash result")
        report = _release_history_report(
            compliant, json.loads((compliant / CONFIG_PATH).read_text(encoding="utf-8"))
        )
        results.append(("compliant-squash-history", not report["findings"]))

        approved = area / "approved"
        approved.mkdir()
        report = _history_merge_case(approved, approved=True)
        results.append(
            (
                "approved-merge-exception",
                not report["findings"] and len(report["exceptions_used"]) == 1,
            )
        )

        superseded = area / "superseded-base"
        superseded.mkdir()
        report = _history_merge_case(
            superseded, approved=True, exception_base_tag="v0.9.0"
        )
        stale_base = [
            item
            for item in report["findings"]
            if item.get("code") == "superseded-history-exception-base"
        ]
        results.append(
            (
                "superseded-base-exception",
                len(stale_base) == 1
                and SHA_RE.fullmatch(str(stale_base[0].get("path", ""))) is not None
                and "v0.9.0" in stale_base[0]["message"]
                and "v1.0.0" in stale_base[0]["message"],
            )
        )

        unapproved = area / "unapproved"
        unapproved.mkdir()
        report = _history_merge_case(unapproved, approved=False)
        codes = {item.get("code") for item in report["findings"]}
        results.append(("unapproved-merge-commit", "unapproved-merge-commit" in codes))

        duplicate = area / "duplicate"
        duplicate.mkdir()
        _history_base_fixture(duplicate)
        _history_commit(duplicate, "one.txt", "one", "Repeated subject")
        _history_commit(duplicate, "two.txt", "two", "Repeated subject")
        report = _release_history_report(
            duplicate, json.loads((duplicate / CONFIG_PATH).read_text(encoding="utf-8"))
        )
        codes = {item.get("code") for item in report["findings"]}
        results.append(("duplicate-subject-history", "duplicate-commit-subject" in codes))
    return results


def run_self_test() -> int:
    cases: list[tuple[str, str, dict[str, Any]]] = []
    stable_versions = ["1.1.0", "1.0.0"]
    cases.append(
        (
            "valid-stable",
            "pass",
            fixture(
                "1.1.0",
                [("1.1.0", "2026-09-05"), ("1.0.0", "2026-09-04")],
                canonical_links(stable_versions),
            ),
        )
    )
    cases.append(
        (
            "valid-same-day-cut-dates",
            "pass",
            fixture(
                "1.1.0",
                [("1.1.0", "2026-09-05"), ("1.0.0", "2026-09-05")],
                canonical_links(stable_versions),
            ),
        )
    )
    pre_versions = ["1.1.0-rc.1", "1.0.0"]
    cases.append(
        (
            "valid-prerelease",
            "pass",
            fixture(
                "1.1.0-rc.1",
                [("1.1.0-rc.1", "2026-09-05"), ("1.0.0", "2026-09-04")],
                canonical_links(pre_versions),
            ),
        )
    )
    cases.append(
        (
            "leading-zero-prerelease",
            "fail",
            fixture(
                "1.1.0-01",
                [("1.1.0-01", "2026-09-05"), ("1.0.0", "2026-09-04")],
                canonical_links(["1.1.0-01", "1.0.0"]),
            ),
        )
    )
    cases.append(
        (
            "out-of-order",
            "fail",
            fixture(
                "1.0.0",
                [("1.0.0", "2026-09-05"), ("1.1.0", "2026-09-04")],
                canonical_links(["1.0.0", "1.1.0"]),
            ),
        )
    )
    cases.append(
        (
            "duplicate-version",
            "fail",
            fixture(
                "1.1.0",
                [("1.1.0", "2026-09-05"), ("1.1.0", "2026-09-04")],
                canonical_links(["1.1.0", "1.1.0"]),
            ),
        )
    )
    cases.append(
        (
            "impossible-date",
            "fail",
            fixture(
                "1.1.0",
                [("1.1.0", "2026-02-30"), ("1.0.0", "2026-09-04")],
                canonical_links(stable_versions),
            ),
        )
    )
    cases.append(
        (
            "cut-date-regression",
            "fail",
            fixture(
                "1.1.0",
                [("1.1.0", "2026-09-04"), ("1.0.0", "2026-09-05")],
                canonical_links(stable_versions),
            ),
        )
    )
    bad_links = canonical_links(stable_versions)
    bad_links["Unreleased"] = "https://github.com/example/repo/compare/v0.9.0...HEAD"
    cases.append(
        (
            "wrong-unreleased-link",
            "fail",
            fixture(
                "1.1.0",
                [("1.1.0", "2026-09-05"), ("1.0.0", "2026-09-04")],
                bad_links,
            ),
        )
    )
    missing_link = canonical_links(stable_versions)
    del missing_link["1.1.0"]
    cases.append(
        (
            "missing-release-link",
            "fail",
            fixture(
                "1.1.0",
                [("1.1.0", "2026-09-05"), ("1.0.0", "2026-09-04")],
                missing_link,
            ),
        )
    )
    cases.append(
        (
            "version-heading-mismatch",
            "fail",
            fixture(
                "1.2.0",
                [("1.1.0", "2026-09-05"), ("1.0.0", "2026-09-04")],
                canonical_links(stable_versions),
            ),
        )
    )

    failures: list[str] = []
    for name, expected, report in cases:
        if report["status"] != expected:
            failures.append(
                f"{name}: expected {expected}, got {report['status']} ({report['findings']})"
            )
        if report["date_policy"]["semantic"] != CHANGELOG_DATE_SEMANTICS:
            failures.append(f"{name}: changelog date semantic is missing or incorrect")
    for name, passed in _history_self_test_cases():
        if not passed:
            failures.append(f"{name}: release-history fixture failed")
    if compare(parse_semver("1.0.0-alpha.2"), parse_semver("1.0.0-alpha.10")) >= 0:
        failures.append("SemVer numeric prerelease precedence is incorrect")
    if compare(parse_semver("1.0.0"), parse_semver("1.0.0-rc.1")) <= 0:
        failures.append("stable release must outrank prerelease")

    if failures:
        for failure in failures:
            print(f"[FAIL] {failure}")
        print(f"SELF-TEST FAIL: {len(failures)} failure(s).")
        return 1
    print(
        "SELF-TEST PASS: SemVer, changelog cut/freeze dates, comparison links, "
        "compliant squash history, reviewed merge exceptions, superseded-base "
        "exception hygiene, unapproved merges, and duplicate-subject rejection passed."
    )
    return 0


def main(argv: list[str] | None = None) -> int:
    options = parse_args(sys.argv[1:] if argv is None else argv)
    return run_gate(
        options,
        build=lambda: run_check(options.root),
        markdown=markdown_report,
        errors=(
            OSError,
            UnicodeError,
            ValueError,
            json.JSONDecodeError,
            subprocess.SubprocessError,
        ),
        self_test=run_self_test,
        self_test_error_prefix="SELF-TEST ERROR",
    )


if __name__ == "__main__":
    raise SystemExit(main())
