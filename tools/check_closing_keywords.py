#!/usr/bin/env python3
"""Find closing keywords in commit messages before they reach the default branch.

GitHub closes an issue when a commit whose message says "closes #N", "fixes
#N" or "resolves #N" lands on the default branch. A commit made on a release
branch carries that keyword until the branch is integrated, and then closes
the issue whatever its state. Commit messages therefore reference issues with
`Refs #N` only; I close an issue by hand, with its Resolution.

The check reads the messages of the commits in BASE..HEAD. A closing keyword
followed by an issue reference is a finding unless the commit is recorded in
.github/closing-keyword-exceptions.json, with the issues it names and what is
done about them. History is never rewritten to remove a keyword: the commit
is recorded, and the issues are verified after the integration.

Usage:
    python tools/check_closing_keywords.py --base origin/main --head HEAD
    python tools/check_closing_keywords.py --base origin/main --head HEAD --list
    python tools/check_closing_keywords.py --self-test

--list prints the recorded and unrecorded commits as a Markdown table, for
the body of an integration pull request. Without --base there is no range and
the check passes without reading anything.
"""

from __future__ import annotations

import argparse
import contextlib
import io
import json
import re
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path

EXCEPTIONS_PATH = ".github/closing-keyword-exceptions.json"
# The nine keywords GitHub documents, an optional colon, then one reference:
# #N, owner/repo#N or an issue or pull request URL.
REFERENCE = (r"(?:[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+)?#(\d+)"
             r"|https://github\.com/[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+/(?:issues|pull)/(\d+)")
KEYWORD = re.compile(
    r"(?<![A-Za-z0-9_-])(close[sd]?|fix(?:e[sd])?|resolve[sd]?):?[ \t]+(?:" + REFERENCE + ")",
    re.IGNORECASE)
# References that continue the phrase: "resolve #163, #173" and "#163 and #173".
CONTINUATION = re.compile(r"(?:[ \t]*,[ \t]*|[ \t]+and[ \t]+)(?:" + REFERENCE + ")", re.IGNORECASE)
SHA = re.compile(r"[0-9a-f]{40}")


@dataclass(frozen=True)
class Hit:
    commit: str
    subject: str
    phrases: tuple[str, ...]
    issues: tuple[int, ...]


def closing_references(message: str) -> tuple[tuple[str, ...], tuple[int, ...]]:
    """The closing phrases of a commit message and every issue they name."""
    phrases: list[str] = []
    issues: list[int] = []
    for match in KEYWORD.finditer(message):
        end = match.end()
        numbers = [int(match.group(2) or match.group(3))]
        while True:
            more = CONTINUATION.match(message, end)
            if not more:
                break
            numbers.append(int(more.group(1) or more.group(2)))
            end = more.end()
        phrases.append(message[match.start():end])
        issues.extend(n for n in numbers if n not in issues)
    return tuple(phrases), tuple(issues)


def git(root: Path, *args: str) -> str:
    done = subprocess.run(["git", "-C", str(root), *args], stdout=subprocess.PIPE,
                          stderr=subprocess.PIPE, text=True)
    if done.returncode:
        raise RuntimeError(f"git {' '.join(args)}: {done.stderr.strip()}")
    return done.stdout


def hits_in_range(root: Path, base: str, head: str) -> list[Hit]:
    out = git(root, "log", "--format=%H%x00%s%x00%B%x01", f"{base}..{head}")
    hits = []
    for record in out.split("\x01"):
        record = record.strip("\n")
        if not record:
            continue
        commit, subject, body = record.split("\x00", 2)
        phrases, issues = closing_references(body)
        if phrases:
            hits.append(Hit(commit, subject, phrases, issues))
    return hits


def load_exceptions(root: Path) -> dict[str, dict[str, object]]:
    path = root / EXCEPTIONS_PATH
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict) or set(data) != {"schema_version", "exceptions"} \
            or data["schema_version"] != 1 or not isinstance(data["exceptions"], list):
        raise ValueError(f"{EXCEPTIONS_PATH}: expected schema_version 1 and an exceptions array")
    table: dict[str, dict[str, object]] = {}
    for item in data["exceptions"]:
        if not isinstance(item, dict) or set(item) != {"commit", "issues", "reason", "action"}:
            raise ValueError(f"{EXCEPTIONS_PATH}: an exception needs commit, issues, reason and action")
        commit, issues = item["commit"], item["issues"]
        if not isinstance(commit, str) or not SHA.fullmatch(commit):
            raise ValueError(f"{EXCEPTIONS_PATH}: commit must be a full lowercase SHA")
        if not isinstance(issues, list) or not issues \
                or any(type(n) is not int or n < 1 for n in issues):
            raise ValueError(f"{EXCEPTIONS_PATH}: {commit[:7]} needs a list of issue numbers")
        for field in ("reason", "action"):
            if not isinstance(item[field], str) or not item[field].strip():
                raise ValueError(f"{EXCEPTIONS_PATH}: {commit[:7]} needs a non-empty {field}")
        if commit in table:
            raise ValueError(f"{EXCEPTIONS_PATH}: {commit[:7]} is recorded twice")
        table[commit] = item
    return table


def findings(hits: list[Hit], exceptions: dict[str, dict[str, object]]) -> list[str]:
    problems = []
    for hit in hits:
        recorded = exceptions.get(hit.commit)
        named = ", ".join(f"#{n}" for n in hit.issues)
        if recorded is None:
            problems.append(f"{hit.commit[:7]} \"{hit.subject}\": closing keyword for {named} "
                            f"({'; '.join(hit.phrases)}). Use `Refs #N` in commit messages; "
                            f"a commit that cannot change is recorded in {EXCEPTIONS_PATH}")
        elif sorted(recorded["issues"]) != sorted(hit.issues):  # type: ignore[call-overload]
            problems.append(f"{hit.commit[:7]}: the exception records {recorded['issues']} "
                            f"but the message names {list(hit.issues)}")
    return problems


def table(hits: list[Hit], exceptions: dict[str, dict[str, object]]) -> str:
    if not hits:
        return "No commit in the range uses a closing keyword for an issue."
    lines = ["| Commit | Closing phrase | Issues | Recorded action |", "| --- | --- | --- | --- |"]
    for hit in hits:
        recorded = exceptions.get(hit.commit)
        action = str(recorded["action"]) if recorded else "**not recorded**"
        lines.append(f"| `{hit.commit[:7]}` | {'; '.join(hit.phrases)} | "
                     f"{', '.join(f'#{n}' for n in hit.issues)} | {action} |")
    return "\n".join(lines)


def run(root: Path, base: str | None, head: str, listing: bool) -> int:
    if not base:
        print("closing keywords: no base given, so there is no range to read")
        return 0
    try:
        hits = hits_in_range(root, base, head)
        exceptions = load_exceptions(root)
    except (RuntimeError, ValueError) as error:
        print(f"FAIL: {error}")
        return 1
    if listing:
        print(table(hits, exceptions))
    problems = findings(hits, exceptions)
    for problem in problems:
        print(f"FAIL: {problem}")
    if problems:
        return 1
    if not listing:
        recorded = sum(1 for hit in hits if hit.commit in exceptions)
        print(f"PASS: no unrecorded closing keyword in {base}..{head} "
              f"({recorded} recorded exception(s) in the range)")
    return 0


def self_test() -> int:
    failures = []

    def expect(condition: bool, text: str) -> None:
        if not condition:
            failures.append(text)

    cases = {
        "Resolve #13": (13,),
        "this resolves #163 and #173": (163, 173),
        "resolve #163, #173": (163, 173),
        "Fixes: #7": (7,),
        "closed owner/repo#12": (12,),
        "Fixed https://github.com/o/r/issues/44": (44,),
        "CLOSES #5": (5,),
    }
    for message, issues in cases.items():
        expect(closing_references(message)[1] == issues, f"{message!r} must name {issues}")
    for message in ("Refs #13", "Related to #14", "prefixes #3", "fix the parser (#12)",
                    "resolved in #", "closes the gap, see #9", "unresolved #4"):
        expect(closing_references(message)[1] == (), f"{message!r} is not a closing reference")

    # The fixtures print their own verdicts; only the self-test's is shown.
    with tempfile.TemporaryDirectory() as folder, contextlib.redirect_stdout(io.StringIO()):
        root = Path(folder)

        def commit(message: str) -> str:
            git(root, "-c", "user.name=t", "-c", "user.email=t@example.invalid",
                "commit", "-q", "--allow-empty", "-m", message)
            return git(root, "rev-parse", "HEAD").strip()

        git(root, "init", "-q")
        base = commit("start")
        commit("docs: a rule\n\nRefs #1")
        expect(run(root, base, "HEAD", False) == 0, "Refs is accepted")
        bad = commit("fix: a thing\n\nResolves #2 and #3")
        expect(run(root, base, "HEAD", False) == 1, "an unrecorded closing keyword fails")
        (root / ".github").mkdir()
        record = {"commit": bad, "issues": [2], "reason": "pushed", "action": "verify #2"}
        path = root / EXCEPTIONS_PATH
        path.write_text(json.dumps({"schema_version": 1, "exceptions": [record]}))
        expect(run(root, base, "HEAD", False) == 1, "an exception naming other issues fails")
        record["issues"] = [3, 2]
        path.write_text(json.dumps({"schema_version": 1, "exceptions": [record]}))
        expect(run(root, base, "HEAD", False) == 0, "a recorded commit passes")
        expect(run(root, base, "HEAD", True) == 0, "the listing of a recorded range passes")
        record["action"] = " "
        path.write_text(json.dumps({"schema_version": 1, "exceptions": [record]}))
        expect(run(root, base, "HEAD", False) == 1, "an exception without an action fails")
        expect(run(root, None, "HEAD", False) == 0, "no base means no range")
        expect(run(root, "missing-ref", "HEAD", False) == 1, "an unknown base fails")

    if failures:
        for text in failures:
            print(f"SELF-TEST FAIL: {text}")
        return 1
    print("SELF-TEST PASS: the nine keywords, colons, lists, cross-repository and URL "
          "references, words that only contain a keyword, unrecorded commits, exceptions "
          "with wrong issues or no action, and missing or unknown bases")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--root", default=".")
    parser.add_argument("--base", default=None)
    parser.add_argument("--head", default="HEAD")
    parser.add_argument("--list", action="store_true", dest="listing")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args(argv)
    if args.self_test:
        return self_test()
    return run(Path(args.root), args.base, args.head, args.listing)


if __name__ == "__main__":
    sys.exit(main())
