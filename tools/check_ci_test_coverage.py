#!/usr/bin/env python3
"""Fail when a tracked test_*.py file is not executed on every pull request.

The checker is deliberately conservative. It credits only a few explicitly
recognized shapes and reports everything else, so an ambiguous step has to be
restructured rather than trusted. Credit comes from:

* a ``run:`` step of a workflow whose ``pull_request`` trigger has no path
  filter, in the default or ``bash`` shell, without ``continue-on-error``, whose
  every line is recognized: blank lines, comments, a top-level ``cd <dir>``,
  plain commands (optionally with redirections and leading assignments), and
  ``python``/``python3`` running ``<file>``, ``-m unittest <file|module>`` or
  ``-m unittest discover``. A single other line (a command list, pipeline,
  subshell, substitution, heredoc, loop, ``exit``, ``set +e``, any other
  ``if`` ...) withholds credit from the whole step. unittest credits only
  modules that define a ``TestCase``, and discovery only inside packages
  without a ``load_tests`` hook;
* an entry of a registered runner list (``RUNNERS``) that such a step invokes,
  with the runner's own filter for its reduced mode applied.

A job or step ``if:`` counts only when it is exactly ``always()``,
``success()``, ``!cancelled()`` or the selection gate below, and a job earns
nothing when a job it needs has any other condition. A step that may rebind
a selection variable (``VAR=``, ``export``, ``declare``, ``read``,
``printf -v``, ``${VAR:=...}``, ``unset`` ...) earns nothing.

A run is *unconditional* when nothing but job success gates it. A run gated by
the accuracy-proof selection is *selected*. Two forms count:
``needs.select.outputs.full == 'true'`` in an ``if:``, or the ``then`` branch of
one ``if [[ "$VAR" == "true" ]]; then`` ... ``fi``. VAR must be bound exactly
to ``${{ needs.select.outputs.full }}``, or the expression itself may appear
inline. The gate is matched on the raw line: the operand must be
double-quoted and compared with ``true``, so ``'$VAR'``, which Bash never
expands, is not mistaken for a read. The ``else`` branch earns nothing. A selected run counts only when
``tools/select_accuracy_checks.py`` does not treat a change to that file as
metadata-only, so editing the file always makes it run. Any other condition
earns nothing. ``EXEMPT`` lists reasoned exceptions and is empty.

Threat model and limits: the guard catches a test file that no CI step runs by
accident, such as a new file nobody listed, a renamed one, or a step that
silently skips it. It does not try to defeat a deliberately misleading
workflow; review does that. It never infers that something runs. A shape
outside the recognized ones is reported rather than credited, so its failure
mode is a false report that makes the author simplify the step, not a false
pass.

Workflows are read with a small standard-library parser for the block-style
YAML subset GitHub workflows use; anything outside it is an error (exit 2).
"""

from __future__ import annotations

import argparse
import ast
import fnmatch
import posixpath
import re
import shlex
import sys
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from _gatelib import run_gate, tracked_files

TOOL_NAME = "Every tracked test file runs in CI"
DEFAULT_BRANCH = "main"
TEST_FILE = re.compile(r"test_[^/]*\.py")
WORKFLOW_DIR = ".github/workflows"
# The job whose `full` output is the accuracy-proof selection (select_accuracy_checks.py).
SELECT_JOB = "select"

# Tracked test files deliberately not run in CI: path -> reason. Keep it empty
# unless a file truly cannot run on a hosted runner; the guard reports an entry
# that is stale (untracked) or redundant (already covered).
EXEMPT: dict[str, str] = {}


@dataclass(frozen=True)
class Runner:
    """A script that executes the command tuples bound to ``variable``."""

    variable: str
    reduced_flag: str | None = None
    # Mirrors the runner's own filter: True when an entry still runs under
    # the reduced flag.
    keeps_in_reduced: Callable[[tuple[str, ...]], bool] = lambda _command: True


def _live_check_only(command: tuple[str, ...]) -> bool:
    return posixpath.basename(command[0]).startswith("check_") and "--self-test" not in command


# Runner lists whose entries execute when CI invokes the runner. A runner that
# CI never invokes (benchmark/refresh_evidence.py is a local orchestration
# script) cannot make a file run in CI, so it is deliberately not listed.
RUNNERS: dict[str, Runner] = {
    "benchmark/test_evidence_tools.py": Runner(
        variable="commands", reduced_flag="--checks-only", keeps_in_reduced=_live_check_only
    ),
}

ALWAYS, SELECTED, NEVER = 0, 1, 2
LEVEL_NAMES = {ALWAYS: "always", SELECTED: "with full proofs"}


class CoverageError(ValueError):
    """The workflows or a runner cannot be read reliably."""


# --------------------------------------------------------------------------
# A strict reader for the block-style YAML subset used by GitHub workflows.
# --------------------------------------------------------------------------

BLOCK_INDICATOR = re.compile(r"[|>][+-]?[1-9]?")


def strip_comment(text: str) -> str:
    quote = ""
    for index, char in enumerate(text):
        if quote:
            if char == quote:
                quote = ""
        elif char in "\"'" and (index == 0 or text[index - 1] in " \t[{,:"):
            quote = char
        elif char == "#" and (index == 0 or text[index - 1] in " \t"):
            return text[:index].rstrip()
    return text.rstrip()


def unquote(text: str) -> str:
    if len(text) >= 2 and text[0] == text[-1] == "'":
        return text[1:-1].replace("''", "'")
    if len(text) >= 2 and text[0] == text[-1] == '"':
        body = text[1:-1]
        return re.sub(r'\\(.)', lambda m: {"n": "\n", "t": "\t"}.get(m.group(1), m.group(1)), body)
    return text


def split_flow(body: str) -> list[str]:
    items: list[str] = []
    current: list[str] = []
    quote = ""
    for char in body:
        if quote:
            current.append(char)
            if char == quote:
                quote = ""
        elif char in "\"'":
            quote = char
            current.append(char)
        elif char in "[]{}":
            raise CoverageError(f"nested flow collection is not supported: {body!r}")
        elif char == ",":
            items.append("".join(current).strip())
            current = []
        else:
            current.append(char)
    if "".join(current).strip():
        items.append("".join(current).strip())
    return items


def scalar(text: str) -> Any:
    text = text.strip()
    if text in ("", "~", "null"):
        return None
    if text[0] in "&*!|>%@`":
        raise CoverageError(f"unsupported YAML construct: {text!r}")
    if text.startswith("["):
        if not text.endswith("]"):
            raise CoverageError(f"unterminated flow sequence: {text!r}")
        return [scalar(item) for item in split_flow(text[1:-1])]
    if text.startswith("{"):
        if not text.endswith("}"):
            raise CoverageError(f"unterminated flow mapping: {text!r}")
        mapping: dict[str, Any] = {}
        for item in split_flow(text[1:-1]):
            key, _, value = item.partition(":")
            mapping[unquote(key.strip())] = scalar(value)
        return mapping
    if text[0] in "\"'" and not (len(text) >= 2 and text[-1] == text[0]):
        raise CoverageError(f"multi-line quoted scalars are not supported: {text!r}")
    return unquote(text)


def split_key(text: str) -> tuple[str, str] | None:
    if text[0] in "\"'":
        end = text.find(text[0], 1)
        if end < 0 or not text[end + 1:].startswith(":"):
            return None
        rest = text[end + 2:]
        if rest and not rest[0].isspace():
            return None
        return unquote(text[: end + 1]), rest.strip()
    if text[0] in "[{":
        return None
    match = re.search(r":(\s|$)", text)
    if match is None:
        return None
    return text[: match.start()].strip(), text[match.end():].strip()


class YamlReader:
    def __init__(self, text: str) -> None:
        self.raw = text.splitlines()
        self.views: dict[int, tuple[int, str]] = {}

    def view(self, index: int) -> tuple[int, str] | None:
        if index in self.views:
            return self.views[index]
        line = self.raw[index]
        stripped = line.lstrip(" ")
        if stripped.startswith("\t"):
            raise CoverageError(f"line {index + 1}: tab indentation")
        content = strip_comment(stripped)
        if not content:
            return None
        return len(line) - len(stripped), content

    def next_content(self, index: int) -> int | None:
        while index < len(self.raw):
            if self.view(index) is not None:
                return index
            index += 1
        return None

    def document(self) -> Any:
        start = self.next_content(0)
        if start is not None and self.view(start) == (0, "---"):
            start = self.next_content(start + 1)
        if start is None:
            return None
        node, end = self.node(start)
        leftover = self.next_content(end)
        if leftover is not None:
            raise CoverageError(f"line {leftover + 1}: unparsed content")
        return node

    def node(self, index: int) -> tuple[Any, int]:
        view = self.view(index)
        assert view is not None
        indent, text = view
        if text == "-" or text.startswith("- "):
            return self.sequence(index, indent)
        return self.mapping(index, indent)

    def child(self, index: int, indent: int, allow_same_indent_sequence: bool) -> tuple[Any, int]:
        nxt = self.next_content(index)
        if nxt is None:
            return None, index
        view = self.view(nxt)
        assert view is not None
        child_indent, text = view
        if child_indent > indent or (
            allow_same_indent_sequence and child_indent == indent
            and (text == "-" or text.startswith("- "))
        ):
            return self.node(nxt)
        return None, index

    def value(self, index: int, indent: int, rest: str, in_mapping: bool) -> tuple[Any, int]:
        """Parse the value that follows a key or a sequence dash on line ``index``."""
        if not rest:
            return self.child(index + 1, indent, allow_same_indent_sequence=in_mapping)
        if BLOCK_INDICATOR.fullmatch(rest):
            return self.block_scalar(index + 1, indent, rest)
        parts = [rest]
        end = index + 1
        while True:
            nxt = self.next_content(end)
            if nxt is None:
                break
            view = self.view(nxt)
            assert view is not None
            if view[0] <= indent:
                break
            parts.append(view[1])
            end = nxt + 1
        return scalar(" ".join(parts)), end

    def block_scalar(self, index: int, parent: int, indicator: str) -> tuple[str, int]:
        lines: list[str] = []
        content_indent: int | None = None
        explicit = re.search(r"[1-9]", indicator)
        if explicit:
            content_indent = parent + int(explicit.group())
        end = index
        while end < len(self.raw):
            line = self.raw[end]
            if line.strip():
                indent = len(line) - len(line.lstrip(" "))
                if indent <= parent:
                    break
                if content_indent is None:
                    content_indent = indent
                if indent < content_indent:
                    raise CoverageError(f"line {end + 1}: block scalar under-indented")
                lines.append(line[content_indent:])
            else:
                lines.append("")
            end += 1
        while lines and lines[-1] == "" and not indicator.endswith("+"):
            lines.pop()
        if indicator.startswith(">"):
            folded: list[str] = []
            for line in lines:
                if folded and line and not line[0].isspace() and folded[-1] and not folded[-1][0].isspace():
                    folded[-1] += " " + line
                else:
                    folded.append(line)
            lines = folded
        text = "\n".join(lines)
        if "-" not in indicator and lines:
            text += "\n"
        return text, end

    def mapping(self, index: int, indent: int) -> tuple[dict[str, Any], int]:
        result: dict[str, Any] = {}
        position = index
        while True:
            line_index = self.next_content(position)
            if line_index is None:
                return result, len(self.raw)
            view = self.view(line_index)
            assert view is not None
            line_indent, text = view
            if line_indent < indent or text == "-" or text.startswith("- "):
                return result, line_index
            if line_indent > indent:
                raise CoverageError(f"line {line_index + 1}: unexpected indentation")
            pair = split_key(text)
            if pair is None:
                raise CoverageError(f"line {line_index + 1}: expected 'key: value'")
            key, rest = pair
            if key in result:
                raise CoverageError(f"line {line_index + 1}: duplicate key {key!r}")
            result[key], position = self.value(line_index, indent, rest, in_mapping=True)

    def sequence(self, index: int, indent: int) -> tuple[list[Any], int]:
        result: list[Any] = []
        position = index
        while True:
            line_index = self.next_content(position)
            if line_index is None:
                return result, len(self.raw)
            view = self.view(line_index)
            assert view is not None
            line_indent, text = view
            if line_indent < indent or not (text == "-" or text.startswith("- ")):
                if line_indent > indent:
                    raise CoverageError(f"line {line_index + 1}: unexpected indentation")
                return result, line_index
            if line_indent > indent:
                raise CoverageError(f"line {line_index + 1}: unexpected indentation")
            rest = text[1:]
            stripped = rest.lstrip(" ")
            item_indent = indent + 1 + len(rest) - len(stripped)
            if stripped and not BLOCK_INDICATOR.fullmatch(stripped) and split_key(stripped):
                self.views[line_index] = (item_indent, stripped)
                item, position = self.mapping(line_index, item_indent)
            else:
                item, position = self.value(line_index, indent, stripped, in_mapping=False)
            result.append(item)


def load_yaml(text: str) -> Any:
    return YamlReader(text).document()


# --------------------------------------------------------------------------
# Workflow semantics.
# --------------------------------------------------------------------------


def runs_on_every_pull_request(triggers: Any) -> tuple[bool, str]:
    """Whether ``on:`` runs the workflow on every pull request to the default branch."""
    if triggers == "pull_request" or (isinstance(triggers, list) and "pull_request" in triggers):
        return True, "every pull request"
    if not isinstance(triggers, dict) or "pull_request" not in triggers:
        return False, "no pull_request trigger"
    config = triggers["pull_request"] or {}
    if not isinstance(config, dict):
        return False, "unreadable pull_request trigger"
    for key in ("paths", "paths-ignore"):
        if key in config:
            return False, f"pull_request is filtered by {key}"
    # GitHub evaluates branch patterns in order, and a later `!pattern` can
    # exclude what an earlier one included. Rather than model that, any
    # negative pattern or branches-ignore withholds credit (conservative).
    if "branches-ignore" in config:
        return False, "pull_request uses branches-ignore, which is not evaluated"
    branches = config.get("branches")
    if branches is not None:
        patterns = [str(p) for p in (branches if isinstance(branches, list) else [branches])]
        if any(p.startswith("!") for p in patterns):
            return False, "pull_request branches use a negative pattern, which is not evaluated"
        if not any(fnmatch.fnmatchcase(DEFAULT_BRANCH, p) for p in patterns):
            return False, f"pull_request branches exclude {DEFAULT_BRANCH}"
    types = config.get("types")
    if types is not None:
        listed = set(types if isinstance(types, list) else [types])
        if not {"opened", "synchronize", "reopened"} <= listed:
            return False, "pull_request activity types omit opened/synchronize/reopened"
    return True, "every pull request"


# The only `if:` expressions credited, on a job or a step. Anything else
# (step outcomes, github.event fields, conjunctions ...) earns no credit:
# whether it holds would depend on other steps, jobs or the event.
SELECTION_TERM = re.compile(r"needs\." + re.escape(SELECT_JOB) + r"\.outputs\.full\s*==\s*'true'")
UNGATED_TERMS = {"always()", "!cancelled()"}


def condition_level(expression: Any) -> tuple[int, bool]:
    """Classify an ``if:`` as (level, runs even when a needed job was skipped)."""
    if expression is None:
        return ALWAYS, False
    text = str(expression).strip()
    if text.startswith("${{") and text.endswith("}}"):
        text = text[3:-2].strip()
    if text in UNGATED_TERMS:
        return ALWAYS, True
    if text == "success()":
        return ALWAYS, False
    if SELECTION_TERM.fullmatch(text):
        return SELECTED, False
    return NEVER, False


def truthy(value: Any) -> bool:
    return value is not None and str(value).strip().lower() not in ("false", "")


@dataclass
class Invocation:
    """One interpreter run: working directory (repository-relative) and argv.

    ``reason`` explains why a NEVER-level invocation earns no credit.
    """

    cwd: str
    argv: list[str]
    level: int
    origin: str
    reason: str = ""


# ``run:`` scripts are not interpreted as bash. Credit needs one of a few
# explicitly recognized line shapes, and a single line outside them withholds
# credit from the whole step: a false report makes the author restructure the
# step, a false pass would hide a test that never runs.
PYTHON_NAMES = {"python", "python3"}
# Tokens that join or group commands: any of them makes a line unrecognized.
CONTROL = {";", "&&", "||", "|", "&", ";;", "|&", "(", ")", "{", "}", "!"}
REDIRECTS = {">", ">>", "<", ">&", "&>", "&>>", ">|"}
HEREDOCS = {"<<", "<<<", "<<-"}
# Commands that can end the script early, change error handling, or start
# shell syntax this reader does not follow.
UNSAFE_HEADS = {
    "exit", "return", "exec", "eval", "source", ".", "trap", "shopt", "alias",
    "function", "unset", "pushd", "popd", "break", "continue", "kill", "wait",
    "builtin", "command", "enable", "ulimit", "coproc", "time", "if", "then",
    "else", "elif", "fi", "for", "while", "until", "do", "done", "case", "esac",
    "select", "[[",
}
SELECTION_OUTPUT = re.compile(r"\$\{\{\s*needs\." + re.escape(SELECT_JOB) + r"\.outputs\.full\s*\}\}")


def logical_lines(script: str) -> list[str]:
    lines: list[str] = []
    pending = ""
    for line in script.splitlines():
        if line.endswith("\\"):
            pending += line[:-1] + " "
            continue
        lines.append(pending + line)
        pending = ""
    if pending:
        lines.append(pending)
    return lines


def comparison_operands(words: list[str]) -> tuple[str, str] | None:
    """Operands of a lone `[ x = y ]`, `[[ x == y ]]` or `test x = y`."""
    if words[:1] == ["test"]:
        words = words[1:]
    elif len(words) >= 2 and (words[0], words[-1]) in (("[", "]"), ("[[", "]]")):
        words = words[1:-1]
    else:
        return None
    if len(words) == 3 and words[1] in ("=", "=="):
        return words[0], words[2]
    return None


# The one recognised gate, matched on the raw line before shell lexing: the
# operand must be a double-quoted variable or selection-output expression, so
# that quoting which stops expansion (`'$VAR'`) can never pass as a read.
GATE_LINE = re.compile(
    r'if (?:\[\[ "(?P<a>OPERAND)" ==? (?:"true"|true) \]\]'
    r'|\[ "(?P<b>OPERAND)" ==? (?:"true"|true) \]); then'
    .replace("OPERAND", r"\$[A-Za-z_]\w*|\$\{[A-Za-z_]\w*\}|\$\{\{\s*needs\."
             + re.escape(SELECT_JOB) + r"\.outputs\.full\s*\}\}"))


def is_selection_gate(tokens: list[str], selection_vars: set[str]) -> bool:
    """`if [[ "$VAR" == "true" ]]; then` with VAR bound exactly to the selection output."""
    if tokens[:1] != ["if"] or tokens[-2:] != [";", "then"]:
        return False
    operands = comparison_operands(tokens[1:-2])
    if operands is None or "true" not in operands:
        return False
    refs = {f"${var}" for var in selection_vars} | {f"${{{var}}}" for var in selection_vars}
    return any(op in refs or SELECTION_OUTPUT.fullmatch(op) for op in operands)


def names_variable(token: str, var: str) -> bool:
    """Whether ``token`` names ``var`` other than as a plain `$VAR`/`${VAR}` read.

    Covers every way a step could rebind the selection variable before the
    gate: `VAR=`, `export`/`declare`/`readonly VAR=`, `read VAR`,
    `printf -v VAR`, `${VAR:=...}`, a nameref `ref=VAR`. Some harmless mentions
    are caught too; that is a false report, never a false pass.
    """
    name = re.escape(var)
    reads_removed = re.sub(r"\$\{" + name + r"\}|\$" + name + r"(?!\w)", "", token)
    return re.search(r"(?<!\w)" + name + r"(?!\w)", reads_removed) is not None


def strip_redirections(tokens: list[str]) -> list[str]:
    words: list[str] = []
    index = 0
    while index < len(tokens):
        if tokens[index] in REDIRECTS:
            if words and words[-1].isdigit():
                words.pop()  # the file descriptor of `2>file`
            index += 2
            continue
        words.append(tokens[index])
        index += 1
    return words


def analyze_script(script: str, selection_vars: set[str]) -> tuple[list[tuple[str, list[str], int]], str]:
    """Return (cd target or "", python argv, level) per recognized line, and a reason.

    A non-empty reason means the step is unrecognized and earns no credit.
    """
    found: list[tuple[str, list[str], int]] = []
    reason = ""
    branch: int | None = None  # None outside the gate; level of the current branch inside
    for number, line in enumerate(logical_lines(script), 1):
        lexer = shlex.shlex(line, posix=True, punctuation_chars=";&|()<>")
        lexer.whitespace_split = True
        lexer.commenters = "#"
        try:
            tokens = list(lexer)
        except ValueError as error:
            reason = reason or f"line {number}: {error}"
            continue
        if not tokens:
            continue
        where = f"line {number} `{line.strip()}`"
        if any("`" in token or "$(" in token for token in tokens) or any(
            token in HEREDOCS for token in tokens
        ):
            reason = reason or f"{where}: command substitution or heredoc"
            continue
        rebound = sorted(var for var in selection_vars
                         if any(names_variable(token, var) for token in tokens))
        if rebound:
            reason = reason or f"{where}: the step may rebind the selection variable `{rebound[0]}`"
            continue
        if (branch is None and tokens[0] == "if" and GATE_LINE.fullmatch(line.strip())
                and is_selection_gate(tokens, selection_vars)):
            branch = SELECTED
            continue
        if branch is not None and tokens == ["else"]:
            branch = NEVER
            continue
        if branch is not None and tokens == ["fi"]:
            branch = None
            continue
        words = strip_redirections(tokens)
        while words and re.fullmatch(r"[A-Za-z_]\w*=\S*", words[0]):
            words.pop(0)  # leading variable assignments
        if not words:
            continue
        if any(token in CONTROL for token in words):
            reason = reason or f"{where}: command lists, pipelines, grouping and `!` are not recognized"
            continue
        head, level = words[0], ALWAYS if branch is None else branch
        if head in UNSAFE_HEADS:
            reason = reason or (
                f"{where}: only `if [[ \"$VAR\" == \"true\" ]]; then` on the selection output is recognized"
                if head == "if" else f"{where}: `{head}` is not recognized")
            continue
        if head == "set" and any(word.startswith("+") for word in words[1:]):
            reason = reason or f"{where}: `set +...` can disable error handling"
            continue
        if head == "cd":
            if len(words) != 2 or words[1].startswith("-") or branch is not None:
                reason = reason or f"{where}: only a top-level `cd <dir>` is recognized"
                continue
            found.append((words[1], [], level))
        elif head in PYTHON_NAMES:
            found.append(("", words[1:], level))
    if branch is not None:
        reason = reason or "unterminated `if`"
    return found, reason


def python_invocations(script: str, cwd: str, level: int, origin: str,
                       selection_vars: set[str], reason: str = "") -> list[Invocation]:
    lines, unrecognized = analyze_script(script, selection_vars)
    reason = reason or unrecognized
    found: list[Invocation] = []
    for target, argv, line_level in lines:
        if target:
            moved = normalize(cwd, target)
            if moved is None:
                reason = reason or f"`cd {target}` leaves the repository"
            else:
                cwd = moved
            continue
        if not argv or (argv[0] == "-m" and argv[1:2] != ["unittest"]):
            continue  # no script, or another module: not a test runner
        why = reason
        if not why and argv[0].startswith("-") and argv[0] != "-m":
            why = f"interpreter option `{argv[0]}` is not recognized"
        if not why and line_level == NEVER:
            why = "it runs only in the `else` branch of the full-proof gate"
        final = NEVER if why else max(level, line_level)
        found.append(Invocation(cwd, argv, final, origin, why))
    return found


def normalize(cwd: str, relative: str) -> str | None:
    joined = posixpath.normpath(posixpath.join(cwd, relative))
    if joined == ".." or joined.startswith("../") or posixpath.isabs(joined):
        return None
    return joined


def workflow_invocations(path: str, data: Any) -> tuple[list[Invocation], str]:
    if not isinstance(data, dict):
        raise CoverageError(f"{path}: not a mapping")
    every_pr, reason = runs_on_every_pull_request(data.get("on"))
    if not every_pr:
        return [], reason
    defaults = ((data.get("defaults") or {}).get("run") or {}) if isinstance(data.get("defaults"), dict) else {}
    jobs = data.get("jobs") or {}
    if not isinstance(jobs, dict):
        raise CoverageError(f"{path}: jobs is not a mapping")

    levels: dict[str, int] = {}

    def job_level(job_id: str, trail: tuple[str, ...] = ()) -> int:
        if job_id in levels:
            return levels[job_id]
        if job_id in trail or job_id not in jobs or not isinstance(jobs[job_id], dict):
            return NEVER
        job = jobs[job_id]
        level, ungated = condition_level(job.get("if"))
        needs = job.get("needs") or []
        needs = needs if isinstance(needs, list) else [needs]
        for need in needs:
            needed = job_level(str(need), trail + (job_id,))
            # A needed job outside the whitelist withholds credit even from an
            # always() job; a selected one gates this job unless always() runs it.
            if needed == NEVER or not ungated:
                level = max(level, needed)
        if truthy(job.get("continue-on-error")):
            level = NEVER
        levels[job_id] = level
        return level

    found: list[Invocation] = []
    for job_id, job in jobs.items():
        if not isinstance(job, dict) or "steps" not in job:
            continue
        level = job_level(job_id)
        job_run = ((job.get("defaults") or {}).get("run") or {}) if isinstance(job.get("defaults"), dict) else {}
        job_env = job.get("env") or {}
        steps = job.get("steps") or []
        if not isinstance(steps, list):
            raise CoverageError(f"{path}: jobs.{job_id}.steps is not a list")
        for number, step in enumerate(steps, 1):
            if not isinstance(step, dict) or "run" not in step:
                continue
            origin = f"{path} › {job.get('name', job_id)} › {step.get('name', f'step {number}')}"
            step_level = max(level, condition_level(step.get("if"))[0])
            why = ""
            if level == NEVER:
                why = "its job has an unrecognized condition, needs, or continue-on-error"
            elif step_level == NEVER:
                why = "its `if:` condition is not recognized"
            elif truthy(step.get("continue-on-error")):
                why = "the step is continue-on-error"
            shell = step.get("shell") or job_run.get("shell") or defaults.get("shell")
            if not why and shell not in (None, "bash"):
                why = f"shell `{shell}` is not recognized (only the default or `bash`)"
            cwd = step.get("working-directory") or job_run.get("working-directory") \
                or defaults.get("working-directory") or "."
            cwd = normalize(".", str(cwd))
            if cwd is None or "${{" in cwd:
                continue
            env = dict(job_env) if isinstance(job_env, dict) else {}
            env.update(step.get("env") or {})
            selection_vars = {str(name) for name, value in env.items()
                              if SELECTION_OUTPUT.fullmatch(str(value).strip())}
            found.extend(python_invocations(str(step["run"]), cwd, step_level, origin,
                                            selection_vars, why))
    return found, reason


# --------------------------------------------------------------------------
# Resolving invocations to test files.
# --------------------------------------------------------------------------


def defines_test_case(source: str) -> bool:
    try:
        tree = ast.parse(source)
    except SyntaxError:
        return False
    for node in ast.walk(tree):
        if isinstance(node, ast.FunctionDef) and node.name == "load_tests":
            return True
        if isinstance(node, ast.ClassDef):
            for base in node.bases:
                name = base.attr if isinstance(base, ast.Attribute) else getattr(base, "id", "")
                if name.endswith("TestCase"):
                    return True
    return False


def defines_load_tests(source: str) -> bool:
    try:
        tree = ast.parse(source)
    except SyntaxError:
        return True  # cannot establish what discovery does here
    return any(
        isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.Assign, ast.ImportFrom))
        and (getattr(node, "name", None) == "load_tests"
             or any(isinstance(t, ast.Name) and t.id == "load_tests" for t in getattr(node, "targets", []))
             or any(alias.asname == "load_tests" or alias.name == "load_tests"
                    for alias in getattr(node, "names", [])))
        for node in tree.body
    )


def runner_commands(root: Path, runner_path: str, runner: Runner) -> list[tuple[str, ...]]:
    try:
        tree = ast.parse((root / runner_path).read_text(encoding="utf-8"))
    except (OSError, SyntaxError) as error:
        raise CoverageError(f"cannot read runner {runner_path}: {error}") from error
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
            isinstance(target, ast.Name) and target.id == runner.variable for target in node.targets
        ):
            if not isinstance(node.value, (ast.Tuple, ast.List)):
                break
            commands: list[tuple[str, ...]] = []
            for item in node.value.elts:
                words: list[str] = []
                if isinstance(item, (ast.Tuple, ast.List)):
                    for element in item.elts:
                        if isinstance(element, ast.Constant) and isinstance(element.value, str):
                            words.append(element.value)
                        else:
                            words = []
                            break
                if not words:
                    raise CoverageError(
                        f"{runner_path}: {runner.variable} must hold literal string tuples"
                    )
                commands.append(tuple(words))
            return commands
    raise CoverageError(f"{runner_path}: no literal '{runner.variable} = (...)' list")


@dataclass
class Coverage:
    runs: dict[str, list[dict[str, Any]]] = field(default_factory=dict)
    notes: dict[str, list[str]] = field(default_factory=dict)
    stale: list[dict[str, str]] = field(default_factory=list)

    def credit(self, path: str, invocation: Invocation, how: str) -> None:
        self.runs.setdefault(path, []).append(
            {"origin": invocation.origin, "how": how, "level": invocation.level}
        )


def unittest_targets(root: Path, tracked: set[str], invocation: Invocation,
                     coverage: Coverage) -> list[str]:
    args = list(invocation.argv[2:])
    flags = {"-v", "--verbose", "-q", "--quiet", "-f", "--failfast", "-c", "--catch",
             "-b", "--buffer", "--locals", "--durations"}
    if any(arg.startswith("-k") for arg in args):  # a name filter may skip whole modules
        return []
    discover = not args or args[0] == "discover" or all(a.startswith("-") for a in args)
    if args and args[0] == "discover":
        args = args[1:]
    if not discover:
        targets: list[str] = []
        for arg in args:
            if arg in flags:
                continue
            if arg.startswith("-"):
                return []  # an option this reader does not know
            candidate = arg if arg.endswith(".py") else None
            if candidate is None:
                parts = arg.split(".")
                for size in range(len(parts), 0, -1):
                    probe = "/".join(parts[:size]) + ".py"
                    if normalize(invocation.cwd, probe) in tracked:
                        candidate = probe
                        break
            resolved = normalize(invocation.cwd, candidate) if candidate else None
            if resolved is None:
                continue
            # unittest loads only TestCase classes (or a load_tests hook) from a
            # named module; checks under `if __name__ == "__main__"` never run.
            if resolved in tracked and not defines_test_case(
                (root / resolved).read_text(encoding="utf-8", errors="replace")
            ):
                coverage.notes.setdefault(resolved, []).append(
                    f"named by {invocation.origin} but defines no unittest.TestCase"
                )
                continue
            targets.append(resolved)
        return targets
    start, pattern = ".", "test*.py"
    positional: list[str] = []
    index = 0
    while index < len(args):
        arg = args[index]
        if arg in ("-s", "--start-directory") and index + 1 < len(args):
            start = args[index + 1]
            index += 2
        elif arg in ("-p", "--pattern") and index + 1 < len(args):
            pattern = args[index + 1]
            index += 2
        elif arg in ("-t", "--top-level-directory") and index + 1 < len(args):
            index += 2
        elif arg in flags:
            index += 1
        elif arg.startswith("-"):
            return []
        else:
            positional.append(arg)
            index += 1
    if positional:
        start = positional[0]
    if len(positional) > 1:
        pattern = positional[1]
    base = normalize(invocation.cwd, start)
    if base is None:
        return []
    found: list[str] = []
    for path in sorted(tracked):
        name = posixpath.basename(path)
        relative = path if base == "." else (path[len(base) + 1:] if path.startswith(base + "/") else None)
        if relative is None or not TEST_FILE.fullmatch(name) or not fnmatch.fnmatchcase(name, pattern):
            continue
        # Discovery descends only into packages: every directory between the
        # start directory and the module must hold an __init__.py.
        parts = relative.split("/")[:-1]
        initializers = [
            posixpath.join(base, *parts[: depth + 1], "__init__.py").removeprefix("./")
            for depth in range(len(parts))
        ]
        if not all(init in tracked for init in initializers):
            continue
        # A package load_tests hook replaces automatic discovery below it and
        # may drop any module, so modules under one are not credited.
        start_init = posixpath.join(base, "__init__.py").removeprefix("./")
        hooked = [init for init in ([start_init] if start_init in tracked else []) + initializers
                  if defines_load_tests((root / init).read_text(encoding="utf-8", errors="replace"))]
        if hooked:
            coverage.notes.setdefault(path, []).append(
                f"discovered by {invocation.origin} under the load_tests hook in {hooked[0]}"
            )
            continue
        source = (root / path).read_text(encoding="utf-8", errors="replace")
        if defines_test_case(source):
            found.append(path)
        else:
            coverage.notes.setdefault(path, []).append(
                f"discovered by {invocation.origin} but defines no unittest.TestCase"
            )
    return found


def resolve(root: Path, tracked: set[str], invocations: list[Invocation]) -> Coverage:
    coverage = Coverage()
    queue = list(invocations)
    seen: set[tuple[str, tuple[str, ...], int, str]] = set()
    while queue:
        invocation = queue.pop(0)
        key = (invocation.cwd, tuple(invocation.argv), invocation.level, invocation.origin)
        if key in seen:
            continue
        seen.add(key)
        argv = invocation.argv
        if invocation.level == NEVER:
            # Name the withheld step on the file it would have run, so the
            # report says which step to restructure.
            script = normalize(invocation.cwd, argv[0]) if argv[0] != "-m" else None
            if script is not None and TEST_FILE.fullmatch(posixpath.basename(script)):
                coverage.notes.setdefault(script, []).append(
                    f"not credited from {invocation.origin}: {invocation.reason or 'never runs'}"
                )
            continue
        if argv[0] == "-m":
            if len(argv) > 1 and argv[1] == "unittest":
                for path in unittest_targets(root, tracked, invocation, coverage):
                    if path in tracked:
                        coverage.credit(path, invocation, "python -m unittest")
                    elif TEST_FILE.fullmatch(posixpath.basename(path)):
                        coverage.stale.append({"origin": invocation.origin, "path": path})
            continue
        script = normalize(invocation.cwd, argv[0])
        if script is None:
            continue
        if script not in tracked:
            if TEST_FILE.fullmatch(posixpath.basename(script)):
                coverage.stale.append({"origin": invocation.origin, "path": script})
            continue
        if TEST_FILE.fullmatch(posixpath.basename(script)):
            coverage.credit(script, invocation, "python " + shlex.join(argv))
        runner = RUNNERS.get(script)
        if runner is None:
            continue
        reduced = runner.reduced_flag is not None and runner.reduced_flag in argv[1:]
        for command in runner_commands(root, script, runner):
            if reduced and not runner.keeps_in_reduced(command):
                continue
            queue.append(Invocation(
                posixpath.dirname(script) or ".", list(command), invocation.level,
                f"{invocation.origin} › {script} {runner.variable}",
            ))
    return coverage


def default_metadata_only(paths: list[str]) -> bool:
    here = str(Path(__file__).resolve().parent)
    if here not in sys.path:
        sys.path.insert(0, here)
    import select_accuracy_checks

    return bool(select_accuracy_checks.metadata_only(paths))


def run_check(root: Path, *, exempt: dict[str, str] | None = None,
              metadata_only: Callable[[list[str]], bool] = default_metadata_only) -> dict[str, Any]:
    exempt = EXEMPT if exempt is None else exempt
    tracked = tracked_files(root)
    tests = sorted(p for p in tracked if TEST_FILE.fullmatch(posixpath.basename(p)))
    workflows: list[dict[str, Any]] = []
    invocations: list[Invocation] = []
    for path in sorted(p for p in tracked if posixpath.dirname(p) == WORKFLOW_DIR
                       and p.endswith((".yml", ".yaml"))):
        try:
            data = load_yaml((root / path).read_text(encoding="utf-8"))
        except CoverageError as error:
            raise CoverageError(f"{path}: {error}") from error
        found, reason = workflow_invocations(path, data)
        workflows.append({"path": path, "credited": bool(reason == "every pull request"),
                          "reason": reason})
        invocations.extend(found)
    coverage = resolve(root, tracked, invocations)

    findings: list[dict[str, str]] = []
    rows: list[dict[str, Any]] = []
    for path in tests:
        runs = coverage.runs.get(path, [])
        best = min((run["level"] for run in runs), default=NEVER)
        row: dict[str, Any] = {"path": path, "runs": runs, "level": LEVEL_NAMES.get(best, "never")}
        rows.append(row)
        if path in exempt:
            row["exempt"] = exempt[path]
            if runs:
                findings.append({"path": path, "message": "exempt, but CI already runs it; "
                                 "remove the exemption"})
            continue
        if not runs:
            extra = "; ".join(coverage.notes.get(path, []))
            findings.append({"path": path, "message": (
                "no CI step or runner list executes it on every pull request; add it to "
                "benchmark/test_evidence_tools.py (benchmark tests) or to a workflow step"
                + (f" ({extra})" if extra else ""))})
        elif best == SELECTED and metadata_only([path]):
            findings.append({"path": path, "message": (
                "runs only with full proofs, but tools/select_accuracy_checks.py treats a "
                "change to it as metadata-only, so editing it alone would skip it")})
    for path in sorted(set(exempt) - set(tests)):
        findings.append({"path": path, "message": "exemption names an untracked test file"})
    for item in coverage.stale:
        findings.append({"path": item["path"], "message": f"{item['origin']} runs a test file "
                         "that is not tracked"})
    return {
        "tool": TOOL_NAME,
        "status": "fail" if findings else "pass",
        "tests": rows,
        "workflows": workflows,
        "findings": findings,
    }


def markdown_report(report: dict[str, Any]) -> str:
    status = "PASS" if report["status"] == "pass" else "FAIL"
    lines = [f"## {TOOL_NAME}: {status}", ""]
    lines.append(f"{len(report['tests'])} tracked test files; "
                 f"{sum(1 for w in report['workflows'] if w['credited'])} workflows run on every "
                 "pull request.")
    lines += ["", "| Test file | Runs | Executed by |", "| --- | --- | --- |"]
    for row in report["tests"]:
        origins = sorted({run["origin"] for run in row["runs"]})
        where = "<br>".join(origins) if origins else row.get("exempt", "nothing")
        lines.append(f"| `{row['path']}` | {row['level']} | {where} |")
    if report["findings"]:
        lines += ["", "### Findings", ""]
        lines += [f"- `{item['path']}`: {item['message']}" for item in report["findings"]]
    return "\n".join(lines) + "\n"


def parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--output", type=Path, help="write the report as JSON")
    parser.add_argument("--summary", type=Path, help="write the Markdown report to this file")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    options = parse_args(argv)
    root = options.root.resolve()
    return run_gate(
        options,
        build=lambda: run_check(root),
        markdown=markdown_report,
        errors=(CoverageError, RuntimeError, OSError, UnicodeDecodeError),
    )


if __name__ == "__main__":
    raise SystemExit(main())
