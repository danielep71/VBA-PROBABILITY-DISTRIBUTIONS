"""Shared reading and lexing of exported VBA source for the VBA source checks.

Statically splits VBA into logical statements: comments removed, line
continuations joined, string literals kept or blanked. It is not a VBA
parser or compiler. tools/README.md, "VBA source checks", lists the rules the
checks enforce and what they cannot see.

check_vba_source.py, check_vba_conditionals.py and check_vba_visibility.py use
it. check_vba_jumps.py keeps its own lexer.
"""

from __future__ import annotations

import re
import subprocess
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path

VBA_SUFFIXES = frozenset({".bas", ".cls", ".frm"})
REM = re.compile(r"Rem(?=\s|$)", re.IGNORECASE)
CONTINUATION = re.compile(r"\s_\s*$")
LINE_NUMBER = re.compile(r"\s*\d+(?=\s)")


@dataclass(frozen=True)
class Finding:
    """One defect, located by repository path and 1-based line when known."""

    path: str
    line: int | None
    message: str

    def __str__(self) -> str:
        return f"{self.path}:{self.line}: {self.message}" if self.line else \
            f"{self.path}: {self.message}"


def report(gate: str, findings: Sequence[Finding], counted: str) -> int:
    """Print the findings and a one-line verdict; return the exit code.

    ``counted`` says what was examined, so that a pass over nothing reads as
    such: these are static checks, never a compilation or a calculation.
    """
    for finding in findings:
        print(f"FAIL: {finding}")
    status = "FAIL" if findings else "PASS"
    print(f"{gate}: {status} ({counted}; static check, nothing compiled or run)")
    return 1 if findings else 0


def tracked_files(root: Path) -> list[str]:
    """Paths tracked by Git (staged ones included), sorted, with "/" separators."""
    completed = subprocess.run(["git", "-C", str(root), "ls-files", "-z"],
                               capture_output=True, check=False)
    if completed.returncode:
        raise RuntimeError(completed.stderr.decode("utf-8", errors="replace").strip())
    return sorted(item.decode("utf-8", errors="surrogateescape")
                  for item in completed.stdout.split(b"\0") if item)


def tracked_components(root: Path) -> list[str]:
    """Tracked exported VBA components: .bas, .cls and .frm files."""
    return [path for path in tracked_files(root) if Path(path).suffix.lower() in VBA_SUFFIXES]


def read_component(root: Path, path: str) -> str:
    """Component text with LF line ends. A non-ASCII byte becomes U+FFFD; the
    vba-source check reports it, the other checks still analyze the file."""
    text = (root / path).read_bytes().decode("ascii", errors="replace")
    return text.replace("\r\n", "\n")


def strip_code(raw: str, *, blank_strings: bool = False) -> str:
    """One physical line without its comment: from a ' outside a string and
    outside [...] (where it quotes a sheet name, as in ['Sheet 1'!A1]), or
    from Rem where a statement starts outside [...] (inside, as in
    [A1:Rem], the colon is a range operator and Rem a name) (at the line start, after a numeric line
    label such as ``100``, or after a ``:`` separator, not after a ``:=``
    named argument). With ``blank_strings`` the content of string literals
    and of [...] expressions becomes spaces, so that keywords quoted in a
    string or a sheet name are not read as code; the brackets stay, and
    columns are preserved either way."""
    label = LINE_NUMBER.match(raw)
    index = label.end() if label else 0
    result: list[str] = [raw[:index]]
    in_string = False
    statement_start = True
    brackets = 0
    while index < len(raw):
        char = raw[index]
        if char == '"':
            if in_string and raw[index + 1:index + 2] == '"':
                result.append("  " if blank_strings else '""')
                index += 2
                continue
            in_string = not in_string
            result.append(" " if blank_strings else char)
        elif not in_string and not brackets and (
                char == "'" or (statement_start and REM.match(raw, index))):
            break
        else:
            hidden = in_string or (brackets and char not in "[]")
            result.append(" " if blank_strings and hidden else char)
        if not in_string and char == "[":
            brackets += 1
        elif not in_string and char == "]" and brackets:
            brackets -= 1
        if not in_string and not brackets and char == ":" and raw[index + 1:index + 2] != "=":
            statement_start = True
        elif not char.isspace():
            statement_start = False
        index += 1
    return "".join(result).rstrip()


def date_delimiter(code: str, index: int, in_date: bool) -> bool:
    """Whether the # at ``index`` opens or closes a date literal such as #1/2/2026#."""
    if in_date:
        return True
    previous = code[index - 1] if index else ""
    if previous and (previous.isalnum() or previous in "_.)]"):
        return False  # a type suffix, as in Total#
    return "#" in code[index + 1:]  # a literal needs its closing #; Print #1 has none


def split_statements(code: str) -> list[str]:
    """Split one logical line at colons outside strings, date literals and
    [...] expressions (such as [A1:B2]); a ``:=`` named argument is not a
    separator, and a Rem statement ends the line."""
    statements: list[str] = []
    start = brackets = 0
    in_string = in_date = False
    index = 0
    while index < len(code):
        char = code[index]
        if char == '"' and not in_date:
            if in_string and code[index:index + 2] == '""':
                index += 2
                continue
            in_string = not in_string
        elif in_string or (in_date and char != "#"):
            pass
        elif char == "#" and date_delimiter(code, index, in_date):
            in_date = not in_date
        elif char in "[]":
            brackets = brackets + 1 if char == "[" else max(brackets - 1, 0)
        elif char == ":" and not brackets and code[index:index + 2] != ":=":
            statements.append(code[start:index].strip())
            start = index + 1
        index += 1
    statements.append(code[start:].strip())
    rem = next((i for i, item in enumerate(statements) if re.match(r"^Rem(?:\s|$)", item, re.I)),
               len(statements))
    return statements[:rem]


def logical_units(lines: list[str], *, blank_strings: bool = False
                  ) -> list[tuple[int, int, str, str]]:
    """``(first line, last line, kind, code)`` per logical line, numbered from 1.

    ``kind`` is ``"directive"`` for a conditional-compilation line (starting
    with ``#``) and ``"code"`` otherwise; continuation lines are joined. A
    comment that ends in " _" continues on the next physical line: those
    lines are comment, not code, and extend the unit the comment belongs to.
    """
    units: list[tuple[int, int, str, str]] = []
    buffer: list[str] = []
    start = 0
    in_comment = False
    for number, raw in enumerate(lines, start=1):
        if in_comment:
            in_comment = bool(CONTINUATION.search(raw))
            if units and not buffer:
                first, _, kind, code = units[-1]
                units[-1] = (first, number, kind, code)
            continue
        code = strip_code(raw, blank_strings=blank_strings)
        # strip_code keeps every column before a comment, so a shorter
        # result means the line ends in a comment.
        in_comment = len(code) < len(raw.rstrip()) and bool(CONTINUATION.search(raw))
        if not buffer and code.strip().startswith("#"):
            units.append((number, number, "directive", code.strip()))
            continue
        if not buffer:
            start = number
        if CONTINUATION.search(code):
            buffer.append(CONTINUATION.sub(" ", code))
            continue
        buffer.append(code)
        units.append((start, number, "code", " ".join(part.strip() for part in buffer)))
        buffer.clear()
    if buffer:
        units.append((start, len(lines), "code", " ".join(part.strip() for part in buffer)))
    return units


def statements(text: str) -> list[tuple[int, str]]:
    """``(line, statement)`` for every code statement, comments removed."""
    return [(line, part) for line, _, kind, code in logical_units(text.split("\n"))
            if kind == "code" for part in split_statements(code)]
