#!/usr/bin/env python3
"""Check or apply the layout of the exported VBA modules.

The layout rules are mechanical and behaviour-neutral (see "VBA source layout"
in tools/README.md):

1. CRLF line endings, one final line ending, no blank lines at the end;
2. no trailing white space;
3. four spaces for each tab of a line's indentation;
4. procedure declarations, their End statements and labels at the outer
   level (column one);
5. exactly two blank lines after each procedure's End statement.

Every other line keeps its indentation, so the repository's indented comment
headings inside procedures ('Declare, 'Invert, ...) and its continuation-line
layout are left as they are. No blank line is added to or removed from a
logical line: a line that a continuation carries over, in code or in a
comment, is kept as it stands, blank or not. No rule changes a string literal or a comment's
text: rules work on indentation, on the end of a line, or on whole lines they
recognise outside literals, and a comment starts only at an apostrophe outside
a literal. The header the VBE writes (VERSION, BEGIN ... END, Object and
Attribute lines) is kept byte for byte.

--check reports every tracked .bas, .cls and .frm module (or each given path)
that differs from its formatted form. Until the out-of-layout modules are
reformatted together with an Excel evidence refresh, the committed baseline
tools/vba_layout_baseline.json lists them, and the check ratchets it:

- a module out of layout that the baseline does not list fails;
- a listed module that is in layout fails until its entry is removed;
- a listed path that is not a tracked VBA module fails.

--write rewrites modules in place. Every module is bound by hash to Excel
evidence, so rewrite one only together with its re-import into the VBE and the
evidence refresh. The tool neither compiles nor runs VBA.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from collections.abc import Sequence
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BASELINE = "tools/vba_layout_baseline.json"
EXTENSIONS = frozenset({".bas", ".cls", ".frm"})
INDENT = "    "
BLANK_LINES_AFTER_END = 2

PROCEDURE = re.compile(r"^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
                       r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+[A-Za-z_]\w*",
                       re.IGNORECASE)
PROCEDURE_END = re.compile(r"^End\s+(?:Sub|Function|Property)\b", re.IGNORECASE)
LABEL = re.compile(r"^([A-Za-z_]\w*):$")
# Statement keywords that a colon separator can follow; none of them is a label.
NOT_LABELS = frozenset({"do", "else", "end", "loop", "next", "rem", "resume", "return",
                        "stop", "wend"})
# The line-continuation rule of tools/vba_lexer.py: white space, then a final _.
CONTINUATION = re.compile(r"\s_$")
HEADER_LINE = re.compile(r"^(?:VERSION\s|Object\s*=|Attribute\s)", re.IGNORECASE)


class BaselineError(ValueError):
    """The baseline file is missing, malformed or lists an invalid path."""


def split_code_comment(line: str) -> tuple[str, str]:
    """Split a line at the apostrophe that starts its comment, if any.

    An apostrophe inside a string literal is data. A doubled quote inside a
    literal toggles the scan twice, which leaves it inside the literal.
    """
    in_literal = False
    for index, char in enumerate(line):
        if char == '"':
            in_literal = not in_literal
        elif char == "'" and not in_literal:
            return line[:index], line[index:]
    return line, ""


def header_length(lines: Sequence[str]) -> int:
    """The number of leading lines the VBE writes and the formatter keeps."""
    depth = 0
    for index, line in enumerate(lines):
        word = line.strip().split(" ", 1)[0].upper()
        if word == "BEGIN":
            depth += 1
        elif word == "END" and depth:
            depth -= 1
        elif not depth and not HEADER_LINE.match(line):
            return index
    return len(lines)


def expand_indentation(line: str) -> str:
    """Replace each tab of the leading white space with four spaces."""
    body = line.lstrip(" \t")
    return line[:len(line) - len(body)].replace("\t", INDENT) + body


def is_label(code: str) -> bool:
    match = LABEL.match(code)
    return match is not None and match.group(1).casefold() not in NOT_LABELS


def outer_level(line: str) -> bool:
    """Whether a line is a procedure declaration, its End statement or a label."""
    code = split_code_comment(line.strip())[0].strip()
    return bool(PROCEDURE.match(code) or PROCEDURE_END.match(code) or is_label(code))


def continues(line: str) -> bool:
    """Whether the next physical line belongs to this line's logical line.

    A line continuation also continues a comment, so the whole line is tested.
    """
    return bool(CONTINUATION.search(line))


def format_body(lines: Sequence[str]) -> list[str]:
    """The formatted body lines.

    No blank line is ever inserted into, or removed from, one logical line:
    a physical line that a continuation (in code or in a comment) carries
    over is emitted as it stands, blank or not, and the blank lines after a
    procedure's End wait until its whole logical line has been emitted.
    """
    out: list[str] = []
    continued = False
    pending_end = False
    for raw in lines:
        line = expand_indentation(raw.rstrip())
        if continued:
            out.append(line)
            continued = continues(line)
            continue
        if not line:
            if not pending_end:
                out.append(line)
            continue
        if pending_end:
            # Exactly two blank lines between a procedure's End and what follows it.
            out.extend([""] * BLANK_LINES_AFTER_END)
            pending_end = False
        if outer_level(line):
            line = line.lstrip()
            pending_end = bool(PROCEDURE_END.match(split_code_comment(line)[0].strip()))
        out.append(line)
        continued = continues(line)
    # A blank line that ends a continuation belongs to it and stays.
    while out and not out[-1] and not (len(out) > 1 and continues(out[-2])):
        out.pop()
    return out


def format_text(text: str) -> str:
    """The formatted form of one module's text; idempotent."""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if lines and lines[-1] == "":
        lines.pop()
    split = header_length(lines)
    body = format_body(lines[split:])
    return "".join(f"{line}\r\n" for line in [*lines[:split], *body])


def tracked_modules(root: Path) -> list[str]:
    """Every tracked VBA module, as a sorted list of root-relative POSIX paths."""
    completed = subprocess.run(["git", "-C", str(root), "ls-files", "-z"],
                               capture_output=True, check=False)
    if completed.returncode != 0:
        raise RuntimeError(completed.stderr.decode("utf-8", errors="replace").strip()
                           or "git ls-files failed")
    names = (item.decode("utf-8", errors="surrogateescape")
             for item in completed.stdout.split(b"\0") if item)
    return sorted(name for name in names if Path(name).suffix.casefold() in EXTENSIONS)


def load_baseline(path: Path) -> list[str]:
    """The sorted, unique list of root-relative paths that the baseline file names."""
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        raise BaselineError(f"cannot read baseline {path}: {error}") from error
    if not isinstance(data, dict) or data.get("schema_version") != 1:
        raise BaselineError(f"{path}: expected an object with schema_version 1")
    files = data.get("files")
    if not isinstance(files, list) or not all(isinstance(name, str) for name in files):
        raise BaselineError(f"{path}: 'files' must be a list of paths")
    if files != sorted(set(files)):
        raise BaselineError(f"{path}: 'files' must be sorted and unique")
    for name in files:
        if name.startswith("/") or "\\" in name or ".." in Path(name).parts:
            raise BaselineError(f"{path}: {name!r} is not a root-relative POSIX path")
    return files


def shown(path: Path, root: Path) -> str:
    """A path relative to the root when it is inside it, else absolute."""
    resolved = path.resolve()
    if resolved.is_relative_to(root):
        return resolved.relative_to(root).as_posix()
    return resolved.as_posix()


def read(path: Path) -> str:
    # Latin-1 maps every byte to one character, so any byte round-trips.
    return path.read_bytes().decode("latin-1")


def check(root: Path, paths: Sequence[Path], everything: bool,
          baseline: list[str], baseline_name: str) -> int:
    listed = set(baseline)
    failures = 0
    for path in paths:
        name = shown(path, root)
        text = read(path)
        in_layout = format_text(text) == text
        if name in listed:
            if in_layout:
                failures += 1
                print(f"FAIL: {name}: in layout but still listed in {baseline_name}; "
                      "remove its entry")
            else:
                print(f"baselined: {name}: out of layout, listed in {baseline_name}")
        elif not in_layout:
            failures += 1
            print(f"FAIL: {name}: not in layout; run python tools/format_vba.py --write "
                  f"{name}, re-import it into the VBE and refresh its Excel evidence")
    if everything:
        known = {shown(path, root) for path in paths}
        for name in baseline:
            if name not in known:
                failures += 1
                print(f"FAIL: {name}: listed in {baseline_name} but not a tracked VBA module")
    status = "FAIL" if failures else "PASS"
    print(f"vba-layout: {status} ({len(paths)} module(s), {len(baseline)} baselined; "
          "layout only, nothing compiled)")
    return 1 if failures else 0


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--check", action="store_true",
                      help="report modules out of layout, ratcheting the baseline")
    mode.add_argument("--write", action="store_true", help="rewrite modules in place")
    parser.add_argument("--root", type=Path, default=ROOT)
    source = parser.add_mutually_exclusive_group()
    source.add_argument("--baseline", type=Path,
                        help=f"baseline file (default: {BASELINE} under the root, "
                             "if it exists)")
    source.add_argument("--no-baseline", action="store_true",
                        help="check without a baseline: every module must be in layout")
    parser.add_argument("paths", nargs="*", type=Path,
                        help="modules to process (default: every tracked VBA module)")
    args = parser.parse_args(argv)
    root = args.root.resolve()
    # Relative paths are relative to the root, wherever the tool runs from.
    paths = [path if path.is_absolute() else root / path for path in args.paths]
    baseline_path = args.baseline or Path(BASELINE)
    baseline_path = baseline_path if baseline_path.is_absolute() else root / baseline_path
    baseline: list[str] = []
    try:
        if not paths:
            paths = [root / name for name in tracked_modules(root)]
        # The default baseline is optional: it is deleted once every module is in layout.
        if not args.no_baseline and (args.baseline or baseline_path.is_file()):
            baseline = load_baseline(baseline_path)
    except (BaselineError, RuntimeError) as error:
        print(f"vba-layout: ERROR: {error}")
        return 2
    if not paths:
        print("vba-layout: PASS (no VBA module found)")
        return 0
    if args.check:
        return check(root, paths, not args.paths, baseline, shown(baseline_path, root))
    written = 0
    for path in paths:
        text = read(path)
        formatted = format_text(text)
        if formatted != text:
            path.write_bytes(formatted.encode("latin-1"))
            written += 1
            print(f"formatted: {shown(path, root)}")
    print(f"vba-layout: {written} of {len(paths)} module(s) rewritten")
    if written:
        print("Re-import each formatted module into the VBE, refresh its Excel evidence "
              f"and remove it from {BASELINE}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
