#!/usr/bin/env python3
"""Fail unless the Python df-cap mirrors equal the VBA envelope constants, bit for bit (#86).

The Accuracy Gate decides whether a Student-t, chi-square or F row lies in an
envelope-reject region from F_MAX_DF, T_MAX_DF and CHI_MAX_DF in
_contract_eval.py. The library enforces the same caps with PROB_F_MAX_DF,
PROB_T_MAX_DF and PROB_CHI_MAX_DF in src/M_STATS_PROBDIST_TFAMILY.bas. If one
side changes alone, the gate scores rows the library refuses, or skips rows it
now accepts. This check reads both sides and compares the binary64 bit patterns.

VBA side. tools/vba_lexer.py removes comments and joins line continuations. Each
cap must be declared exactly once, in the module, outside conditional
compilation, in the form

    Private|Public Const PROB_X_MAX_DF As Double = <numeric literal>

Any other statement that declares the name, or mentions it inside a declaration
(Const, Dim, Static, ReDim, Private, Public, Global, Friend, Sub, Function,
Property, Declare, Enum, Type or Event, or an Enum member), in that module or in
any other src/*.bas module, fails: it is either a duplicate or a form this check
cannot verify. An expression, a reference to another constant, a string, a
hexadecimal or octal literal, or a declared type other than Double also fails.

Literal conversion. The literal is decimal digits with an optional fraction,
an optional E or D exponent and an optional type suffix (# Double, ! Single,
@ Currency, % Integer, & Long, ^ LongLong). It is read as an exact rational
number. It must be exactly representable in the type its suffix selects and,
after the conversion to the declared Double, exactly representable in
binary64. With an exact value every correctly working parser yields the same
Double, so the result does not depend on how VBA rounds a decimal literal:
that rounding is not proven against Python (#32). A cap that is not exactly
representable fails instead of being guessed.

Python side. _contract_eval.py must bind each mirror exactly once at module
level, to a float literal, and never through a global statement; the imported
module must hold that same value.

Run: python3 check_df_cap_mirrors.py   (exit 0 = mirrors agree, 1 = drift or
an unverifiable declaration). test_df_cap_mirrors.py holds the negative
fixtures. Standard library only.
"""
from __future__ import annotations

import argparse
import ast
import importlib.util
import math
import re
import struct
import sys
from fractions import Fraction
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT / "tools"))
import vba_lexer  # noqa: E402

SOURCE = "src/M_STATS_PROBDIST_TFAMILY.bas"
PYTHON = HERE / "_contract_eval.py"
# (VBA constant, Python mirror)
PAIRS = (
    ("PROB_F_MAX_DF", "F_MAX_DF"),
    ("PROB_T_MAX_DF", "T_MAX_DF"),
    ("PROB_CHI_MAX_DF", "CHI_MAX_DF"),
)

CANONICAL = re.compile(
    r"^(?P<scope>Private|Public)\s+Const\s+(?P<name>[A-Za-z]\w*)\s+As\s+(?P<type>\w+)"
    r"\s*=\s*(?P<literal>.*?)\s*$", re.IGNORECASE)
DECLARATION = re.compile(
    r"^(?:Const|Dim|Static|ReDim|Private|Public|Global|Friend|Sub|Function|Property|"
    r"Declare|Enum|Type|Event)\b", re.IGNORECASE)
ENUM_START = re.compile(r"^(?:(?:Private|Public)\s+)?Enum\b", re.IGNORECASE)
ENUM_END = re.compile(r"^End\s+Enum\b", re.IGNORECASE)
LINE_LABEL = re.compile(r"^\d+\s+")
LITERAL = re.compile(
    r"^(?P<sign>[+-]?)\s*(?P<int>\d*)(?:\.(?P<frac>\d*))?"
    r"(?:[EeDd](?P<exp>[+-]?\d+))?(?P<suffix>[#!@%&^]?)$")
INTEGER_LIMITS = {"%": 2**15 - 1, "&": 2**31 - 1, "^": 2**63 - 1}
CURRENCY_LIMIT = Fraction(2**63 - 1, 10_000)
MAX_EXPONENT = 10_000


class CapError(ValueError):
    """A cap that cannot be read, or cannot be converted without guessing."""


def bits(value: float) -> str:
    """The 16-hex-digit IEEE-754 binary64 bit pattern of value."""
    return struct.pack(">d", value).hex()


def _exact_double(value: Fraction, what: str) -> float:
    try:
        result = float(value)  # correctly rounded, ties to even
    except OverflowError:
        raise CapError(f"{what} is outside the Double range") from None
    if not math.isfinite(result) or Fraction(result) != value:
        raise CapError(f"{what} is not exactly representable as a binary64 Double; "
                       "VBA's rounding of such a literal is not proven (#32), so give "
                       "the cap as an exactly representable value")
    return result


def vba_double(literal: str) -> float:
    """The Double a VBA numeric literal assigned to an As Double constant denotes.

    Raises CapError for anything that is not one exactly representable numeric
    literal; see the module docstring for the accepted grammar."""
    text = literal.strip()
    match = LITERAL.match(text)
    if not match or not (match["int"] or match["frac"]):
        raise CapError(f"unparsable literal {literal!r}: expected one decimal numeric "
                       "literal (an expression, constant name, string, or hexadecimal or "
                       "octal literal is not accepted)")
    exponent = int(match["exp"]) if match["exp"] is not None else 0
    if abs(exponent) > MAX_EXPONENT:
        raise CapError(f"literal {literal!r} is outside the Double range")
    digits = (match["int"] or "") + (match["frac"] or "")
    magnitude = Fraction(int(digits)) * Fraction(10) ** (exponent - len(match["frac"] or ""))
    suffix = match["suffix"]
    if suffix in INTEGER_LIMITS:
        if match["frac"] is not None or match["exp"] is not None:
            raise CapError(f"literal {literal!r} puts an integer type suffix on a "
                           "non-integer literal")
        if magnitude > INTEGER_LIMITS[suffix]:
            raise CapError(f"literal {literal!r} overflows its {suffix} integer type")
    elif suffix == "@":
        if match["exp"] is not None:
            raise CapError(f"literal {literal!r} puts the Currency suffix on an exponent form")
        if magnitude > CURRENCY_LIMIT or (magnitude * 10_000).denominator != 1:
            raise CapError(f"literal {literal!r} is not exactly a Currency value")
    if match["sign"] == "-" and magnitude == 0:
        raise CapError(f"literal {literal!r} is a negative zero, whose VBA sign is not proven")
    value = -magnitude if match["sign"] == "-" else magnitude
    result = _exact_double(value, f"literal {literal!r}")
    if suffix == "!":
        try:
            single = struct.unpack(">f", struct.pack(">f", result))[0]
        except OverflowError:
            raise CapError(f"literal {literal!r} is outside the Single range") from None
        if single != result:
            raise CapError(f"literal {literal!r} is not exactly representable as a Single")
    return result


def _mentions(name: str, statement: str) -> bool:
    return re.search(rf"(?<![\w]){re.escape(name)}(?![\w])", statement, re.IGNORECASE) \
        is not None


def scan_module(text: str, names: tuple[str, ...]
                ) -> tuple[dict[str, list[tuple[int, str]]], list[str]]:
    """Canonical declarations per name as ``(line, literal)``, and the problems found.

    A problem is any other declaration that mentions one of the names, or a
    canonical declaration inside conditional compilation."""
    found: dict[str, list[tuple[int, str]]] = {name: [] for name in names}
    problems: list[str] = []
    depth = 0
    in_enum = False
    for first, _, kind, code in vba_lexer.logical_units(text.split("\n"), blank_strings=True):
        if kind == "directive":
            if re.match(r"^#If\b", code, re.IGNORECASE):
                depth += 1
            elif re.match(r"^#End\s*If\b", code, re.IGNORECASE):
                depth = max(depth - 1, 0)
            continue
        for statement in vba_lexer.split_statements(code):
            statement = LINE_LABEL.sub("", statement).strip()
            if ENUM_START.match(statement):
                in_enum = True
            elif ENUM_END.match(statement):
                in_enum = False
            declares = in_enum or DECLARATION.match(statement) is not None
            for name in names:
                if not declares or not _mentions(name, statement):
                    continue
                canonical = CANONICAL.match(statement)
                if canonical and canonical["name"].lower() == name.lower():
                    if canonical["type"].lower() != "double":
                        problems.append(f"line {first}: {name} is declared As "
                                        f"{canonical['type']}, not As Double")
                    elif depth:
                        problems.append(f"line {first}: {name} is declared inside "
                                        "conditional compilation")
                    else:
                        found[name].append((first, canonical["literal"]))
                    continue
                problems.append(f"line {first}: a declaration mentions {name} in a form "
                                f"this check cannot verify: {statement}")
    return found, problems


def vba_caps(root: Path) -> tuple[dict[str, float], list[str]]:
    """The VBA cap values by constant name, and every failure found."""
    names = tuple(vba for vba, _ in PAIRS)
    failures: list[str] = []
    values: dict[str, float] = {}
    source = root / SOURCE
    if not source.is_file():
        return values, [f"{SOURCE} not found"]
    found, problems = scan_module(vba_lexer.read_component(root, SOURCE), names)
    failures += [f"{SOURCE} {problem}" for problem in problems]
    for name in names:
        declarations = found[name]
        if not declarations:
            failures.append(f"{SOURCE}: {name} is not declared as "
                            f"'Private|Public Const {name} As Double = <literal>'")
            continue
        if len(declarations) > 1:
            lines = ", ".join(str(line) for line, _ in declarations)
            failures.append(f"{SOURCE}: {name} is declared {len(declarations)} times "
                            f"(lines {lines})")
            continue
        line, literal = declarations[0]
        try:
            values[name] = vba_double(literal)
        except CapError as exc:
            failures.append(f"{SOURCE} line {line}: {name}: {exc}")
    for other in sorted((root / "src").glob("*.bas")):
        relative = other.relative_to(root).as_posix()
        if relative == SOURCE:
            continue
        found, problems = scan_module(vba_lexer.read_component(root, relative), names)
        failures += [f"{relative} {problem}" for problem in problems]
        for name, declarations in found.items():
            failures += [f"{relative} line {line}: duplicate declaration of {name}; "
                         f"it belongs only in {SOURCE}" for line, _ in declarations]
    return values, failures


def _module_bindings(tree: ast.Module, name: str) -> list[ast.AST]:
    """Module-scope statements and expressions that bind name."""
    found: list[ast.AST] = []
    stack: list[ast.AST] = list(tree.body)
    while stack:
        node = stack.pop()
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            if node.name == name:
                found.append(node)
            stack.extend(node.decorator_list)
            continue  # a nested scope binds its own names
        if isinstance(node, (ast.Lambda, ast.ListComp, ast.SetComp, ast.DictComp,
                             ast.GeneratorExp)):
            continue
        if isinstance(node, ast.Name) and node.id == name \
                and isinstance(node.ctx, (ast.Store, ast.Del)):
            found.append(node)
        if isinstance(node, (ast.Import, ast.ImportFrom)):
            found += [node for alias in node.names
                      if (alias.asname or alias.name.split(".")[0]) in (name, "*")]
        stack.extend(ast.iter_child_nodes(node))
    return found


def python_caps(path: Path) -> tuple[dict[str, float], list[str]]:
    """The Python mirror values by name, and every failure found."""
    failures: list[str] = []
    values: dict[str, float] = {}
    label = path.name
    try:
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    except (OSError, SyntaxError, UnicodeDecodeError) as exc:
        return values, [f"{label} unreadable: {exc}"]
    globals_declared = {name for node in ast.walk(tree) if isinstance(node, ast.Global)
                        for name in node.names}
    literals: dict[str, float] = {}
    for _, name in PAIRS:
        if name in globals_declared:
            failures.append(f"{label}: {name} is rebound through a global statement")
            continue
        bindings = _module_bindings(tree, name)
        if not bindings:
            failures.append(f"{label}: {name} is not assigned at module level")
            continue
        if len(bindings) > 1:
            lines = ", ".join(str(getattr(node, "lineno", "?")) for node in bindings)
            failures.append(f"{label}: {name} is bound {len(bindings)} times (lines {lines})")
            continue
        node = bindings[0]
        statement = next((s for s in tree.body
                          if isinstance(s, ast.Assign) and len(s.targets) == 1
                          and s.targets[0] is node), None)
        try:
            value = ast.literal_eval(statement.value) if statement else None
        except ValueError:
            value = None
        if type(value) is not float or not math.isfinite(value):
            failures.append(f"{label} line {getattr(node, 'lineno', '?')}: {name} is not "
                            "a single module-level assignment of a finite float literal")
            continue
        literals[name] = value
    if failures:
        return values, failures
    spec = importlib.util.spec_from_file_location(f"_df_cap_mirror_{id(path)}", path)
    if spec is None or spec.loader is None:
        return values, [f"{label}: cannot be imported"]
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as exc:  # report any import failure as a check failure
        return values, [f"{label}: import failed: {exc}"]
    for name, literal in literals.items():
        runtime = getattr(module, name, None)
        if type(runtime) is not float or bits(runtime) != bits(literal):
            failures.append(f"{label}: imported {name} is {runtime!r}, not its "
                            f"literal {literal!r}")
        else:
            values[name] = runtime
    return values, failures


def check(root: Path, python: Path) -> list[str]:
    """Every failure; an empty list means each mirror equals its cap bit for bit."""
    vba, failures = vba_caps(root)
    mirrors, python_failures = python_caps(python)
    failures += python_failures
    for vba_name, python_name in PAIRS:
        if vba_name in vba and python_name in mirrors \
                and bits(vba[vba_name]) != bits(mirrors[python_name]):
            failures.append(
                f"{python_name} = {mirrors[python_name]!r} (bits {bits(mirrors[python_name])}) "
                f"differs from {vba_name} = {vba[vba_name]!r} (bits {bits(vba[vba_name])})")
    return failures


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("--root", type=Path, default=ROOT,
                        help="repository root holding src/ (default: this repository)")
    parser.add_argument("--python", type=Path, default=PYTHON,
                        help="Python module holding the mirrors (default: _contract_eval.py)")
    args = parser.parse_args(argv)
    failures = check(args.root, args.python)
    if failures:
        print("FAIL: df-cap mirrors")
        for failure in failures:
            print("  - " + failure)
        return 1
    print(f"PASS: df-cap mirrors ({len(PAIRS)} caps; each {SOURCE} literal converts "
          "exactly to its _contract_eval.py mirror, bit for bit)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
