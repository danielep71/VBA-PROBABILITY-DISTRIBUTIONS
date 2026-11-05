"""
True binary64 ULP distance, and the binding of binary64_route_model.py to the
VBA source it mirrors (#77, finding F1 of PRODUCTION_FEASIBILITY.md).

ULP distance
    ulp_distance(a, b) is the number of representable binary64 steps between
    a and b: the difference of their ordered-integer bit patterns. Positive
    patterns keep their unsigned value; negative ones map to minus their
    magnitude bits, so the order is monotone across the sign. +0 and -0 are
    the same point; subnormals and binade boundaries are ordinary steps; the
    largest finite Double is one step from infinity. NaN has no distance and
    raises ValueError. This replaces |a - b| / (|a| * 2^-52), which is relative
    error in units of machine epsilon: below a power of two it understates the
    step count by up to almost 2x.

Source binding
    source_binding.json records the SHA-256 of the text of every VBA procedure
    and constant the model depends on in src/M_STATS_PROBDIST_NORMALFAMILY.bas,
    and of the model file itself. stale() compares them with the files as they
    are now, and also checks that the extraction is faithful:

    * Statement inventory. INVENTORY lists every executable statement of the
      four bound procedures, as the lexer below yields them (comments removed,
      continuations joined), in source order. Each is either CHECKED or carries
      the one-line reason it is intentionally unchecked. A statement in neither
      group, a listed statement that is gone, or a change of order is reported,
      so any new or changed statement forces a decision.
    * Translation. The CHECKED statements of each procedure are translated to
      Python (identifiers through NAMES, numeric literals to floats, VBA
      operators, signs and loop bounds kept exactly), normalized, and compared
      by syntax tree with the model function that mirrors it (TARGETS). A
      changed operator, sign, literal, loop bound, branch or statement order
      on either side is a difference.
    * Constants. The value expression of each mirrored VBA constant is compared
      the same way with the model constant; each guard constant must lie
      outside the domain the model evaluates, which is what makes its
      unchecked guard harmless.

    Any difference is reported; an empty list means the model and its source
    are the ones finding F1 measured. A change to a bound procedure, constant
    or the model therefore fails test_binary64_route_model.py until the
    finding is re-measured, the study documents and INVENTORY are updated,
    and the binding is re-recorded:

        python3 model_fidelity.py                      check (exit 1 if stale)
        python3 model_fidelity.py --record --reason T  re-record after review

    --record refuses while the extraction is not faithful, and appends the
    changed identities to the record's history.

The .bas text is hashed with CRLF normalized to LF, so a Git checkout and a VBE
export of the same source bind identically. Standard library only.
"""
from __future__ import annotations

import argparse
import ast
import copy
import datetime
import hashlib
import json
import math
import os
import re
import struct
import sys
from typing import Any

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
RECORD = os.path.join(HERE, "source_binding.json")

SOURCE = "src/M_STATS_PROBDIST_NORMALFAMILY.bas"
MODEL = "benchmark/student_t_large_df_study/binary64_route_model.py"

# The fidelity limit of the port against committed Excel observations, in true ULPs.
FIDELITY_LIMIT_ULPS = 2

# VBA procedures the model mirrors (PROB_NormalPDF: the route's phi).
PROCEDURES = ("PROB_NormalPDF", "PROB_NormalUpperTailPositive", "PROB_NormalCDF", "PROB_NormalSurvival")

# VBA constant -> the model constant that mirrors it.
CONSTANTS = (
    ("PROB_INV_SQRT_TWO_PI", "INV_SQRT_TWO_PI"),
    ("PROB_TAIL_UNDERFLOW_Z", "TAIL_UNDERFLOW_Z"),
    ("PROB_NORMAL_TAIL_CF_TERMS", "NORMAL_TAIL_CF_TERMS"),
    ("PROB_CDF_SPLIT", "CDF_SPLIT"),
)

# VBA guard constant -> the model constant bounding the domain the model evaluates.
# The guard's statements are intentionally unchecked; it is harmless only while it
# cannot fire inside that domain, so its value must exceed the bound.
GUARDS = (
    ("PROB_PDF_UNDERFLOW_Z", "T_MAX"),
)

BOUND_CONSTANTS = tuple(name for name, _ in CONSTANTS) + tuple(name for name, _ in GUARDS)

# VBA procedure -> (model function, model variable or None, model name of the parameter).
# With a variable, the procedure's result is compared with the one value assigned to it
# in the function; without, the procedure's body is compared with the function's body.
TARGETS = {
    "PROB_NormalPDF": ("route", "phi", "t"),
    "PROB_NormalUpperTailPositive": ("vba_upper_tail", None, "z"),
    "PROB_NormalCDF": ("vba_cdf", None, "z"),
    "PROB_NormalSurvival": ("vba_survival", None, "z"),
}

# VBA identifier -> model identifier, for every name a CHECKED statement may use.
# The parameter is mapped per procedure, from TARGETS.
NAMES = {
    "Exponential": "exponential",
    "SumA": "a",
    "SumB": "b",
    "Denominator": "denominator",
    "TermIndex": "k",
    "Exp": "math.exp",
    "CDbl": "float",
    "PROB_NormalUpperTailPositive": "vba_upper_tail",
    **{vba: model for vba, model in CONSTANTS},
}

CHECKED = None
PDF_GUARD = ("density underflow guard at |Z| >= PROB_PDF_UNDERFLOW_Z; the route evaluates phi only for "
             "|t| <= T_MAX, and the binding requires PROB_PDF_UNDERFLOW_Z > T_MAX")

# Every executable statement of the bound procedures, in source order: CHECKED, or the
# reason it is intentionally unchecked.
INVENTORY: dict[str, tuple[tuple[str, str | None], ...]] = {
    "PROB_NormalPDF": (
        ("Private Function PROB_NormalPDF( ByVal Z As Double) As Double", CHECKED),
        ("If Abs(Z) >= PROB_PDF_UNDERFLOW_Z Then", PDF_GUARD),
        ("PROB_NormalPDF = 0#", PDF_GUARD),
        ("Else", PDF_GUARD),
        ("PROB_NormalPDF = PROB_INV_SQRT_TWO_PI * Exp(-0.5 * Z * Z)", CHECKED),
        ("End If", PDF_GUARD),
        ("End Function", CHECKED),
    ),
    "PROB_NormalUpperTailPositive": (
        ("Private Function PROB_NormalUpperTailPositive( ByVal Z As Double) As Double", CHECKED),
        ("Dim Denominator As Double", CHECKED),
        ("Dim Exponential As Double", CHECKED),
        ("Dim SumA As Double", CHECKED),
        ("Dim SumB As Double", CHECKED),
        ("Dim TermIndex As Long", CHECKED),
        ("If Z >= PROB_TAIL_UNDERFLOW_Z Then", CHECKED),
        ("PROB_NormalUpperTailPositive = 0#", CHECKED),
        ("Exit Function", CHECKED),
        ("End If", CHECKED),
        ("Exponential = Exp(-0.5 * Z * Z)", CHECKED),
        ("If Z < PROB_CDF_SPLIT Then", CHECKED),
        ("SumA = 3.52624965998911E-02 * Z + 0.700383064443688", CHECKED),
        ("SumA = SumA * Z + 6.37396220353165", CHECKED),
        ("SumA = SumA * Z + 33.912866078383", CHECKED),
        ("SumA = SumA * Z + 112.079291497871", CHECKED),
        ("SumA = SumA * Z + 221.213596169931", CHECKED),
        ("SumA = SumA * Z + 220.206867912376", CHECKED),
        ("SumB = 8.83883476483184E-02 * Z + 1.75566716318264", CHECKED),
        ("SumB = SumB * Z + 16.064177579207", CHECKED),
        ("SumB = SumB * Z + 86.7807322029461", CHECKED),
        ("SumB = SumB * Z + 296.564248779674", CHECKED),
        ("SumB = SumB * Z + 637.333633378831", CHECKED),
        ("SumB = SumB * Z + 793.826512519948", CHECKED),
        ("SumB = SumB * Z + 440.413735824752", CHECKED),
        ("PROB_NormalUpperTailPositive = Exponential * SumA / SumB", CHECKED),
        ("Else", CHECKED),
        ("Denominator = Z", CHECKED),
        ("For TermIndex = PROB_NORMAL_TAIL_CF_TERMS To 1 Step -1", CHECKED),
        ("Denominator = Z + CDbl(TermIndex) / Denominator", CHECKED),
        ("Next TermIndex", CHECKED),
        ("PROB_NormalUpperTailPositive = Exponential * PROB_INV_SQRT_TWO_PI / Denominator", CHECKED),
        ("End If", CHECKED),
        ("End Function", CHECKED),
    ),
    "PROB_NormalCDF": (
        ("Private Function PROB_NormalCDF( ByVal Z As Double) As Double", CHECKED),
        ("If Z < 0# Then", CHECKED),
        ("PROB_NormalCDF = PROB_NormalUpperTailPositive(-Z)", CHECKED),
        ("Else", CHECKED),
        ("PROB_NormalCDF = 1# - PROB_NormalUpperTailPositive(Z)", CHECKED),
        ("End If", CHECKED),
        ("End Function", CHECKED),
    ),
    "PROB_NormalSurvival": (
        ("Private Function PROB_NormalSurvival( ByVal Z As Double) As Double", CHECKED),
        ("If Z >= 0# Then", CHECKED),
        ("PROB_NormalSurvival = PROB_NormalUpperTailPositive(Z)", CHECKED),
        ("Else", CHECKED),
        ("PROB_NormalSurvival = 1# - PROB_NormalUpperTailPositive(-Z)", CHECKED),
        ("End If", CHECKED),
        ("End Function", CHECKED),
    ),
}

RULE = ("SHA-256 of the text of each bound VBA procedure and constant declaration in the source "
        "(CRLF normalized to LF) and of the model file. A change to any of them makes finding F1 stale "
        "until it is re-measured, the study documents and the statement inventory are updated, and "
        "model_fidelity.py --record --reason TEXT re-records the binding (#77).")


# ---------------------------------------------------------------------------
# ULP distance
# ---------------------------------------------------------------------------

_MAGNITUDE = (1 << 63) - 1


def ordered(x: float) -> int:
    """The ordered-integer image of x: monotone in x, +0 and -0 both map to 0."""
    if math.isnan(x):
        raise ValueError("NaN has no position in the binary64 order")
    bits = struct.unpack("<Q", struct.pack("<d", x))[0]
    return -(bits & _MAGNITUDE) if bits >> 63 else bits


def ulp_distance(a: float, b: float) -> int:
    """Number of representable binary64 steps between a and b."""
    return abs(ordered(a) - ordered(b))


def within_ulps(observed: float, modelled: float, limit: int = FIDELITY_LIMIT_ULPS) -> bool:
    return ulp_distance(observed, modelled) <= limit


# ---------------------------------------------------------------------------
# VBA lexing
# ---------------------------------------------------------------------------

def read_text(path: str) -> str:
    with open(path, "rb") as f:
        return f.read().replace(b"\r\n", b"\n").decode("latin-1")


def sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("latin-1")).hexdigest()


def procedure_text(source: str, name: str) -> str:
    """The text of procedure `name`, from its declaration line to its End line."""
    starts = [m.start() for m in re.finditer(
        r"^[ \t]*(?:(?:Public|Private|Friend)[ \t]+)?(?:Static[ \t]+)?(?:Function|Sub|Property[ \t]+(?:Get|Let|Set))"
        rf"[ \t]+{re.escape(name)}[ \t]*\(", source, re.M)]
    if len(starts) != 1:
        raise LookupError(f"procedure {name}: found {len(starts)} declarations, expected 1")
    end = re.compile(r"^[ \t]*End[ \t]+(?:Function|Sub|Property)[ \t]*$", re.M).search(source, starts[0])
    if end is None:
        raise LookupError(f"procedure {name}: no End line")
    return source[starts[0]:end.end()]


def constant_text(source: str, name: str) -> str:
    """The declaration of constant `name`, with its continuation lines."""
    matches = list(re.finditer(
        rf"^[ \t]*(?:(?:Public|Private|Global)[ \t]+)?Const[ \t]+{re.escape(name)}\b.*$", source, re.M))
    if len(matches) != 1:
        raise LookupError(f"constant {name}: found {len(matches)} declarations, expected 1")
    lines = [matches[0].group(0)]
    rest = source[matches[0].end() + 1:].split("\n")
    while lines[-1].rstrip().endswith(" _") and rest:
        lines.append(rest.pop(0))
    return "\n".join(lines)


def _strip_comment(line: str) -> str:
    # A string literal would make this split wrong, but the translator rejects any
    # statement containing a quote, so such a statement can never pass as CHECKED.
    return line.split("'", 1)[0]


def statements(text: str) -> list[str]:
    """Executable lines of VBA text: comments removed, continuations joined, blanks collapsed."""
    out: list[str] = []
    pending = ""
    for line in text.split("\n"):
        code = _strip_comment(line).rstrip()
        if code.endswith(" _"):
            pending += code[:-2] + " "
            continue
        joined = " ".join((pending + code).split())
        pending = ""
        if joined:
            out.append(joined)
    return out


# ---------------------------------------------------------------------------
# VBA -> Python translation
# ---------------------------------------------------------------------------

_TOKEN = re.compile(r"\s*(?:(?P<num>\d+(?:\.\d*)?(?:[Ee][+-]?\d+)?#?)|(?P<name>[A-Za-z_]\w*)"
                    r"|(?P<op><=|>=|<>|[-+*/(),<>=]))")
_RESULT = "__result__"


class TranslationError(ValueError):
    pass


def vba_expression(text: str, names: dict[str, str], condition: bool = False) -> str:
    """Python text with exactly the semantics of a VBA arithmetic or comparison expression."""
    out: list[str] = []
    pos = 0
    text = text.strip()
    while pos < len(text):
        m = _TOKEN.match(text, pos)
        if m is None or m.end() == pos:
            raise TranslationError(f"unsupported VBA syntax at {text[pos:]!r}")
        pos = m.end()
        if m.group("num"):
            out.append(repr(float(m.group("num").rstrip("#"))))
        elif m.group("name"):
            if m.group("name") not in names:
                raise TranslationError(f"identifier {m.group('name')} has no model counterpart")
            out.append(names[m.group("name")])
        else:
            op = m.group("op")
            if op in ("=", "<>"):
                if not condition:
                    raise TranslationError(f"comparison {op!r} outside a condition in {text!r}")
                op = "==" if op == "=" else "!="
            out.append(op)
    return " ".join(out)


def translate(procedure: str, lines: list[str], param: str) -> list[ast.stmt]:
    """The CHECKED statements of a VBA Function as the body of an equivalent Python function."""
    names = dict(NAMES)
    if not lines:
        raise TranslationError(f"{procedure}: no checked statements")
    header = re.fullmatch(rf"(?:Public |Private )?Function {re.escape(procedure)}\( ?ByVal (\w+) As Double\) "
                          r"As Double", lines[0])
    if header is None:
        raise TranslationError(f"{procedure}: declaration {lines[0]!r} is not Function(ByVal X As Double) As Double")
    names[header.group(1)] = param
    if lines[-1] != "End Function":
        raise TranslationError(f"{procedure}: does not end with End Function")
    counters = {m.group(1) for line in lines if (m := re.match(r"For (\w+) = ", line))}
    out: list[str] = []
    blocks: list[tuple[str, int]] = []          # (kind, len(out) when the block opened)
    indent = 1

    def emit(code: str) -> None:
        out.append("    " * indent + code)

    def close(kind: str) -> None:
        nonlocal indent
        if not blocks or blocks[-1][0] != kind:
            raise TranslationError(f"{procedure}: unbalanced {kind} block")
        if len(out) == blocks.pop()[1]:
            emit("pass")
        indent -= 1

    for line in lines[1:-1]:
        if m := re.fullmatch(r"Dim (\w+) As (Double|Long)", line):
            if m.group(2) == "Long" and m.group(1) not in counters:
                raise TranslationError(f"{procedure}: {line!r}: a Long is modelled only as a loop counter")
            continue
        if m := re.fullmatch(r"If (.+) Then", line):
            emit(f"if {vba_expression(m.group(1), names, condition=True)}:")
            blocks.append(("If", len(out)))
            indent += 1
        elif line == "Else":
            close("If")
            emit("else:")
            blocks.append(("If", len(out)))
            indent += 1
        elif line == "End If":
            close("If")
        elif m := re.fullmatch(r"For (\w+) = (.+) To (.+?)(?: Step (.+))?", line):
            var, start, stop, step = m.group(1), m.group(2), m.group(3), m.group(4) or "1"
            if var not in names:
                raise TranslationError(f"{procedure}: loop counter {var} has no model counterpart")
            step_value = ast.literal_eval(vba_expression(step, names))
            if not isinstance(step_value, float) or step_value == 0 or step_value != int(step_value):
                raise TranslationError(f"{procedure}: {line!r}: the step must be a nonzero integer literal")
            # VBA's To bound is inclusive; range's stop is exclusive.
            last = vba_expression(stop, names)
            try:
                stop_py = repr(ast.literal_eval(last) + (1.0 if step_value > 0 else -1.0))
            except ValueError:
                stop_py = f"({last}) {'+' if step_value > 0 else '-'} 1.0"
            emit(f"for {names[var]} in range({vba_expression(start, names)}, {stop_py}, "
                 f"{vba_expression(step, names)}):")
            blocks.append((f"For {var}", len(out)))
            indent += 1
        elif m := re.fullmatch(r"Next (\w+)", line):
            close(f"For {m.group(1)}")
        elif line == "Exit Function":
            emit(f"return {_RESULT}")
        elif m := re.fullmatch(r"(\w+) = (.+)", line):
            target = _RESULT if m.group(1) == procedure else names.get(m.group(1))
            if target is None or "." in target:
                raise TranslationError(f"{procedure}: assignment target {m.group(1)} has no model counterpart")
            emit(f"{target} = {vba_expression(m.group(2), names)}")
        else:
            raise TranslationError(f"{procedure}: unsupported statement {line!r}")
    if blocks:
        raise TranslationError(f"{procedure}: unclosed {blocks[-1][0]} block")
    emit(f"return {_RESULT}")
    tree = ast.parse("def f():\n" + "\n".join(out) + "\n")
    function = tree.body[0]
    assert isinstance(function, ast.FunctionDef)
    return function.body


# ---------------------------------------------------------------------------
# Normalization: VBA's assign-the-name-then-exit and Python's return
# ---------------------------------------------------------------------------

class _Floats(ast.NodeTransformer):
    def visit_Constant(self, node: ast.Constant) -> ast.Constant:
        if isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
            return ast.copy_location(ast.Constant(float(node.value)), node)
        return node


def _is_result(node: ast.AST | None) -> bool:
    return isinstance(node, ast.Name) and node.id == _RESULT


def _canon(body: list[ast.stmt]) -> list[ast.stmt]:
    body = [copy.deepcopy(s) for s in body]
    for s in body:
        if isinstance(s, ast.If):
            s.body, s.orelse = _canon(s.body), _canon(s.orelse)
        elif isinstance(s, ast.For):
            s.body = _canon(s.body)
    changed = True
    while changed:
        changed = False
        # A return after an if/else runs at the end of both branches.
        if len(body) >= 2 and isinstance(body[-1], ast.Return) and isinstance(body[-2], ast.If) \
                and body[-2].orelse and not isinstance(body[-2].body[-1], ast.Return):
            tail = body.pop()
            branch = body[-1]
            assert isinstance(branch, ast.If)
            branch.body = _canon(branch.body + [copy.deepcopy(tail)])
            branch.orelse = _canon(branch.orelse + [copy.deepcopy(tail)])
            changed = True
        # Assigning the result and leaving is returning the value.
        for i in range(len(body) - 1):
            s, nxt = body[i], body[i + 1]
            if isinstance(s, ast.Assign) and len(s.targets) == 1 and _is_result(s.targets[0]) \
                    and isinstance(nxt, ast.Return) and _is_result(nxt.value):
                body[i:i + 2] = [ast.Return(s.value)]
                changed = True
                break
        # Statements after a return are unreachable.
        for i, s in enumerate(body[:-1]):
            if isinstance(s, ast.Return):
                del body[i + 1:]
                changed = True
                break
        for i, s in enumerate(body):
            if not isinstance(s, ast.If) or not s.orelse or not s.body or not isinstance(s.body[-1], ast.Return):
                continue
            if len(s.body) == 1 and len(s.orelse) == 1 and isinstance(s.orelse[0], ast.Return) \
                    and s.body[0].value is not None and s.orelse[0].value is not None:
                # if c: return x / else: return y  is  return x if c else y
                body[i:i + 1] = [ast.Return(ast.IfExp(s.test, s.body[0].value, s.orelse[0].value))]
            else:
                # An else after a returning branch is the code that follows the if.
                orelse, s.orelse = s.orelse, []
                body[i + 1:i + 1] = orelse
            changed = True
            break
    return body


def normalize(body: list[ast.stmt]) -> list[ast.stmt]:
    module = _Floats().visit(ast.Module(body=_canon(body), type_ignores=[]))
    return list(module.body)


def _dump(nodes: list[ast.stmt] | ast.expr) -> str:
    if isinstance(nodes, list):
        return "\n".join(ast.dump(n) for n in nodes)
    return ast.dump(nodes)


def _render(nodes: list[ast.stmt]) -> str:
    return ast.unparse(ast.Module(body=nodes, type_ignores=[])).replace("\n", "\n      ")


# ---------------------------------------------------------------------------
# Model extraction (by parsing, so this module needs no mpmath)
# ---------------------------------------------------------------------------

def _evaluate(node: ast.expr) -> float:
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
        return float(node.value)
    if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.USub):
        return -_evaluate(node.operand)
    if isinstance(node, ast.BinOp) and isinstance(node.op, (ast.Add, ast.Sub)):
        left, right = _evaluate(node.left), _evaluate(node.right)
        return left + right if isinstance(node.op, ast.Add) else left - right
    raise ValueError(f"unsupported constant expression: {ast.unparse(node)}")


def model_parts(model_source: str) -> tuple[dict[str, ast.expr], dict[str, ast.FunctionDef]]:
    tree = ast.parse(model_source)
    constants: dict[str, ast.expr] = {}
    functions: dict[str, ast.FunctionDef] = {}
    for node in tree.body:
        if isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
            constants[node.targets[0].id] = node.value
        elif isinstance(node, ast.FunctionDef):
            functions[node.name] = node
    return constants, functions


def _model_body(function: ast.FunctionDef) -> list[ast.stmt]:
    body = list(function.body)
    if body and isinstance(body[0], ast.Expr) and isinstance(body[0].value, ast.Constant) \
            and isinstance(body[0].value.value, str):
        body = body[1:]
    return body


# ---------------------------------------------------------------------------
# Binding
# ---------------------------------------------------------------------------

def inventory_counts() -> tuple[int, int]:
    """(CHECKED, intentionally unchecked) statements in INVENTORY."""
    entries = [reason for rows in INVENTORY.values() for _, reason in rows]
    return sum(r is CHECKED for r in entries), sum(r is not CHECKED for r in entries)


def inventory_problems(name: str, found: list[str]) -> list[str]:
    """Every way a procedure's statements differ from its inventory."""
    listed = [text for text, _ in INVENTORY.get(name, ())]
    out: list[str] = []
    remaining = list(listed)
    for text in found:
        if text in remaining:
            remaining.remove(text)
        else:
            out.append(f"{name}: statement '{text}' is neither cross-checked nor listed as intentionally "
                       "unchecked in INVENTORY")
    for text in remaining:
        out.append(f"{name}: INVENTORY lists '{text}', which the source no longer contains")
    if not out and found != listed:
        out.append(f"{name}: the statements are in a different order from INVENTORY")
    for text, reason in INVENTORY.get(name, ()):
        if reason is not CHECKED and not (isinstance(reason, str) and reason.strip()):
            out.append(f"{name}: '{text}' is unchecked without a reason")
    return out


def identities(root: str = ROOT) -> dict[str, Any]:
    """The identities of the bound source and model as they are under root now."""
    source = read_text(os.path.join(root, SOURCE))
    return {
        "procedures": {name: sha256_text(procedure_text(source, name)) for name in PROCEDURES},
        "constants": {name: sha256_text(constant_text(source, name)) for name in BOUND_CONSTANTS},
        "model_sha256": sha256_text(read_text(os.path.join(root, MODEL))),
    }


def _constant_expression(declaration: str) -> tuple[str, str]:
    """(declared type, value expression) of a VBA constant declaration."""
    m = re.fullmatch(r"(?:(?:Public|Private|Global) )?Const \w+ As (Double|Long) = (.+)",
                     " ".join(statements(declaration)))
    if m is None:
        raise TranslationError(f"unsupported constant declaration {declaration!r}")
    return m.group(1), m.group(2)


def extraction_problems(root: str = ROOT) -> list[str]:
    """Every way the model differs from the VBA source under root, statement by statement."""
    out: list[str] = []
    source = read_text(os.path.join(root, SOURCE))
    constants, functions = model_parts(read_text(os.path.join(root, MODEL)))

    for vba_name, model_name in CONSTANTS:
        kind, expression = _constant_expression(constant_text(source, vba_name))
        try:
            vba = ast.parse(vba_expression(expression, {}), mode="eval").body
        except TranslationError as e:
            out.append(f"constant {vba_name}: {e}")
            continue
        model = constants.get(model_name)
        if model is None:
            out.append(f"constant {vba_name}: the model has no {model_name}")
        elif _dump(_Floats().visit(copy.deepcopy(vba))) != _dump(_Floats().visit(copy.deepcopy(model))):
            out.append(f"constant {vba_name} = {expression} in VBA, but the model's {model_name} = "
                       f"{ast.unparse(model)}")
        elif kind == "Long" and not (isinstance(model, ast.Constant) and type(model.value) is int):
            out.append(f"constant {vba_name} is a Long, but the model's {model_name} is not an int literal")

    for vba_name, model_name in GUARDS:
        _, expression = _constant_expression(constant_text(source, vba_name))
        guard = _evaluate(ast.parse(vba_expression(expression, {}), mode="eval").body)
        bound = constants.get(model_name)
        bound_value = _evaluate(bound) if bound is not None else None
        if bound_value is None or not guard > bound_value:
            out.append(f"guard constant {vba_name} = {guard!r} must exceed the model's {model_name} = "
                       f"{bound_value!r}, or its unchecked guard can fire inside the modelled domain")

    for name in PROCEDURES:
        found = statements(procedure_text(source, name))
        problems = inventory_problems(name, found)
        out += problems
        if problems:
            continue
        function_name, variable, param = TARGETS[name]
        function = functions.get(function_name)
        if function is None:
            out.append(f"{name}: the model has no function {function_name}")
            continue
        if param not in [a.arg for a in function.args.args]:
            out.append(f"{name}: model {function_name} has no parameter {param}")
            continue
        checked = [text for text, reason in INVENTORY[name] if reason is CHECKED]
        try:
            vba = normalize(translate(name, checked, param))
        except (TranslationError, SyntaxError, ValueError) as e:
            out.append(f"{name}: cannot translate the checked statements: {e}")
            continue
        if variable is None:
            model = normalize(_model_body(function))
            if _dump(vba) != _dump(model):
                out.append(f"{name} does not match model {function_name}:\n    VBA:   {_render(vba)}"
                           f"\n    model: {_render(model)}")
            continue
        values = [n.value for n in ast.walk(function) if isinstance(n, ast.Assign)
                  and any(isinstance(t, ast.Name) and t.id == variable for t in n.targets)]
        if len(values) != 1:
            out.append(f"{name}: model {function_name} assigns {variable} {len(values)} times, expected once")
            continue
        model_value = _Floats().visit(copy.deepcopy(values[0]))
        if len(vba) != 1 or not isinstance(vba[0], ast.Return) or vba[0].value is None \
                or _dump(vba[0].value) != _dump(model_value):
            out.append(f"{name} does not match {variable} in model {function_name}:\n    VBA:   {_render(vba)}"
                       f"\n    model: {variable} = {ast.unparse(model_value)}")
    return out


def _compare(recorded: Any, now: dict[str, Any]) -> list[str]:
    out: list[str] = []
    if not isinstance(recorded, dict):
        return ["source_binding.json is missing or is not a JSON object"]
    for group in ("procedures", "constants"):
        rec = recorded.get(group) if isinstance(recorded.get(group), dict) else {}
        for name, sha in now[group].items():
            if rec.get(name) != sha:
                out.append(f"{SOURCE} {group[:-1]} {name}: recorded {rec.get(name)}, now {sha}")
        for name in sorted(set(rec) - set(now[group])):
            out.append(f"{SOURCE} {group[:-1]} {name}: recorded but no longer bound")
    if recorded.get("model_sha256") != now["model_sha256"]:
        out.append(f"{MODEL}: recorded {recorded.get('model_sha256')}, now {now['model_sha256']}")
    if recorded.get("source") != SOURCE or recorded.get("model") != MODEL:
        out.append("source_binding.json names a different source or model")
    return out


def stale(recorded: Any, root: str = ROOT) -> list[str]:
    """Every way the recorded binding differs from the files under root; [] if it is current."""
    try:
        now = identities(root)
        extraction = extraction_problems(root)
    except (OSError, LookupError, ValueError, SyntaxError) as e:
        return [f"cannot extract the bound source or model: {e}"]
    return _compare(recorded, now) + extraction


def load_record(path: str = RECORD) -> Any:
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return None


def record(reason: str, root: str = ROOT, path: str = RECORD, date: str | None = None) -> list[str]:
    """Re-record the binding; refuse (returning the problems) while the extraction is not faithful."""
    try:
        problems = extraction_problems(root)
    except (OSError, LookupError, ValueError, SyntaxError) as e:
        problems = [f"cannot extract the bound source or model: {e}"]
    if problems:
        return problems
    previous = load_record(path)
    now = identities(root)
    changed = _compare(previous, now) if previous is not None else ["initial binding"]
    history = list(previous.get("history", [])) if isinstance(previous, dict) else []
    history.append({"date": date or datetime.date.today().isoformat(), "reason": reason,
                    "identities_changed": changed})
    body = {"rule": RULE, "source": SOURCE, "model": MODEL,
            "evidence": ["benchmark/student_t_large_df_study/PRODUCTION_FEASIBILITY.md (finding F1)",
                         "benchmark/student_t_large_df_study/test_binary64_route_model.py"],
            **now, "history": history}
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        json.dump(body, f, indent=1)
        f.write("\n")
    return []


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--record", action="store_true", help="re-record the binding after review")
    parser.add_argument("--reason", help="why the binding is re-recorded (required with --record)")
    parser.add_argument("--date", help="date of the history entry (default: today)")
    args = parser.parse_args(argv)
    if args.record:
        if not args.reason:
            parser.error("--record needs --reason")
        problems = record(args.reason, date=args.date)
        if problems:
            print("REFUSED: the model does not mirror the source:\n  - " + "\n  - ".join(problems))
            return 1
        print(f"recorded {os.path.relpath(RECORD, ROOT)}")
        return 0
    problems = stale(load_record())
    if problems:
        print("STALE: finding F1's model no longer matches its bound VBA source:\n  - " + "\n  - ".join(problems))
        return 1
    checked, unchecked = inventory_counts()
    print(f"CURRENT: the model is bound to {len(PROCEDURES)} procedures and {len(BOUND_CONSTANTS)} constants of "
          f"{SOURCE}; {checked} statements cross-checked, {unchecked} intentionally unchecked")
    return 0


if __name__ == "__main__":
    sys.exit(main())
