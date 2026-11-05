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
    and constant the model mirrors in src/M_STATS_PROBDIST_NORMALFAMILY.bas,
    and of the model file itself. stale() compares them with the files as they
    are now, and also checks that the extraction is faithful: the Hart
    coefficients of PROB_NormalUpperTailPositive and the four constants must
    equal the literals in the model, and each route predicate of the VBA must
    have its counterpart in the model. Any difference is reported; an empty
    list means the model and its source are the ones finding F1 measured.

    A change to a bound procedure, constant or the model therefore fails
    test_binary64_route_model.py until the finding is re-measured, the study
    documents are updated, and the binding is re-recorded:

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

# Route predicates and statements: (VBA procedure, VBA statement, model function, model expression).
PREDICATES = (
    ("PROB_NormalUpperTailPositive", "If Z >= PROB_TAIL_UNDERFLOW_Z Then",
     "vba_upper_tail", "z >= TAIL_UNDERFLOW_Z"),
    ("PROB_NormalUpperTailPositive", "Exponential = Exp(-0.5 * Z * Z)",
     "vba_upper_tail", "math.exp(-0.5 * z * z)"),
    ("PROB_NormalUpperTailPositive", "If Z < PROB_CDF_SPLIT Then",
     "vba_upper_tail", "z < CDF_SPLIT"),
    ("PROB_NormalUpperTailPositive", "PROB_NormalUpperTailPositive = Exponential * SumA / SumB",
     "vba_upper_tail", "exponential * a / b"),
    ("PROB_NormalUpperTailPositive", "For TermIndex = PROB_NORMAL_TAIL_CF_TERMS To 1 Step -1",
     "vba_upper_tail", "range(NORMAL_TAIL_CF_TERMS, 0, -1)"),
    ("PROB_NormalUpperTailPositive", "Denominator = Z + CDbl(TermIndex) / Denominator",
     "vba_upper_tail", "z + float(k) / denominator"),
    ("PROB_NormalUpperTailPositive",
     "PROB_NormalUpperTailPositive = Exponential * PROB_INV_SQRT_TWO_PI / Denominator",
     "vba_upper_tail", "exponential * INV_SQRT_TWO_PI / denominator"),
    ("PROB_NormalSurvival", "If Z >= 0# Then", "vba_survival", "z >= 0.0"),
    ("PROB_NormalSurvival", "PROB_NormalSurvival = 1# - PROB_NormalUpperTailPositive(-Z)",
     "vba_survival", "1.0 - vba_upper_tail(-z)"),
    ("PROB_NormalCDF", "If Z < 0# Then", "vba_cdf", "z < 0.0"),
    ("PROB_NormalCDF", "PROB_NormalCDF = 1# - PROB_NormalUpperTailPositive(Z)",
     "vba_cdf", "1.0 - vba_upper_tail(z)"),
    ("PROB_NormalPDF", "PROB_NormalPDF = PROB_INV_SQRT_TWO_PI * Exp(-0.5 * Z * Z)",
     "route", "INV_SQRT_TWO_PI * math.exp(-0.5 * t * t)"),
)

RULE = ("SHA-256 of the text of each bound VBA procedure and constant declaration in the source "
        "(CRLF normalized to LF) and of the model file. A change to any of them makes finding F1 stale "
        "until it is re-measured, the study documents are updated, and model_fidelity.py --record "
        "--reason TEXT re-records the binding (#77).")


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
# VBA extraction
# ---------------------------------------------------------------------------

_LITERAL = r"\d+(?:\.\d*)?(?:[Ee][+-]?\d+)?#?"


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
    # The bound text has no string literals, so the first apostrophe opens a comment.
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


def _vba_number(token: str) -> float:
    return float(token.rstrip("#"))


def constant_value(declaration: str) -> float:
    """The Double a VBA constant declaration builds; a sum is evaluated left to right in Double."""
    expression = " ".join(statements(declaration)).split("=", 1)[1].strip()
    if not re.fullmatch(rf"{_LITERAL}(?:\s*\+\s*{_LITERAL})*", expression):
        raise ValueError(f"unsupported constant expression: {expression!r}")
    value = 0.0
    for i, token in enumerate(re.findall(_LITERAL, expression)):
        value = _vba_number(token) if i == 0 else value + _vba_number(token)
    return value


def hart_literals(procedure: str) -> dict[str, list[float]]:
    """The numeric literals of the SumA and SumB Horner statements, in source order."""
    out: dict[str, list[float]] = {"SumA": [], "SumB": []}
    for stmt in statements(procedure):
        target = stmt.split("=", 1)[0].strip()
        if target in out:
            out[target] += [_vba_number(t) for t in re.findall(rf"(?<![\w.]){_LITERAL}", stmt.split("=", 1)[1])]
    return out


# ---------------------------------------------------------------------------
# Model extraction (by parsing, so this module needs no mpmath)
# ---------------------------------------------------------------------------

def _evaluate(node: ast.expr) -> float:
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
        return float(node.value)
    if isinstance(node, ast.BinOp) and isinstance(node.op, ast.Add):
        return _evaluate(node.left) + _evaluate(node.right)
    raise ValueError(f"unsupported model constant expression: {ast.unparse(node)}")


def model_facts(model_source: str) -> dict[str, Any]:
    tree = ast.parse(model_source)
    constants: dict[str, float] = {}
    functions: dict[str, ast.FunctionDef] = {}
    for node in tree.body:
        if isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
            try:
                constants[node.targets[0].id] = _evaluate(node.value)
            except ValueError:
                pass
        elif isinstance(node, ast.FunctionDef):
            functions[node.name] = node
    hart: dict[str, list[float]] = {"a": [], "b": []}
    if "vba_upper_tail" in functions:
        for node in ast.walk(functions["vba_upper_tail"]):
            if isinstance(node, ast.Assign) and len(node.targets) == 1 \
                    and isinstance(node.targets[0], ast.Name) and node.targets[0].id in hart:
                literals = sorted((n.lineno, n.col_offset, float(n.value)) for n in ast.walk(node.value)
                                  if isinstance(n, ast.Constant) and isinstance(n.value, (int, float)))
                hart[node.targets[0].id] += [value for _, _, value in literals]
    expressions = {name: {ast.unparse(n) for n in ast.walk(fn) if isinstance(n, ast.expr)}
                   for name, fn in functions.items()}
    return {"constants": constants, "hart": hart, "expressions": expressions}


# ---------------------------------------------------------------------------
# Binding
# ---------------------------------------------------------------------------

def identities(root: str = ROOT) -> dict[str, Any]:
    """The identities of the bound source and model as they are under root now."""
    source = read_text(os.path.join(root, SOURCE))
    return {
        "procedures": {name: sha256_text(procedure_text(source, name)) for name in PROCEDURES},
        "constants": {name: sha256_text(constant_text(source, name)) for name, _ in CONSTANTS},
        "model_sha256": sha256_text(read_text(os.path.join(root, MODEL))),
    }


def extraction_problems(root: str = ROOT) -> list[str]:
    """Every way the model's numbers and predicates differ from the VBA source under root."""
    out: list[str] = []
    source = read_text(os.path.join(root, SOURCE))
    facts = model_facts(read_text(os.path.join(root, MODEL)))
    for vba_name, model_name in CONSTANTS:
        vba = constant_value(constant_text(source, vba_name))
        model = facts["constants"].get(model_name)
        if model != vba:
            out.append(f"constant {vba_name} = {vba!r} in VBA, but the model's {model_name} = {model!r}")
    hart = hart_literals(procedure_text(source, "PROB_NormalUpperTailPositive"))
    for vba_name, model_name in (("SumA", "a"), ("SumB", "b")):
        if hart[vba_name] != facts["hart"][model_name]:
            out.append(f"Hart {vba_name} coefficients {hart[vba_name]} in VBA, "
                       f"but the model's {model_name} has {facts['hart'][model_name]}")
    bodies = {name: statements(procedure_text(source, name)) for name in PROCEDURES}
    for procedure, statement, function, expression in PREDICATES:
        if statement not in bodies[procedure]:
            out.append(f"{procedure} no longer contains the statement '{statement}' that the model mirrors")
        if expression not in facts["expressions"].get(function, set()):
            out.append(f"model {function} no longer contains '{expression}' (mirrors '{statement}')")
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
    problems = extraction_problems(root)
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
    args = parser.parse_args(argv)
    if args.record:
        if not args.reason:
            parser.error("--record needs --reason")
        problems = record(args.reason)
        if problems:
            print("REFUSED: the model does not mirror the source:\n  - " + "\n  - ".join(problems))
            return 1
        print(f"recorded {os.path.relpath(RECORD, ROOT)}")
        return 0
    problems = stale(load_record())
    if problems:
        print("STALE: finding F1's model no longer matches its bound VBA source:\n  - " + "\n  - ".join(problems))
        return 1
    print(f"CURRENT: the model is bound to {len(PROCEDURES)} procedures and {len(CONSTANTS)} constants of {SOURCE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
