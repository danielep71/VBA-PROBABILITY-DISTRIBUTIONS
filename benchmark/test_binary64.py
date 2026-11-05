"""
Proofs for the shared binary64 argument helpers (#32).

Python side of the parser-parity requirement: over ordinary decimals, 17-digit
round trips, the normal/subnormal boundary, representative subnormals, the
minimum subnormal, one-ulp seams and the maximum finite Double,

  * the emitted token parses back to exactly the Double it was emitted for;
  * as_reference_arg evaluates at that exact Double;
  * different spellings of one Double share one key;
  * the token grammar is one Val() also reads (sign, digits, '.', 'e' exponent);
  * every argument in the committed main and holdout grids parses, and
    re-emitting it reproduces the same Double.

It also pins every grid tool - reconcile_grid, promote_grid_rows,
migrate_references, check_promotion, report_promotion_candidates,
check_grid_coverage, audit_references and the generator - to the one shared
definition (#90), and shows on one fixture that they all key the edge rows
identically: non-finite tokens in Python's and in VBA's spellings, a blank
regime, -0 against +0 (two identities: the key is the bit pattern), and
subnormals. A static guard fails if a grid tool imports struct again, which is
how the private copies were written. Parity of VBA Val() with float() on these
tokens needs Excel and stays open in #32.

Run: python3 test_binary64.py   (exit 0 = pass, nonzero = fail)
"""
import ast
import importlib.util
import math
import os
import re
import struct
import sys

import mpmath as mp

import _binary64 as B
import _grid_safety as gs

HERE = os.path.dirname(os.path.abspath(__file__))
fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


def bits(x):
    return struct.pack(">d", x).hex()


def from_bits(h):
    return struct.unpack(">d", bytes.fromhex(h))[0]


MIN_SUBNORMAL = from_bits("0000000000000001")
MAX_SUBNORMAL = from_bits("000fffffffffffff")
MIN_NORMAL = from_bits("0010000000000000")
MAX_FINITE = from_bits("7fefffffffffffff")

VALUES = {
    "ordinary decimals": [0.1, 0.5, 0.85, 1.5, 2.0, 10.0, 123456.789, 1e-8, 1e22, 1e23],
    "17-digit round trips": [0.84999999999999998, 3.0702906297578495e-7,
                             0.30000000000000004, 2.0 / 3.0],
    "normal/subnormal boundary": [MIN_NORMAL, MAX_SUBNORMAL],
    "representative subnormals": [1e-310, 2.5e-320, from_bits("0000000080000000")],
    "minimum subnormal": [MIN_SUBNORMAL],
    "one-ulp seams": [math.nextafter(1.0, 2.0), math.nextafter(1.0, 0.0),
                      math.nextafter(2.0 ** 53, math.inf), math.nextafter(0.5, 0.0),
                      math.nextafter(MIN_NORMAL, 0.0)],
    "maximum finite": [MAX_FINITE, math.nextafter(MAX_FINITE, 0.0)],
    "integers above 2^52": [6235507341273914.0, 6235507341273916.0, 2.0 ** 53],
    "negatives": [-0.85, -MIN_SUBNORMAL, -MAX_FINITE, -1e-8],
}
TOKEN = re.compile(r"^-?[0-9]+(\.[0-9]+)?(e[+-]?[0-9]+)?$")

mp.mp.dps = 50
for family, values in VALUES.items():
    for v in values:
        token = B.format_double_arg(v)
        check(bool(TOKEN.fullmatch(token)),
              f"{family}: token {token!r} for {v!r} uses only Val()-readable characters")
        digits = re.sub(r"e.*$", "", token).replace("-", "").replace(".", "").lstrip("0")
        check(len(digits) <= 17, f"{family}: token {token!r} has more than 17 significant digits")
        check(B.double_key(token) == bits(v),
              f"{family}: token {token!r} parses back to exactly {v!r}")
        check(B.as_reference_arg(v) == mp.mpf(v),
              f"{family}: as_reference_arg({v!r}) is the exact Double")
        check(B.as_reference_arg(token) == mp.mpf(v),
              f"{family}: as_reference_arg({token!r}) is the Double the token denotes")
        check(B.format_double_arg(token) == token,
              f"{family}: re-emitting {token!r} is byte-stable")
        check(B.exact_binary64_mpf(v) == mp.mpf(v),
              f"{family}: exact_binary64_mpf({v!r}) is exact")

# Spellings of one Double share one key; blanks and garbage are distinct and marked.
check(B.double_key("0.85") == B.double_key("0.84999999999999998") == bits(0.85),
      "decimal spellings of one Double share a key")
check(B.double_key("1E-8") == B.double_key("1.0e-8") == B.double_key(" 1e-08 "),
      "exponent spellings and surrounding blanks share a key")
check(B.double_key("") == "" and B.double_key(None) == "" and B.double_key("  ") == "",
      "blank cells have the empty key")
check(B.double_key("abc") == "?abc" and B.double_key("1e999") == "?1e999",
      "unparsable or overflowing tokens are marked, never coerced")
check(B.double_key("4.9406564584124654e-324") == "0000000000000001",
      "the minimum subnormal token keys to bit pattern 1")

# The documented double-rounding case: follow the token, not the design value.
design = mp.mpf("6235507341273914.5000001")
check(B.format_double_arg(design) == "6235507341273914.0"
      and B.as_reference_arg(design) == mp.mpf(6235507341273914),
      "a design value just above a tie follows its 17-digit token to 6235507341273914")
check(float(design) == 6235507341273915.0,
      "the nearest Double to that design value is the other neighbour")
check(B.double_key("6235507341273914.5") == bits(6235507341273914.0),
      "the committed tie token rounds half to even (6235507341273914)")
check(B.double_key("inf") == "?inf" and B.double_key("nan") == "?nan"
      and B.double_key("-1e400") == "?-1e400",
      "non-finite tokens are marked, because Val() never yields them")

# Count inputs pass through untouched; sequences are canonicalised elementwise.
check(B.as_reference_arg(7) == 7 and isinstance(B.as_reference_arg(7), int),
      "integer count inputs pass through as int")
pair = B.as_reference_arg((mp.mpf("0.85"), 3))
check(isinstance(pair, tuple) and pair[0] == mp.mpf(0.85) and pair[1] == 3,
      "tuples are canonicalised element by element")

# Every committed argument parses and is reproduced by re-emission.
for grid in (gs.COMMITTED_GRID, gs.COMMITTED_HOLDOUT):
    rows = gs.read_grid(grid)
    unparsable = []
    unstable = []
    for row in rows:
        for column in ("arg1", "arg2", "arg3", "arg4"):
            text = (row.get(column) or "").strip()
            if not text:
                continue
            key = B.double_key(text)
            if key.startswith("?"):
                unparsable.append(text)
            elif B.double_key(B.format_double_arg(mp.mpf(float(text)))) != key:
                unstable.append(text)
    name = os.path.relpath(grid, HERE)
    check(not unparsable, f"{name}: every argument parses ({unparsable[:3]})")
    check(not unstable, f"{name}: every argument survives re-emission ({unstable[:3]})")

# One shared definition: each tool binds the shared helpers, not a copy.


def load(name, filename):
    spec = importlib.util.spec_from_file_location(name, os.path.join(HERE, filename))
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


check(gs.f64_bits is B.double_key, "_grid_safety.f64_bits is the shared double_key")
for filename in ("reconcile_grid.py", "promote_grid_rows.py", "migrate_references.py",
                 "check_promotion.py"):
    module = load("pin_" + filename[:-3], filename)
    check(module.key is gs.canonical_key, f"{filename} key is _grid_safety.canonical_key")
check(load("pin_reconcile_f64", "reconcile_grid.py").f64 is B.double_key,
      "reconcile_grid.py f64 is the shared double_key")
check(load("pin_candidates", "report_promotion_candidates.py").bits is B.double_key,
      "report_promotion_candidates.py bits is the shared double_key")
generator = load("pin_generator", "generate_reference_values.py")
check(generator.D is B.as_reference_arg and generator.format_double_arg is B.format_double_arg,
      "generate_reference_values.py uses the shared as_reference_arg and format_double_arg")
coverage = load("pin_coverage", "check_grid_coverage.py")
check(getattr(coverage, "canonical_key", None) is gs.canonical_key,
      "check_grid_coverage.py builds its identity from _grid_safety.canonical_key")
audit = load("pin_audit", "audit_references.py")
check(getattr(audit, "key", None) is gs.canonical_key
      and getattr(audit, "bits", None) is B.double_key,
      "audit_references.py key and bits are canonical_key and double_key")

# One fixture, keyed identically by every tool (#90). Before #90 the coverage
# gate raised on a non-finite token and read a blank regime as 'all', and the
# audit gave "inf" the pattern 7ff0000000000000; double_key marks it instead.
# Signed zeros: the identity is the bit pattern, so "-0" and "0" are two
# identities everywhere, exactly as double_key keys them.
NONFINITE = ["inf", "-inf", "+inf", "Infinity", "nan", "NaN", "-nan", "1e999",
             "-1e400",                                    # Python reads these
             "1.#INF", "-1.#INF", "1.#QNAN", "-1.#IND"]   # VBA's own spellings
ZEROS = {"0": "0000000000000000", "0.0": "0000000000000000",
         "1e-400": "0000000000000000",                    # underflows to +0
         "-0": "8000000000000000", "-0.0": "8000000000000000",
         "-1e-400": "8000000000000000"}                   # underflows to -0
SUBNORMALS = {"4.9406564584124654e-324": "0000000000000001",
              "5e-324": "0000000000000001",
              "-5e-324": "8000000000000001",
              "2.2250738585072009e-308": "000fffffffffffff",
              "2.2250738585072014e-308": "0010000000000000",  # minimum normal
              "1e-310": bits(1e-310)}
for token in NONFINITE:
    check(B.double_key(token) == "?" + token, f"non-finite token {token!r} is marked")
for token, pattern in {**ZEROS, **SUBNORMALS}.items():
    check(B.double_key(token) == pattern, f"{token!r} keys to {pattern}")
check(B.double_key("-0") != B.double_key("0"), "-0 and +0 are distinct identities")


def fixture_row(arg1, regime="all"):
    return {"function": "Fixture_Edge", "vba_kernel": "", "claim": "", "metric": "",
            "arg1": arg1, "arg2": "2", "arg3": "", "arg4": "", "reference": "",
            "observed_vba": "", "regime": regime, "evidence_set": "main grid",
            "expected_error": ""}


edge_rows = ([fixture_row(t) for t in NONFINITE]
             + [fixture_row(t) for t in {**ZEROS, **SUBNORMALS}]
             + [fixture_row("0.85", regime=""), fixture_row("inf", regime="")])
tool_keys = {
    "_grid_safety.canonical_key": gs.canonical_key,
    "check_grid_coverage.row_identity": lambda r: coverage.identity_key(coverage.row_identity(r)),
    "audit_references.key": getattr(audit, "key", lambda r: "no module-level key"),
}
for filename in ("reconcile_grid.py", "promote_grid_rows.py", "migrate_references.py",
                 "check_promotion.py"):
    tool_keys[filename + " key"] = load("edge_" + filename[:-3], filename).key
for r in edge_rows:
    expected = ("Fixture_Edge", B.double_key(r["arg1"]), B.double_key("2"), "", "",
                r["regime"], "main grid")
    for name, fn in tool_keys.items():
        try:
            got = fn(r)
        except Exception as exc:  # a tool that cannot key the row disagrees
            got = f"{type(exc).__name__}: {exc}"
        check(got == expected,
              f"{name} keys arg1={r['arg1']!r} regime={r['regime']!r} as {expected}, not {got}")
check(gs.canonical_key(fixture_row("0.85", regime="")) != gs.canonical_key(fixture_row("0.85")),
      "a blank regime is its own identity, not an alias of 'all'")
check(gs.canonical_key(fixture_row("-0")) != gs.canonical_key(fixture_row("0")),
      "rows at -0 and +0 have distinct identities")

# The coverage gate keys those rows the same way, then refuses the ones no
# disposition may carry, instead of giving them an identity of its own.
refused = coverage.evaluate_coverage(
    [fixture_row("inf"), fixture_row("1.#INF"), fixture_row("0.85", regime="")], [], [])
check(refused["main_rows"] == 0 and len(refused["errors"]) == 3
      and sum("non-finite or unparsable" in e for e in refused["errors"]) == 2
      and sum("blank function or regime" in e for e in refused["errors"]) == 1,
      f"the coverage gate refuses non-finite and blank-regime rows ({refused['errors']})")
signed = coverage.evaluate_coverage(
    [fixture_row("-0"), fixture_row("0"), fixture_row("5e-324"),
     fixture_row("4.9406564584124654e-324")], [], [])
check(signed["main_rows"] == 4 and len(signed["main_identities"]) == 3,
      "the coverage gate counts -0 and +0 apart and one subnormal's spellings together")

# Guard: no grid tool keeps a private bit-pattern helper again. Every
# benchmark script that reads grid argument cells (it names arg1), except the
# shared module and the tests, must not import struct - the private copies
# were all struct.pack(">d", ...).
scanned = []
for folder in (HERE, os.path.join(HERE, "holdout")):
    for name in sorted(os.listdir(folder)):
        path = os.path.join(folder, name)
        if not name.endswith(".py") or name.startswith("test_") or name == "_binary64.py":
            continue
        with open(path, encoding="utf-8") as fh:
            source = fh.read()
        if "arg1" not in source:
            continue
        scanned.append(os.path.relpath(path, HERE))
        imports_struct = any(
            (isinstance(node, ast.Import) and any(a.name == "struct" for a in node.names))
            or (isinstance(node, ast.ImportFrom) and node.module == "struct")
            for node in ast.walk(ast.parse(source, filename=path)))
        check(not imports_struct,
              f"{scanned[-1]} imports struct; use _binary64.double_key for argument identity")
check({"check_grid_coverage.py", "audit_references.py", "reconcile_grid.py",
       "generate_reference_values.py", "_grid_safety.py"} <= set(scanned),
      f"the struct guard covers the grid tools ({scanned})")

if fails:
    print("FAIL: binary64 argument helpers")
    for f in fails:
        print("  - " + f)
    raise SystemExit(1)
print(f"PASS: binary64 argument helpers ({sum(map(len, VALUES.values()))} edge values "
      "across ordinary, 17-digit, boundary, subnormal, seam, maximum and negative "
      "families; spelling, blank and unparsable keys; double-rounding token rule; "
      "committed main and holdout arguments; one shared definition in every tool, "
      f"edge rows keyed identically by {len(tool_keys)} tools, struct guard over "
      f"{len(scanned)} grid scripts)")
