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

It also pins every grid tool to the one shared definition, so a copy cannot be
reintroduced silently. Parity of VBA Val() with float() on these tokens needs
Excel and stays open in #32.

Run: python3 test_binary64.py   (exit 0 = pass, nonzero = fail)
"""
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

if fails:
    print("FAIL: binary64 argument helpers")
    for f in fails:
        print("  - " + f)
    raise SystemExit(1)
print(f"PASS: binary64 argument helpers ({sum(map(len, VALUES.values()))} edge values "
      "across ordinary, 17-digit, boundary, subnormal, seam, maximum and negative "
      "families; spelling, blank and unparsable keys; double-rounding token rule; "
      "committed main and holdout arguments; one shared definition in every tool)")
