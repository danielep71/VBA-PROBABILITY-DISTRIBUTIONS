"""
Source-fidelity fixture for the #72 kernel: the VBA that ships is the kernel
the study validated.

  1. Constants. Every PROB_NT_* constant in M_STATS_PROBDIST_SPECIALFUNCS
     evaluates, as VBA evaluates "a + b" or "a - b" in Double, to exactly the
     Double recorded in kernel_coefficients.json. No literal exceeds the 15
     significant digits the VBE keeps. The split, the continued-fraction depth
     and the underflow cut-off match the preregistered design, and the
     normal-family PROB_INV_SQRT_TWO_PI is the same Double.
  2. Arithmetic order. The Horner steps, the continued fraction, the final
     products and the split exponential are written in the order
     kernel_model.py evaluates them, and PROB_NormalPDF uses the split
     exponential.
  3. Wiring. PROB_NormalUpperTailPositive and its constants are gone; the
     cumulative, survival and interval kernels call PROB_NormalUpperTail.
  4. Test references. Every reference in Test_Core_ExpNegHalfSquare and in
     the upper-tail, density and lower-tail checks of
     Test_NF_TailRelativeAccuracy is the correctly rounded Double of the
     60-digit value. The binary64 model reproduces it within the test's own
     tolerance, so the VBA suite passes exactly when the VBA matches the model.

Run: python3 test_vba_kernel_source.py   (exit 0 = pass, nonzero = fail)
"""
import os
import re
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, HERE)
import kernel_model as km                                        # noqa: E402

fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


def source(rel):
    with open(os.path.join(ROOT, rel), "rb") as f:
        return f.read().decode("cp1252").replace("\r\n", "\n")


def joined(text):
    """Fold VBA line continuations (' _' at end of line) into single lines."""
    return re.sub(r" _\n\s*", " ", text)


def procedure(text, name):
    m = re.search(rf"^(Public|Private) (Function|Sub) {name}\(.*?^End (Function|Sub)$", text, re.S | re.M)
    check(m is not None, f"procedure {name} not found")
    return m.group(0) if m else ""


LITERAL = r"\d+(?:\.\d+)?(?:E[-+]\d{2,3})?#?"


def vba_constant(expr):
    """Evaluate 'a', '-a', 'a + b', 'a - b' (Double literals) as VBA does."""
    m = re.fullmatch(rf"\s*(-?)({LITERAL})(?:\s*([-+])\s*({LITERAL}))?\s*", expr)
    check(m is not None, f"unsupported constant expression {expr!r}")
    if not m:
        return None
    for lit in (m.group(2), m.group(4)):
        if lit:
            digits = lit.rstrip("#").split("E")[0].replace(".", "").lstrip("0")
            check(len(digits) <= 15, f"literal {lit} has more than 15 significant digits")
    value = float(m.group(2).rstrip("#"))
    if m.group(1):
        value = -value
    if m.group(3):
        other = float(m.group(4).rstrip("#"))
        value = value + other if m.group(3) == "+" else value - other
    return value


sf = joined(source("src/M_STATS_PROBDIST_SPECIALFUNCS.bas"))
core = joined(source("src/M_STATS_PROBDIST_CORE.bas"))
nf = joined(source("src/M_STATS_PROBDIST_NORMALFAMILY.bas"))
tests = joined(source("tests/M_STATS_PROBDIST_TEST.bas"))

# 1. Constants ---------------------------------------------------------------
consts = dict(re.findall(r"^Private Const (PROB_NT_\w+)\s+As (?:Double|Long) = ([^'\n]+?)\s*(?:'.*)?$", sf, re.M))
raw = km.COEFFICIENTS["raw"]
for i, rec in enumerate(raw["numerator"]):
    check(vba_constant(consts.get(f"PROB_NT_P{i}", "")) == float(rec["double"]),
          f"PROB_NT_P{i} is not the study's Double {rec['double']}")
for i, rec in enumerate(raw["denominator"][1:], start=1):
    check(vba_constant(consts.get(f"PROB_NT_D{i}", "")) == float(rec["double"]),
          f"PROB_NT_D{i} is not the study's Double {rec['double']}")
check(float(raw["denominator"][0]["double"]) == 1.0, "D(0) must be 1 for the literal 1# in the Horner step")
check(vba_constant(consts.get("PROB_NT_INV_SQRT_TWO_PI", "")) == float(raw["inv_sqrt_two_pi"]["double"]),
      "PROB_NT_INV_SQRT_TWO_PI is not the study's Double")
check(vba_constant(consts.get("PROB_NT_SPLIT", "")) == raw["split"], "PROB_NT_SPLIT differs from the design")
check(consts.get("PROB_NT_CF_LEVELS") == str(raw["continued_fraction_levels"]), "PROB_NT_CF_LEVELS differs from the design")
check(vba_constant(consts.get("PROB_NT_UNDERFLOW_Z", "")) == km.TAIL_UNDERFLOW_Z, "PROB_NT_UNDERFLOW_Z differs from the model")
nf_c = re.search(r"^Private Const PROB_INV_SQRT_TWO_PI As Double = ([^'\n]+)$", nf, re.M)
check(nf_c is not None and vba_constant(nf_c.group(1)) == float(raw["inv_sqrt_two_pi"]["double"]),
      "NORMALFAMILY PROB_INV_SQRT_TWO_PI is not the same Double as the kernel's")
pdf_z = re.search(r"^Private Const PROB_PDF_UNDERFLOW_Z As Double = ([\d.]+)", nf, re.M)
check(pdf_z is not None and float(pdf_z.group(1)) == km.PDF_UNDERFLOW_Z, "PROB_PDF_UNDERFLOW_Z differs from the model")

# 2. Arithmetic order --------------------------------------------------------
kernel = procedure(sf, "PROB_NormalUpperTail")
expected_lines = (
    ["If Z >= PROB_NT_UNDERFLOW_Z Then", "Exponential = PROB_ExpNegHalfSquare(Z)", "If Z < PROB_NT_SPLIT Then",
     "SumP = PROB_NT_P7 * Z + PROB_NT_P6"]
    + [f"SumP = SumP * Z + PROB_NT_P{i}" for i in range(5, -1, -1)]
    + ["SumD = PROB_NT_D7 * Z + PROB_NT_D6"]
    + [f"SumD = SumD * Z + PROB_NT_D{i}" for i in range(5, 0, -1)]
    + ["SumD = SumD * Z + 1#", "PROB_NormalUpperTail = Exponential * (SumP / SumD)", "Denominator = Z",
       "For Level = PROB_NT_CF_LEVELS To 1 Step -1", "Denominator = Z + CDbl(Level) / Denominator",
       "PROB_NormalUpperTail = Exponential * (PROB_NT_INV_SQRT_TWO_PI / Denominator)"])
code = [ln.strip() for ln in kernel.split("\n") if ln.strip() and not ln.strip().startswith("'")]
position = 0
for line in expected_lines:
    try:
        position = code.index(line, position) + 1
    except ValueError:
        fails.append(f"PROB_NormalUpperTail: expected, in order, {line!r}")
exp_helper = procedure(core, "PROB_ExpNegHalfSquare")
check("ZHigh = Int(Z * 256#) / 256#" in exp_helper, "PROB_ExpNegHalfSquare: split must be Int(Z * 256#) / 256#")
check("PROB_ExpNegHalfSquare = Exp(-0.5 * (ZHigh * ZHigh)) * Exp(-0.5 * ((Z - ZHigh) * (Z + ZHigh)))" in exp_helper,
      "PROB_ExpNegHalfSquare: product is not written in the modelled order")
check(float(raw["exponent_split_scale"]) == 256.0, "the study's exponent split scale is not 256")
pdf = procedure(nf, "PROB_NormalPDF")
check("PROB_NormalPDF = PROB_INV_SQRT_TWO_PI * PROB_ExpNegHalfSquare(Abs(Z))" in pdf,
      "PROB_NormalPDF does not use the split exponential as modelled")
check("If Abs(Z) >= PROB_PDF_UNDERFLOW_Z Then" in pdf, "PROB_NormalPDF lost its underflow guard")

# 3. Wiring ------------------------------------------------------------------
for gone in ("PROB_NormalUpperTailPositive", "PROB_TAIL_UNDERFLOW_Z As", "PROB_NORMAL_TAIL_CF_TERMS As",
             "PROB_CDF_SPLIT As"):
    check(not re.search(rf"^[^'\n]*{gone}", nf, re.M), f"{gone} still appears in NORMALFAMILY code")
check(procedure(nf, "PROB_NormalCDF").count("PROB_NormalUpperTail(") == 2, "PROB_NormalCDF must call PROB_NormalUpperTail twice")
check(procedure(nf, "PROB_NormalSurvival").count("PROB_NormalUpperTail(") == 2, "PROB_NormalSurvival must call PROB_NormalUpperTail twice")
check(procedure(nf, "PROB_NormalIntervalProbability").count("PROB_NormalUpperTail(") == 4,
      "PROB_NormalIntervalProbability must call PROB_NormalUpperTail four times")

# 4. Test references ---------------------------------------------------------
tol_m = re.search(r"^Private Const TOL_REL_NORMAL_TAIL As Double = ([\d.E-]+)", tests, re.M)
check(tol_m is not None, "TOL_REL_NORMAL_TAIL not declared")
tolerance = float(tol_m.group(1)) if tol_m else 0.0
check(tolerance <= 2.5 * km.BOUND, f"TOL_REL_NORMAL_TAIL {tolerance} is looser than about twice the 2^-50 bound")

with mp.workdps(60):
    references = {
        "PROB_ExpNegHalfSquare": (lambda z: mp.exp(-mp.mpf(z) ** 2 / 2), km.exp_neg_half_square),
        "K_STATS_NormalStandard_Survival": (lambda z: mp.erfc(mp.mpf(z) / mp.sqrt(2)) / 2, km.survival),
        "K_STATS_NormalStandard_Cumulative": (lambda z: mp.erfc(-mp.mpf(z) / mp.sqrt(2)) / 2, km.cdf),
        "K_STATS_NormalStandard_Density": (lambda z: mp.exp(-mp.mpf(z) ** 2 / 2) / mp.sqrt(2 * mp.pi), km.candidate_pdf),
    }
    body = procedure(tests, "Test_Core_ExpNegHalfSquare") + procedure(tests, "Test_NF_TailRelativeAccuracy")
    pattern = re.compile(r'AssertRelClose "([^"]+)", (\w+)\(([^)]*)\), ([^,]+), TOL_REL_NORMAL_TAIL')
    checked = 0
    for label, function, argument, expected in pattern.findall(body):
        if function not in references:
            continue
        z = vba_constant(argument)
        exact, model = references[function]
        target = float(exact(z))
        check(vba_constant(expected) == target, f"{label}: reference {expected} is not the correctly rounded value")
        predicted = model(z)
        rel = float(abs((mp.mpf(predicted) - exact(z)) / exact(z)))
        check(rel <= tolerance, f"{label}: the model is {rel:.2e} from the reference, outside the test tolerance")
        checked += 1
    check(checked >= 17, f"expected at least 17 kernel-level test references, parsed {checked}")

if fails:
    print("FAIL: #72 VBA kernel source fidelity")
    for msg in fails:
        print("  -", msg)
    sys.exit(1)
print(f"PASS: #72 VBA kernel source fidelity: {len(raw['numerator']) + len(raw['denominator'])} coefficients and 4 "
      f"design constants match the study exactly; arithmetic order, wiring and {checked} test references verified")
