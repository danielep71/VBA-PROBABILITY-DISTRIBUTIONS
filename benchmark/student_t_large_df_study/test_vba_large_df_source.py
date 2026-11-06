"""
Source-fidelity fixture for the Student-t large-df route (#34): the VBA that
ships is the route PRODUCTION_PREREGISTRATION.md froze and
production_route_model.py measures.

  1. Constants. Every PROB_T_LDF_Gk_j constant in M_STATS_PROBDIST_TFAMILY
     evaluates, as VBA evaluates "a + b" or "a - b" in Double, to exactly the
     coefficient Double of the model, with no literal above 15 significant
     digits, and no coefficient is missing or extra. The predicate bounds are
     1E5 and 8, and PROB_T_INV_SQRT_TWO_PI is the normal kernel's Double.
  2. Arithmetic. The executable statements of PROB_TF_StudentTLargeDFTail are
     exactly the ones §2 prescribes, in that order.
  3. Wiring. PROB_TryStudentTTail dispatches to the route on the preregistered
     predicate after the df = 2 branch and before the incomplete-beta route.
  4. Test references. Every reference in Test_TF_StudentTLargeDFTail is the
     correctly rounded Double of the 60-digit oracle at its arguments; the
     model reproduces each forward reference inside the region within the
     test's tolerance, and each inverse reference's PT2 bound is below it.
  5. Mutations. A changed last digit of a coefficient, a reassociated product,
     a dropped order and a moved predicate bound each fail checks 1-3.
  6. Inverse polish (INVERSE_PREREGISTRATION.md section 2). Its two constants,
     its statements in the order of section 1, and their place after the
     convergence check and before Result = X. A different step count, a
     different tail bound, a missing XNew = X exit and the polish moved
     before the convergence check each fail it.
  7. Quantile references. Every reference in Test_TF_StudentTLargeDFQuantile
     is the correctly rounded quantile; every outcome predict_inverse.py
     allows there is within the test's tolerance, and the tolerance is below
     the first wave's error at that row (excel_wave_check.json).

Run: python3 test_vba_large_df_source.py   (exit 0 = pass)
"""
import json
import os
import re
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, HERE)
import predict_production as PP  # noqa: E402
import predict_inverse as PI  # noqa: E402
import production_route_model as P  # noqa: E402
import validate_expansion as V  # noqa: E402

SOURCE = "src/M_STATS_PROBDIST_TFAMILY.bas"
TESTS = "tests/M_STATS_PROBDIST_TEST.bas"
LITERAL = r"\d+(?:\.\d+)?(?:E[-+]\d{2,3})?#?"
TOLERANCES = {"TOL_REL_LDF_QUANTILE": 5e-16, "TOL_REL_LARGE_DF_TAIL": 2e-15, "TOL_REL_LDF_OUTSIDE_DF": 5e-12, "TOL_REL_TAIL": 1e-9}


def read(rel):
    with open(os.path.join(ROOT, rel), "rb") as f:
        return f.read().decode("cp1252").replace("\r\n", "\n")


def joined(text):
    return re.sub(r" _\n\s*", " ", text)


def procedure(text, name):
    m = re.search(rf"^(?:Public|Private) (?:Function|Sub) {name}\(.*?^End (?:Function|Sub)$",
                  text, re.S | re.M)
    return m.group(0) if m else None


def statements(proc):
    out = []
    for line in joined(proc).split("\n"):
        code = line.split("'", 1)[0].strip()
        if code and not code.startswith("Dim "):
            out.append(code)
    return out


def vba_value(expr, problems, name):
    m = re.fullmatch(rf"\s*(-?)({LITERAL})(?:\s*([-+])\s*({LITERAL}))?\s*", expr)
    if not m:
        problems.append(f"{name}: unsupported constant expression {expr!r}")
        return None
    for lit in (m.group(2), m.group(4)):
        if lit and len(lit.rstrip("#").split("E")[0].replace(".", "").lstrip("0")) > 15:
            problems.append(f"{name}: literal {lit} has more than 15 significant digits")
    value = float(m.group(2).rstrip("#"))
    if m.group(1):
        value = -value
    if m.group(3):
        other = float(m.group(4).rstrip("#"))
        value = value + other if m.group(3) == "+" else value - other
    return value


def expected_body():
    lines = ["U = AbsX * AbsX", "W = 1# / DegreesFreedom"]
    for k in range(P.ORDER, 0, -1):
        top = len(P.COEFFICIENTS[k - 1]) - 1
        lines.append(f"Horner = PROB_T_LDF_G{k}_{top}")
        lines += [f"Horner = Horner * U + PROB_T_LDF_G{k}_{j}" for j in range(top - 1, -1, -1)]
        lines.append("Series = Horner" if k == P.ORDER else "Series = Series * W + Horner")
    lines += ["Series = Series * W",
              "Correction = AbsX * Series",
              "Correction = PROB_T_INV_SQRT_TWO_PI * PROB_ExpNegHalfSquare(AbsX) * Correction",
              "PROB_TF_StudentTLargeDFTail = PROB_NormalUpperTail(AbsX) + Correction"]
    return lines


def source_problems(text):
    """Checks 1-3 on a module text; [] when the VBA is the preregistered route."""
    problems = []
    flat = joined(text)
    consts = dict(re.findall(r"^Private Const (PROB_T_\w+)\s+As (?:Double|Long) = ([^'\n]+)", flat, re.M))
    expected = {f"PROB_T_LDF_G{k}_{j}": d for k, row in enumerate(P.COEFFICIENTS, start=1)
                for j, (_, d, _, _) in enumerate(row)}
    found = {n for n in consts if re.fullmatch(r"PROB_T_LDF_G\d+_\d+", n)}
    for name in sorted(set(expected) - found):
        problems.append(f"{name} is missing")
    for name in sorted(found - set(expected)):
        problems.append(f"{name} is not a coefficient of the route")
    for name, double in expected.items():
        if name in consts:
            value = vba_value(consts[name], problems, name)
            if value is not None and value != double:
                problems.append(f"{name} = {value!r}, the model's Double is {double!r}")
    for name, want in (("PROB_T_LDF_MIN_DF", P.DF_MIN), ("PROB_T_LDF_MAX_ABS_T", P.T_MAX),
                       ("PROB_T_INV_SQRT_TWO_PI", P.INV_SQRT_TWO_PI)):
        value = vba_value(consts.get(name, "x"), problems, name) if name in consts else None
        if value != want:
            problems.append(f"{name} is {value!r}, expected {want!r}")
    route = procedure(text, "PROB_TF_StudentTLargeDFTail")
    if route is None:
        problems.append("PROB_TF_StudentTLargeDFTail not found")
    else:
        body = statements(route)[1:-1]        # drop the declaration and End Function
        if body != expected_body():
            for i, (got, want) in enumerate(zip(body + [""] * 99, expected_body() + [""] * 99)):
                if got != want:
                    problems.append(f"PROB_TF_StudentTLargeDFTail statement {i + 1}: "
                                    f"{got!r}, expected {want!r}")
                    break
    tail = procedure(text, "PROB_TryStudentTTail")
    if tail is None:
        problems.append("PROB_TryStudentTTail not found")
    else:
        lines = statements(tail)
        guard = "If DegreesFreedom >= PROB_T_LDF_MIN_DF And AbsX <= PROB_T_LDF_MAX_ABS_T Then"
        if lines.count(guard) != 1:
            problems.append("PROB_TryStudentTTail does not dispatch on the preregistered predicate")
        else:
            i = lines.index(guard)
            if lines[i + 1] != "Result = PROB_TF_StudentTLargeDFTail(AbsX, DegreesFreedom)":
                problems.append("the large-df branch does not return the route")
            df2 = lines.index("If DegreesFreedom = 2# Then") if "If DegreesFreedom = 2# Then" in lines else -1
            beta = next((j for j, s in enumerate(lines) if "PROB_TF_TrySquareRatioPair" in s), -1)
            if not df2 < i < beta:
                problems.append("the large-df branch is not between the df = 2 branch and the beta route")
    return problems


def reference_problems(text):
    """Check 4 on Test_TF_StudentTLargeDFTail."""
    problems = []
    sub = procedure(text, "Test_TF_StudentTLargeDFTail")
    if sub is None:
        return ["Test_TF_StudentTLargeDFTail not found"]
    pattern = (r'AssertRelClose "([^"]+)", (K_STATS_StudentT_(\w+))\(([^)]*)\), '
               r'(-?' + LITERAL + r'(?:\s*[-+]\s*' + LITERAL + r')?), (\w+)')
    checks = re.findall(pattern, joined(sub))
    if len(checks) != 18:
        problems.append(f"expected 18 relative assertions, found {len(checks)}")
    with mp.workdps(60):
        for label, _, kind, argtext, reftext, tol_name in checks:
            args = [float(eval(a.replace("#", ""))) for a in argtext.split(",")]
            reference = vba_value(reftext, problems, label)
            tol = TOLERANCES[tol_name]
            a, df = args
            if kind == "InverseCumulative":
                target = mp.mpf(1) - mp.mpf(a) if a > 0.5 else mp.mpf(a)
                seed = mp.sqrt(2) * mp.erfinv(2 * abs(mp.mpf(a) - mp.mpf("0.5")))
                q = mp.findroot(lambda x: V.oracle_tail(x, mp.mpf(df)) - target, seed)
                exact = q if a > 0.5 else -q
                if PP.reaches_route(abs(float(exact)), df) and PP.quantile_bound(a, float(exact), df) > tol:
                    problems.append(f"{label}: PT2 bound exceeds the test tolerance")
            else:
                tail = V.oracle_tail(mp.mpf(abs(a)), mp.mpf(df))
                upper = (kind == "Survival") == (a >= 0)
                exact = tail if upper else 1 - tail
                if PP.reaches_route(abs(a), df):
                    model = P.survival(a, df) if kind == "Survival" else P.cumulative(a, df)
                    if abs(mp.mpf(model) - exact) / abs(exact) > tol:
                        problems.append(f"{label}: the model misses the reference by more than {tol}")
                elif tol_name == "TOL_REL_LARGE_DF_TAIL":
                    problems.append(f"{label}: outside the region but held to the route's tolerance")
            if reference != float(exact):
                problems.append(f"{label}: reference {reference!r} is not the correctly rounded "
                                f"{float(exact)!r}")
    return problems


POLISH = [
    "If DegreesFreedom >= PROB_T_LDF_MIN_DF And Tail <= PROB_T_LDF_POLISH_MAX_TAIL Then",
    "For PolishIdx = 1 To PROB_T_LDF_POLISH_STEPS",
    "If X > PROB_T_LDF_MAX_ABS_T Then Exit For",
    "If Not PROB_TryStudentTTail(X, DegreesFreedom, TailAtX, FailMsg) Then Exit Function",
    "If Not PROB_TryStudentTPDF(X, DegreesFreedom, Density, FailMsg) Then Exit Function",
    "If Density <= 0# Then Exit For",
    "XNew = X + (TailAtX - Tail) / Density",
    "If XNew = X Then Exit For",
    "X = XNew",
    "Next PolishIdx",
    "End If",
]


def polish_problems(text):
    """Check 6 on a module text; [] when the polish is the preregistered one."""
    problems = []
    flat = joined(text)
    consts = {name: (kind, value.strip()) for name, kind, value in re.findall(
        r"^Private Const (PROB_T_LDF_POLISH_\w+)\s+As (Double|Long) = ([^'\n]+)", flat, re.M)}
    if consts.get("PROB_T_LDF_POLISH_STEPS") != ("Long", str(PI.STEPS)):
        problems.append(f"PROB_T_LDF_POLISH_STEPS is {consts.get('PROB_T_LDF_POLISH_STEPS')}")
    tail = consts.get("PROB_T_LDF_POLISH_MAX_TAIL")
    if not tail or tail[0] != "Double" or float(tail[1].rstrip("#")) != PI.MAX_TAIL:
        problems.append(f"PROB_T_LDF_POLISH_MAX_TAIL is {tail}")
    inverse = procedure(text, "PROB_TryStudentTInvTail")
    if inverse is None:
        return problems + ["PROB_TryStudentTInvTail not found"]
    lines = statements(inverse)
    starts = [i for i in range(len(lines)) if lines[i:i + len(POLISH)] == POLISH]
    if len(starts) != 1:
        return problems + ["the polish statements are not exactly the preregistered ones"]
    start = starts[0]
    try:
        converged = lines.index("If Not Converged Then")
        returned = lines.index("Result = X")
    except ValueError:
        return problems + ["the convergence check or Result = X is missing"]
    if not (converged < start and start + len(POLISH) == returned):
        problems.append("the polish is not between the convergence check and Result = X")
    return problems


def quantile_problems(text):
    """Check 7 on Test_TF_StudentTLargeDFQuantile."""
    problems = []
    sub = procedure(text, "Test_TF_StudentTLargeDFQuantile")
    if sub is None:
        return ["Test_TF_StudentTLargeDFQuantile not found"]
    pattern = (r'AssertRelClose "([^"]+)", K_STATS_StudentT_InverseCumulative\(([^)]*)\), '
               r'(-?' + LITERAL + r'(?:\s*[-+]\s*' + LITERAL + r')?), (\w+)')
    checks = re.findall(pattern, joined(sub))
    if len(checks) != 5:
        problems.append(f"expected 5 quantile assertions, found {len(checks)}")
    with open(os.path.join(HERE, "excel_wave_check.json"), encoding="utf-8") as f:
        first = json.load(f)["PT2"]
    for label, argtext, reftext, tol_name in checks:
        p, df = (float(eval(a.replace("#", ""))) for a in argtext.split(","))
        tol = TOLERANCES[tol_name]
        reference = vba_value(reftext, problems, label)
        qd, found, _, _ = PI.outcomes(p, df)
        if reference != qd:
            problems.append(f"{label}: reference {reference!r} is not the correctly rounded {qd!r}")
            continue
        if any(abs(x - qd) / qd > tol for x in found):
            problems.append(f"{label}: an allowed outcome is outside the tolerance")
        earlier = [r["error"] for r in first if float(r["args"][0]) == p and float(r["args"][1]) == df]
        if not earlier or min(earlier) <= tol:
            problems.append(f"{label}: the tolerance does not discriminate the first wave's error")
    return problems


def mutate(text, old, new):
    if text.count(old) < 1:
        raise AssertionError(f"mutation anchor not found: {old!r}")
    return text.replace(old, new, 1)


def main():
    fails = []
    module = read(SOURCE)
    fails += source_problems(module)
    fails += reference_problems(read(TESTS))
    fails += polish_problems(module)
    fails += quantile_problems(read(TESTS))
    mutations = {
        "a coefficient's last digit": mutate(module, "PROB_T_LDF_G2_1 As Double = -5.20833333333333E-02",
                                             "PROB_T_LDF_G2_1 As Double = -5.20833333333334E-02"),
        "a reassociated product": mutate(
            module,
            "Correction = PROB_T_INV_SQRT_TWO_PI * PROB_ExpNegHalfSquare(AbsX) * Correction",
            "Correction = PROB_T_INV_SQRT_TWO_PI * (PROB_ExpNegHalfSquare(AbsX) * Correction)"),
        "a dropped order": mutate(module, "        Series = Series * W + Horner\n", ""),
        "the predicate bound": mutate(module,
                                      "PROB_T_LDF_MAX_ABS_T As Double = 8#",
                                      "PROB_T_LDF_MAX_ABS_T As Double = 6#"),
    }
    for what, text in mutations.items():
        if not source_problems(text):
            fails.append(f"mutation not caught: {what}")
    block = "\n".join(module.split("\n"))
    polish_start = block.index("        If DegreesFreedom >= PROB_T_LDF_MIN_DF And Tail <= PROB_T_LDF_POLISH_MAX_TAIL Then")
    polish_end = block.index("            Next PolishIdx\n        End If\n", polish_start) + len(
        "            Next PolishIdx\n        End If\n")
    polish_text = block[polish_start:polish_end]
    check_at = block.index("    'Reject a solver that exhausted the iteration budget")
    moved = block[:polish_start] + block[polish_end:]
    moved = moved[:check_at] + polish_text + moved[check_at:]
    polish_mutations = {
        "a different step count": mutate(module, "PROB_T_LDF_POLISH_STEPS As Long = 3",
                                         "PROB_T_LDF_POLISH_STEPS As Long = 4"),
        "a different tail bound": mutate(module, "PROB_T_LDF_POLISH_MAX_TAIL As Double = 0.4",
                                         "PROB_T_LDF_POLISH_MAX_TAIL As Double = 0.5"),
        "a missing XNew = X exit": mutate(module, "                If XNew = X Then Exit For\n", ""),
        "the polish moved before the convergence check": moved,
    }
    for what, text in polish_mutations.items():
        if not polish_problems(text):
            fails.append(f"mutation not caught: {what}")
    if fails:
        print("FAIL: #34 VBA large-df route source fidelity")
        for f in fails:
            print("  - " + f)
        return 1
    print(f"PASS: #34 VBA large-df route: {sum(len(r) for r in P.COEFFICIENTS)} coefficients and "
          f"the predicate match the preregistration exactly; {len(expected_body())} statements in "
          f"order; dispatch wired; 18 test references correctly rounded and reproduced; "
          f"{len(mutations)} mutations caught; inverse polish as preregistered, 5 quantile "
          f"references correctly rounded and discriminating, {len(polish_mutations)} polish mutations caught")
    return 0


if __name__ == "__main__":
    sys.exit(main())
