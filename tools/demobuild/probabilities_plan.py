"""The Probabilities sheet as data: every formula, name, validation rule and text.

Same method as explorer_plan.py: examples/demo/explorer-families.json is the one
source of the families, and this module writes from it the plan that
tools/demobuild/Evolve-Probabilities.ps1 applies in desktop Excel. The sheet
reads the family table that the Explorer sheet keeps in its hidden columns (the
Explorer* names: the labels, limits and messages of the parameters), so the 16
distributions, their limits and their input messages are defined once.

Layout, in the style of the Explorer: in C:D the inputs D6:D14 (distribution, up
to four parameters, the points a and b, the probability p, the number of
decimals), the status D17, the macros cell D18, then the results D20:D28:

    D20 P(X <= a)         D24 P(a < X <= b)       D27 round-trip difference
    D21 P(X > a)          D25 quantile Q(p)       D28 round-trip verdict
    D22 P(X <= b)         D26 F(Q(p))
    D23 P(X > b)

The notes are in F5:I17. The grid starts at F20 under its headings in row 19:
x (or the integer k), the density or mass, and two columns that keep the density
where the first region (x <= a) and the second (a < x <= b) are and hold 0
elsewhere, for the shaded chart areas; one dynamic-array formula per column.

The interval is computed in the tail that avoids cancellation: S(a) - S(b) when
F(a) is above one half, F(b) - F(a) otherwise. The round trip compares F(Q(p))
with p relative to min(p, 1 - p), within ROUNDTRIP_TOLERANCE. A discrete family
has no exact round trip; the verdict states the definition of its quantile,
F(Q - 1) < p <= F(Q).
"""
from __future__ import annotations

import re
from typing import Any

import explorer_plan as E

SHEET = "Probabilities"
GRID_FIRST_ROW = 20
PARAMETER_CELLS = E.PARAMETER_CELLS            # D7:D10, as on the Explorer
A, B, PROBABILITY, DECIMALS = "$D$11", "$D$12", "$D$13", "$D$14"
STATUS = "$D$17"
MACROS = "D18"
F_A, S_A, F_B, S_B, INTERVAL, QUANTILE, F_Q, DIFFERENCE, VERDICT = (
    "D20", "D21", "D22", "D23", "D24", "D25", "D26", "D27", "D28")
SUPPORT = "C30"
MAX_DECIMALS = E.MAX_DECIMALS
# The quantile contracts of the library are relative in x (5E-15 to 4.7E-12).
# F(Q(p)) differs from p by about f(Q)*Q*delta relative to F(Q), which for the
# Normal at p = 1E-15 (|Q| = 7.9) is 63 times the contract: 1E-8 leaves room
# for the tails the sheet accepts and still shows a wrong quantile.
ROUNDTRIP_TOLERANCE = 1e-8
LIMIT_POINTS = 1000     # the grid of a discrete family, as on the Explorer

NAMES_ROW = "ProbabilitiesRow"


def ref(address: str) -> str:
    return f"{SHEET}!{address}"


def dispatch(families: list[dict[str, Any]], operation: str, point: str) -> str:
    """The selected family's call, as one SWITCH over the distribution cell."""
    return E.dispatch(families, operation, point)


def plot_expression(families: list[dict[str, Any]], key: str) -> str:
    """The plot bound the family suggests, as a SWITCH over the distribution."""
    branches = ",".join(f'"{f["name"]}",{E.sheet_expression(f[key])}' for f in families)
    return f"SWITCH($D$6,{branches},NA())"


def constraint_expression(families: list[dict[str, Any]]) -> str:
    """TRUE unless the selected family's parameters contradict each other."""
    branches = ",".join(f'"{f["name"]}",{E.sheet_expression(f["constraint"])}'
                        for f in families if f["constraint"] != "TRUE")
    return f"SWITCH($D$6,{branches},TRUE)"


def constraint_text(families: list[dict[str, Any]]) -> str:
    branches = ",".join(f'"{f["name"]}","{f["constraint_text"]}"' for f in families if f["constraint"] != "TRUE")
    return f'SWITCH($D$6,{branches},"")'


def check_formula() -> str:
    """The first input that is not acceptable, else "": parameters, then constraint, then a and b."""
    parts = []
    for slot in range(1, 5):
        value = PARAMETER_CELLS[slot - 1]
        parts.append(f'IF(OR({slot}>INDEX(ExplorerNP,ProbabilitiesRow),'
                     f'IFERROR(AND(ISNUMBER({value}),{value}>=INDEX(ExplorerMin,ProbabilitiesRow,{slot}),'
                     f'{value}<=INDEX(ExplorerMax,ProbabilitiesRow,{slot}),'
                     f'OR(INDEX(ExplorerInt,ProbabilitiesRow,{slot})=0,{value}=INT({value}))),FALSE)),"",'
                     f'INDEX(ExplorerLabels,ProbabilitiesRow,{slot}))')
    return "TEXTJOIN(\", \",TRUE," + ",".join(parts) + ")"


def status_formula(defaults: dict[str, Any]) -> str:
    limit = E.number(defaults["x_limit"])
    return (
        f'=IFERROR(LET(parameter,{check_formula()},'
        f'IF(parameter<>"","Check: "&parameter,'
        f'IF(NOT(IFERROR({constraint_expression_placeholder()},FALSE)),"Check: "&{constraint_text_placeholder()},'
        f'IF(NOT(AND(ISNUMBER({A}),ISNUMBER({B}),ABS({A})<={limit},ABS({B})<={limit})),"Check: points a and b",'
        f'IF({A}>{B},"Check: a must not exceed b",'
        f'IF(ProbabilitiesWindowTooWide,"Check: a and b are too far apart for a discrete grid","OK")))))),'
        f'"Check: distribution")'
    )


def constraint_expression_placeholder() -> str:
    return "ProbabilitiesConstraint"


def constraint_text_placeholder() -> str:
    return "ProbabilitiesConstraintText"


def effective(point: str) -> str:
    """The point as the library reads it: a discrete family truncates it to a whole number."""
    return f'IF(INDEX(ExplorerKind,ProbabilitiesRow)="Discrete",INT({point}),{point})'


def x_grid_formula() -> str:
    return (
        f'=IF({STATUS}="OK",IF(INDEX(ExplorerKind,ProbabilitiesRow)="Discrete",'
        f'LET(first,-INT(-ProbabilitiesLowerEnd),SEQUENCE(INT(ProbabilitiesUpperEnd)-first+1,,first)),'
        f'LET(i,SEQUENCE(201,,0),ProbabilitiesLowerEnd+(ProbabilitiesUpperEnd-ProbabilitiesLowerEnd)*i/200)),NA())'
    )


def density_formula(families: list[dict[str, Any]]) -> str:
    return (f'=IF({STATUS}="OK",MAP(ProbabilitiesX,LAMBDA(x,IFERROR({dispatch(families, "density", "x")},NA()))),NA())')


def shade_first_formula() -> str:
    """The density where x <= a, and 0 elsewhere: an area chart draws nothing at 0."""
    return (f'=IF({STATUS}="OK",IF(ProbabilitiesX<={effective(A)},ProbabilitiesDensity,0),NA())')


def shade_second_formula() -> str:
    return (f'=IF({STATUS}="OK",IF((ProbabilitiesX>{effective(A)})*(ProbabilitiesX<={effective(B)}),'
            f'ProbabilitiesDensity,0),NA())')


def results_formulas(families: list[dict[str, Any]]) -> list[dict[str, str]]:
    ok = f'{STATUS}="OK"'

    def cdf(point: str) -> str:
        return dispatch(families, "cdf", point)

    def survival(point: str) -> str:
        return dispatch(families, "survival", point)

    quantile = dispatch(families, "quantile", PROBABILITY)
    good_p = f"AND(ISNUMBER({PROBABILITY}),{PROBABILITY}>=0,{PROBABILITY}<=1)"
    discrete = 'INDEX(ExplorerKind,ProbabilitiesRow)="Discrete"'
    tolerance = E.number(ROUNDTRIP_TOLERANCE)
    return [
        {"address": F_A, "formula": f"=IF({ok},{cdf(A)},NA())"},
        {"address": S_A, "formula": f"=IF({ok},{survival(A)},NA())"},
        {"address": F_B, "formula": f"=IF({ok},{cdf(B)},NA())"},
        {"address": S_B, "formula": f"=IF({ok},{survival(B)},NA())"},
        # The tail that avoids the cancellation of the two cumulative values.
        {"address": INTERVAL,
         "formula": f"=IFERROR(MAX(0,IF({F_A}>0.5,{S_A}-{S_B},{F_B}-{F_A})),NA())"},
        {"address": QUANTILE, "formula": f"=IF(AND({ok},{good_p}),{quantile},NA())"},
        {"address": F_Q, "formula": f"=IFERROR({cdf(QUANTILE)},NA())"},
        {"address": DIFFERENCE,
         "formula": (f'=IFERROR(IF({discrete},"n/a",IF(OR({PROBABILITY}<=0,{PROBABILITY}>=1),"n/a",'
                     f'IF({PROBABILITY}<=0.5,ABS({F_Q}-{PROBABILITY})/{PROBABILITY},'
                     f'ABS({survival(QUANTILE)}-(1-{PROBABILITY}))/(1-{PROBABILITY})))),NA())')},
        {"address": VERDICT,
         "formula": (f'=IFERROR(IF(OR({PROBABILITY}<=0,{PROBABILITY}>=1),"End of the support",'
                     f'IF({discrete},IF(AND({F_Q}>={PROBABILITY},{cdf(QUANTILE + "-1")}<{PROBABILITY}),'
                     f'"Q is the smallest k with F(k) >= p","CHECK: Q is not the smallest k"),'
                     f'IF({DIFFERENCE}<={tolerance},"Within the tolerance of {tolerance.upper()}",'
                     f'"OUTSIDE the tolerance of {tolerance.upper()}"))),NA())')},
    ]


def static_messages(defaults: dict[str, Any]) -> dict[str, list[str]]:
    """Input title, input message, error title and error message of the fixed input cells."""
    limit = E.shown(defaults["x_limit"])
    return {
        "D6": ["Distribution", "Choose one of the 16 distributions. Its parameters and the points a and b are "
               "loaded with it.", "Unknown distribution", "Choose a distribution from the list."],
        "D11": ["Lower point a", "Lower end of the interval, in units of x. A discrete family truncates it to a "
                f"whole number. From -{limit} to {limit}, not above b.",
                "Invalid lower point", f"Enter a number from -{limit} to {limit}, not above b."],
        "D12": ["Upper point b", "Upper end of the interval, in units of x. A discrete family truncates it to a "
                f"whole number. From -{limit} to {limit}, not below a.",
                "Invalid upper point", f"Enter a number from -{limit} to {limit}, not below a."],
        "D13": ["Probability p", "Probability for the quantile: the result is the smallest x with F(x) >= p. "
                "From 0 to 1; 0 and 1 give the ends of the support.",
                "Invalid probability", "Enter a probability from 0 to 1."],
        "D14": ["Number of decimals", "Decimals shown in the results and the grid, from 0 to 15. Display only: "
                "the values keep their full precision.",
                "Invalid number of decimals", "Choose a number of decimals from 0 to 15."],
    }


def plan() -> dict[str, Any]:
    data = E.load()
    families, defaults = data["families"], data["defaults"]
    last_row = GRID_FIRST_ROW + max(int(defaults["max_points"]), int(defaults["max_discrete_points"])) - 1
    names = [
        {"name": NAMES_ROW, "refers": f"=MATCH({ref('$D$6')},ExplorerFamily,0)"},
        {"name": "ProbabilitiesPlotLower", "refers": "=" + plot_expression(families, "plot_lower").replace("$D$", f"{SHEET}!$D$")},
        {"name": "ProbabilitiesPlotUpper", "refers": "=" + plot_expression(families, "plot_upper").replace("$D$", f"{SHEET}!$D$")},
        {"name": "ProbabilitiesLowerEnd", "refers": f"=MIN(ProbabilitiesPlotLower,{ref(A)})"},
        {"name": "ProbabilitiesUpperEnd", "refers": f"=MAX(ProbabilitiesPlotUpper,{ref(B)})"},
        {"name": "ProbabilitiesWindowTooWide",
         "refers": (f'=AND(INDEX(ExplorerKind,ProbabilitiesRow)="Discrete",'
                    f'INT(ProbabilitiesUpperEnd)--INT(-ProbabilitiesLowerEnd)>{LIMIT_POINTS})')},
        {"name": "ProbabilitiesConstraint", "refers": "=" + constraint_expression(families).replace("$D$", f"{SHEET}!$D$")},
        {"name": "ProbabilitiesConstraintText", "refers": "=" + constraint_text(families).replace("$D$", f"{SHEET}!$D$")},
        {"name": "ProbabilitiesX", "refers": f"={ref(f'$F${GRID_FIRST_ROW}')}#"},
        {"name": "ProbabilitiesDensity", "refers": f"={ref(f'$G${GRID_FIRST_ROW}')}#"},
        {"name": "ProbabilitiesFirstRegion", "refers": f"={ref(f'$H${GRID_FIRST_ROW}')}#"},
        {"name": "ProbabilitiesSecondRegion", "refers": f"={ref(f'$I${GRID_FIRST_ROW}')}#"},
        {"name": "ProbabilitiesDecimals", "refers": "=ExplorerDecimals"},
        {"name": "ProbabilitiesStatusBad", "refers": f'={ref(STATUS)}<>"OK"'},
    ]
    for slot in range(1, 5):
        value = ref(f"$D${6 + slot}")
        names += [
            {"name": f"ProbabilitiesParameterOk{slot}",
             "refers": (f"=IFERROR(AND(ISNUMBER({value}),{value}>=INDEX(ExplorerMin,ProbabilitiesRow,{slot}),"
                        f"{value}<=INDEX(ExplorerMax,ProbabilitiesRow,{slot}),"
                        f"OR(INDEX(ExplorerInt,ProbabilitiesRow,{slot})=0,{value}=INT({value}))),FALSE)")},
            {"name": f"ProbabilitiesParameterUnused{slot}", "refers": f"={slot}>INDEX(ExplorerNP,ProbabilitiesRow)"},
        ]
    formulas = [
        {"address": STATUS.replace("$", ""), "formula": status_formula(defaults)},
        {"address": MACROS, "formula": '=IF(ISERROR(K_STATS_NormalStandard_Cumulative(0)+0*COUNTA($D$6:$D$13)),"DISABLED","ENABLED")'},
        *results_formulas(families),
        {"address": SUPPORT, "formula": '=IFERROR(INDEX(ExplorerSupport,ProbabilitiesRow),"")'},
        {"address": "F4", "formula": f'=IF(${MACROS[0]}${MACROS[1:]}="ENABLED","","MACROS ARE DISABLED: the results are not live")'},
        {"address": f"F{GRID_FIRST_ROW}", "formula": x_grid_formula()},
        {"address": f"G{GRID_FIRST_ROW}", "formula": density_formula(families)},
        {"address": f"H{GRID_FIRST_ROW}", "formula": shade_first_formula()},
        {"address": f"I{GRID_FIRST_ROW}", "formula": shade_second_formula()},
    ]
    for slot in range(4):
        formulas.append({"address": f"C{7 + slot}",
                         "formula": f'=IFERROR(INDEX(ExplorerLabels,ProbabilitiesRow,{slot + 1}),"")'})
    texts = [
        {"address": "C3", "value": "Choose a distribution, the points a and b and the probability p. Every result is a call "
                                   "to the embedded VBA API (K_STATS_ functions). Macros are required."},
        {"address": "C11", "value": "Lower point a"},
        {"address": "C12", "value": "Upper point b"},
        {"address": "C13", "value": "Probability p"},
        {"address": "C14", "value": "Number of decimals"},
        {"address": "C17", "value": "Input status"},
        {"address": "C18", "value": "Macros"},
        {"address": "C20", "value": "P(X <= a)"},
        {"address": "C21", "value": "P(X > a)"},
        {"address": "C22", "value": "P(X <= b)"},
        {"address": "C23", "value": "P(X > b)"},
        {"address": "C24", "value": "P(a < X <= b)"},
        {"address": "C25", "value": "Quantile Q(p)"},
        {"address": "C26", "value": "F(Q(p))"},
        {"address": "C27", "value": "Round-trip difference"},
        {"address": "C28", "value": "Round-trip verdict"},
        {"address": "F5", "value": "Read the results"},
        {"address": "F6", "value": "Yellow cells are editable. A new distribution loads its parameters and the points a and b."},
        {"address": "F8", "value": "Continuous: a point has probability zero, so < and <= give the same value. "
                                   "Discrete: a and b are truncated to whole numbers; P(X <= a) includes a."},
        {"address": "F10", "value": "P(a < X <= b) is S(a) - S(b) when F(a) > 0.5 and F(b) - F(a) otherwise: the tail that "
                                    "keeps its digits."},
        {"address": "F12", "value": "Round trip: F(Q(p)) against p, relative to min(p, 1 - p), within "
                                    f"{E.number(ROUNDTRIP_TOLERANCE).upper()}. A discrete quantile is the smallest k with F(k) >= p."},
        {"address": "F14", "value": "Educational demo, not an accuracy certification. The library's guarantees are in docs/."},
        {"address": "F16", "value": "Needs macros and Excel 365 or 2024. If Macros says DISABLED, unblock the file and reopen."},
        {"address": "B1", "value": "Probabilities and Quantiles"},
        {"address": "B2", "value": "Lower tail, upper tail and interval probabilities of a distribution, and the quantile that inverts them."},
        {"address": f"F{GRID_FIRST_ROW - 1}", "value": "x / integer k"},
        {"address": f"G{GRID_FIRST_ROW - 1}", "value": "PDF / PMF"},
        {"address": f"H{GRID_FIRST_ROW - 1}", "value": "x <= a"},
        {"address": f"I{GRID_FIRST_ROW - 1}", "value": "a < x <= b"},
    ]
    return {
        "schema": "probabilities-plan-1",
        "sheet": SHEET,
        "names": names,
        "formulas": formulas,
        "texts": texts,
        "grid": {"first_row": GRID_FIRST_ROW, "last_row": last_row, "columns": ["F", "G", "H", "I"]},
        "decimals": {"cell": DECIMALS, "default": 6, "max": MAX_DECIMALS,
                     "cells": f"{F_A}:{F_Q},{DIFFERENCE},G{GRID_FIRST_ROW}:I{last_row}",
                     "x_cells": f"F{GRID_FIRST_ROW}:F{last_row}"},
        "probability": {"cell": PROBABILITY, "default": 0.95},
        "roundtrip_tolerance": ROUNDTRIP_TOLERANCE,
        "chart": {"name": "ProbabilitiesDensity",
                  "series": [{"name": "Density / mass", "x": "ProbabilitiesX", "y": "ProbabilitiesDensity"},
                             {"name": "P(X <= a)", "x": "ProbabilitiesX", "y": "ProbabilitiesFirstRegion"},
                             {"name": "P(a < X <= b)", "x": "ProbabilitiesX", "y": "ProbabilitiesSecondRegion"}]},
        "messages": static_messages(defaults),
        "layout": {"family": "D6", "lower": A.replace("$", ""), "upper": B.replace("$", ""),
                   "probability": PROBABILITY.replace("$", ""), "decimals": DECIMALS.replace("$", ""),
                   "status": STATUS.replace("$", ""), "macros": MACROS, "inputs": "D6:D14",
                   "outputs": [f"{STATUS.replace('$', '')}:{VERDICT}"]},
        "validation": {
            "family_list": "=ExplorerFamily",
            "parameter": "=ProbabilitiesParameterOk{slot}",
            "point_limit": defaults["x_limit"],
            "decimals_list": "=ExplorerDecimals",
            "probability": [0, 1],
        },
        "conditional": {
            "unused_parameter": "=ProbabilitiesParameterUnused{slot}",
            "status_bad": "=ProbabilitiesStatusBad",
            "grid_inside": f"=ROW()-{GRID_FIRST_ROW - 1}<=ROWS(ProbabilitiesX)",
        },
    }


def formula_lengths(result: dict[str, Any]) -> list[tuple[str, int]]:
    """Length of every formula and name, to check them against Excel's limit of 8192 characters."""
    sizes = [(item["address"], len(item["formula"])) for item in result["formulas"]]
    sizes += [(item["name"], len(item["refers"])) for item in result["names"]]
    return sorted(sizes, key=lambda item: -item[1])


if __name__ == "__main__":
    for name, size in formula_lengths(plan())[:6]:
        print(size, name)
    assert re.search("ProbabilitiesRow", status_formula(E.load()["defaults"]))
