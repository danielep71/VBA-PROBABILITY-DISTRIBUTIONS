"""The Parameter Effects sheet as data: every formula, name, validation rule and text.

Same method as explorer_plan.py and probabilities_plan.py: examples/demo/explorer-families.json
is the source of the 16 families, and examples/demo/parameter-effects.json adds what this sheet
needs of each one: four values of each parameter, the parameter that is varied first, a line on
what the parameter does, and the first three moments. The family table that the Explorer keeps
in its hidden columns supplies the labels, limits and input messages (the Explorer* names).

Layout, in the style of the Explorer: in C:D the inputs D6:D16 (distribution, up to four base
parameters, the parameter to vary, the four values it takes, the number of decimals), the status
D19, the macros cell D20; the four curves in rows 22 to 25 (C and D name each curve and its
value, F:I give its mean, standard deviation, skewness and the probability beyond three standard
deviations, under their headings in row 21); the line on what the parameter does in C27:D32.
The notes are in F5:I17. The curves are drawn from hidden columns: x in R, the four densities in
S:V, the four cumulative functions in W:Z, the table of the families from AB, the window and the
legend texts in AX:AZ. The two charts, in K:O, follow them through workbook names.

A curve is the family at the base parameters with the varied parameter set to its value: every
formula of curve k opens with LET(pa, IF(slot = 1, value, $D$7), ...), so the dispatch of the
Explorer (one SWITCH over the distribution) is reused with pa..pd in place of the input cells.
All curves share one x axis: the union of the plot windows of the four. The moments are written
in the parameters and need no library function; a moment that does not exist reads "undefined"
or "infinite". The tail column is P(X > mean + 3 sd), from the library's survival function.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import explorer_plan as E
import probabilities_plan as P

ROOT = Path(__file__).resolve().parents[2]
EFFECTS_FILE = ROOT / "examples" / "demo" / "parameter-effects.json"

SHEET = "Parameters"
CURVES = 4
VARY = "$D$11"
CURVE_VALUES = ["$D$12", "$D$13", "$D$14", "$D$15"]
DECIMALS = "$D$16"
STATUS = "$D$19"
MACROS = "D20"
HEADER_ROW = 21
FIRST_CURVE_ROW = 22                 # rows 22 to 25
EFFECT_CELL = "C27"                  # merged C27:D32
GRID_FIRST_ROW = 20
X_COLUMN = "R"
DENSITY_COLUMNS = ["S", "T", "U", "V"]
CDF_COLUMNS = ["W", "X", "Y", "Z"]
CURVE_TABLE_COLUMN = 28              # AB: 16 values, slot by slot, four per slot
VARY_COLUMN = 44                     # AR
EFFECT_COLUMN = 45                   # AS: four texts
WINDOW_COLUMNS = ("AX", "AY")        # lower and upper bound of each curve's plot window
LEGEND_COLUMN = "AZ"
POINTS = 201
MAX_DISCRETE_SPAN = 1000
TAIL_SIGMAS = 3
WORDS = ("pa", "pb", "pc", "pd")
CELLS = ("$D$7", "$D$8", "$D$9", "$D$10")
TEXT_MOMENTS = ("undefined", "infinite")


def load() -> dict[str, Any]:
    data: dict[str, Any] = json.loads(EFFECTS_FILE.read_text(encoding="utf-8"))
    families = E.load()["families"]
    missing = [family["name"] for family in families if family["name"] not in data["families"]]
    if missing:
        raise ValueError(f"no parameter effects for {missing}")
    for family in families:
        entry = data["families"][family["name"]]
        if len(entry["parameters"]) != len(family["parameters"]):
            raise ValueError(f'{family["name"]}: one effect per parameter')
        if not 1 <= entry["vary"] <= len(family["parameters"]):
            raise ValueError(f'{family["name"]}: the parameter to vary')
        for effect, parameter in zip(entry["parameters"], family["parameters"]):
            if len(effect["curves"]) != CURVES:
                raise ValueError(f'{family["name"]}: {CURVES} values per parameter')
            for value in effect["curves"]:
                if not parameter["min"] <= value <= parameter["max"]:
                    raise ValueError(f'{family["name"]}: {value} is outside the limits of {parameter["label"]}')
                if parameter["integer"] and float(value) != int(value):
                    raise ValueError(f'{family["name"]}: {value} must be a whole number')
    return data


def ref(address: str) -> str:
    return f"{SHEET}!{address}"


def in_words(expression: str) -> str:
    """An expression written in the parameters a, b, c, d, in the names pa, pb, pc, pd."""
    return re.sub(r"\b([abcd])\b", lambda found: WORDS[E.LETTERS.index(found.group(1))], expression)


def cells_to_words(expression: str) -> str:
    """The Explorer's dispatch, written on the input cells $D$7..$D$10, on pa..pd."""
    for cell, word in reversed(list(zip(CELLS, WORDS))):   # $D$10 before $D$1x
        expression = expression.replace(cell, word)
    return expression


def dispatch(families: list[dict[str, Any]], operation: str, point: str) -> str:
    return cells_to_words(E.dispatch(families, operation, point))


def curve(k: int, body: str) -> str:
    """`body` evaluated at the parameters of curve k (1 to 4): the base ones, the varied one replaced."""
    value = CURVE_VALUES[k - 1]
    names = ",".join(f"{word},IF(ParameterEffectsSlot={slot + 1},{value},{CELLS[slot]})"
                     for slot, word in enumerate(WORDS))
    return f"LET({names},{body})"


def switch(families: list[dict[str, Any]], data: dict[str, Any], moment: str) -> str:
    branches = ",".join(f'"{family["name"]}",{in_words(data["families"][family["name"]]["moments"][moment])}'
                        for family in families)
    return f"SWITCH($D$6,{branches},NA())"


def plot_switch(families: list[dict[str, Any]], key: str) -> str:
    branches = ",".join(f'"{family["name"]}",{in_words(E.sheet_expression(family[key]))}' for family in families)
    return cells_to_words(f"SWITCH($D$6,{branches},NA())")


def constraint(families: list[dict[str, Any]]) -> str:
    return cells_to_words(P.constraint_expression(families))


def x_grid_formula() -> str:
    return (f'=IF({STATUS}="OK",IF(INDEX(ExplorerKind,ParameterEffectsRow)="Discrete",'
            f'LET(first,-INT(-ParameterEffectsLower),SEQUENCE(INT(ParameterEffectsUpper)-first+1,,first)),'
            f'LET(i,SEQUENCE({POINTS},,0),ParameterEffectsLower+(ParameterEffectsUpper-ParameterEffectsLower)*i/{POINTS - 1})),NA())')


def grid_formula(families: list[dict[str, Any]], k: int, operation: str) -> str:
    call = dispatch(families, operation, "x")
    return f'=IF({STATUS}="OK",MAP(ParameterEffectsX,LAMBDA(x,{curve(k, f"IFERROR({call},NA())")})),NA())'


def window_formula(families: list[dict[str, Any]], k: int, key: str) -> str:
    return f'=IFERROR({curve(k, plot_switch(families, key))},"")'


def curve_ok(k: int, families: list[dict[str, Any]]) -> str:
    """TRUE when the value of curve k is a number within the limits of the varied parameter, and
    the parameters of the curve do not contradict each other."""
    value = CURVE_VALUES[k - 1]
    inside = (f"AND(ISNUMBER({value}),{value}>=INDEX(ExplorerMin,ParameterEffectsRow,ParameterEffectsSlot),"
              f"{value}<=INDEX(ExplorerMax,ParameterEffectsRow,ParameterEffectsSlot),"
              f"OR(INDEX(ExplorerInt,ParameterEffectsRow,ParameterEffectsSlot)=0,{value}=INT({value})))")
    return f"IFERROR(AND({inside},{curve(k, constraint(families))}),FALSE)"


def check_base() -> str:
    """The first base parameter that is not acceptable, leaving out the one that is varied."""
    parts = []
    for slot in range(1, 5):
        parts.append(f'IF(OR({slot}>INDEX(ExplorerNP,ParameterEffectsRow),{slot}=ParameterEffectsSlot,'
                     f'ParameterEffectsParameterOk{slot}),"",INDEX(ExplorerLabels,ParameterEffectsRow,{slot}))')
    return 'TEXTJOIN(", ",TRUE,' + ",".join(parts) + ")"


def status_formula(families: list[dict[str, Any]]) -> str:
    bad_curves = ",".join(f'IF({curve_ok(k, families)},"","curve {k}")' for k in range(1, CURVES + 1))
    return (f'=IFERROR(LET(base,{check_base()},IF(base<>"","Check: "&base,'
            f'IF(NOT(ISNUMBER(ParameterEffectsSlot)),"Check: parameter to vary",'
            f'LET(curves,TEXTJOIN(", ",TRUE,{bad_curves}),'
            f'IF(curves<>"","Check: "&curves,'
            f'IF(ParameterEffectsWindowTooWide,"Check: curves are too far apart for a discrete grid","OK")))))),'
            f'"Check: distribution")')


def stat_formulas(families: list[dict[str, Any]], data: dict[str, Any], k: int) -> list[dict[str, str]]:
    row = FIRST_CURVE_ROW + k - 1
    ok = f'{STATUS}="OK"'
    mean = curve(k, switch(families, data, "mean"))
    deviation = curve(k, f'LET(v,{switch(families, data, "variance")},IF(ISNUMBER(v),SQRT(v),v))')
    skewness = curve(k, switch(families, data, "skewness"))
    tail = curve(k, dispatch(families, "survival", f"F{row}+{TAIL_SIGMAS}*G{row}"))
    return [
        {"address": f"F{row}", "formula": f'=IF({ok},IFERROR({mean},"undefined"),NA())'},
        {"address": f"G{row}", "formula": f'=IF({ok},IFERROR({deviation},"undefined"),NA())'},
        {"address": f"H{row}", "formula": f'=IF({ok},IFERROR({skewness},"undefined"),NA())'},
        {"address": f"I{row}",
         "formula": f'=IF({ok},IF(AND(ISNUMBER(F{row}),ISNUMBER(G{row})),IFERROR({tail},NA()),"n/a"),NA())'},
    ]


def table_cells(families: list[dict[str, Any]], data: dict[str, Any]) -> list[dict[str, Any]]:
    """Constants of the hidden table, one row per family in the order of the Explorer's table."""
    cells: list[dict[str, Any]] = []
    for index, family in enumerate(families):
        entry = data["families"][family["name"]]
        row = index + 1
        for slot in range(4):
            used = slot < len(family["parameters"])
            for k in range(CURVES):
                value = entry["parameters"][slot]["curves"][k] if used else 0
                cells.append({"address": f"{E.column_letters(CURVE_TABLE_COLUMN + slot * CURVES + k)}{row}", "value": value})
            cells.append({"address": f"{E.column_letters(EFFECT_COLUMN + slot)}{row}",
                          "value": entry["parameters"][slot]["effect"] if used else ""})
        cells.append({"address": f"{E.column_letters(VARY_COLUMN)}{row}", "value": entry["vary"]})
    return cells


def static_messages(defaults: dict[str, Any]) -> dict[str, list[str]]:
    limits = "within its limits (the input message of its base cell gives them), and the parameters of the curve must not contradict each other"
    messages = {
        "D6": ["Distribution", "Choose one of the 16 distributions. Its base parameters, the parameter to vary "
               "and the four curve values are loaded with it.", "Unknown distribution", "Choose a distribution from the list."],
        "D11": ["Parameter to vary", "The parameter that takes the four values below, one curve each. The others stay "
                "at their base values. Choosing another loads its four values.", "Unknown parameter",
                "Choose a parameter from the list."],
        "D16": ["Number of decimals", "Decimals shown in the table, from 0 to 15. Display only: the values keep "
                "their full precision.", "Invalid number of decimals", "Choose a number of decimals from 0 to 15."],
    }
    for k in range(1, CURVES + 1):
        messages[f"D{11 + k}"] = [f"Curve {k} value", f"Value of the parameter to vary for curve {k}: a number {limits}.",
                                  f"Invalid curve {k} value", "Enter a value within the limits of the parameter to vary."]
    return messages


def plan() -> dict[str, Any]:
    data = load()
    families = E.load()["families"]
    defaults = E.load()["defaults"]
    rows = len(families)
    last_row = GRID_FIRST_ROW + defaults["max_discrete_points"] - 1
    names = [
        {"name": "ParameterEffectsRow", "refers": f"=MATCH({ref('$D$6')},ExplorerFamily,0)"},
        {"name": "ParameterEffectsSlot",
         "refers": f"=MATCH({ref(VARY)},INDEX(ExplorerLabels,ParameterEffectsRow,0),0)"},
        {"name": "ParameterEffectsChoices",
         "refers": ("=INDEX(ExplorerLabels,ParameterEffectsRow,1):"
                    "INDEX(ExplorerLabels,ParameterEffectsRow,INDEX(ExplorerNP,ParameterEffectsRow))")},
        {"name": "ParameterEffectsX", "refers": f"={ref(f'${X_COLUMN}${GRID_FIRST_ROW}')}#"},
        {"name": "ParameterEffectsCurves",
         "refers": (f"={ref(f'${E.column_letters(CURVE_TABLE_COLUMN)}$1')}:"
                    f"{f'${E.column_letters(CURVE_TABLE_COLUMN + 15)}${rows}'}")},
        {"name": "ParameterEffectsVary",
         "refers": f"={ref(f'${E.column_letters(VARY_COLUMN)}$1')}:{f'${E.column_letters(VARY_COLUMN)}${rows}'}"},
        {"name": "ParameterEffectsEffects",
         "refers": f"={ref(f'${E.column_letters(EFFECT_COLUMN)}$1')}:{f'${E.column_letters(EFFECT_COLUMN + 3)}${rows}'}"},
    ]
    first, last = FIRST_CURVE_ROW, FIRST_CURVE_ROW + CURVES - 1
    names += [
        {"name": "ParameterEffectsLower",
         "refers": f"=MIN({ref(f'${WINDOW_COLUMNS[0]}${first}')}:{f'${WINDOW_COLUMNS[0]}${last}'})"},
        {"name": "ParameterEffectsUpper",
         "refers": f"=MAX({ref(f'${WINDOW_COLUMNS[1]}${first}')}:{f'${WINDOW_COLUMNS[1]}${last}'})"},
        {"name": "ParameterEffectsWindowTooWide",
         "refers": ('=AND(INDEX(ExplorerKind,ParameterEffectsRow)="Discrete",'
                    f"INT(ParameterEffectsUpper)--INT(-ParameterEffectsLower)>{MAX_DISCRETE_SPAN})")},
        {"name": "ParameterEffectsStatusBad", "refers": f'={ref(STATUS)}<>"OK"'},
    ]
    for k in range(CURVES):
        names += [
            {"name": f"ParameterEffectsDensity{k + 1}", "refers": f"={ref(f'${DENSITY_COLUMNS[k]}${GRID_FIRST_ROW}')}#"},
            {"name": f"ParameterEffectsCdf{k + 1}", "refers": f"={ref(f'${CDF_COLUMNS[k]}${GRID_FIRST_ROW}')}#"},
        ]
        value = ref(CURVE_VALUES[k])
        names.append({"name": f"ParameterEffectsCurveOk{k + 1}",
                      "refers": (f"=IFERROR(AND(ISNUMBER({value}),{value}>=INDEX(ExplorerMin,ParameterEffectsRow,ParameterEffectsSlot),"
                                 f"{value}<=INDEX(ExplorerMax,ParameterEffectsRow,ParameterEffectsSlot),"
                                 f"OR(INDEX(ExplorerInt,ParameterEffectsRow,ParameterEffectsSlot)=0,{value}=INT({value}))),FALSE)")})
    for slot in range(1, 5):
        value = ref(f"$D${6 + slot}")
        names += [
            {"name": f"ParameterEffectsParameterOk{slot}",
             "refers": (f"=IFERROR(AND(ISNUMBER({value}),{value}>=INDEX(ExplorerMin,ParameterEffectsRow,{slot}),"
                        f"{value}<=INDEX(ExplorerMax,ParameterEffectsRow,{slot}),"
                        f"OR(INDEX(ExplorerInt,ParameterEffectsRow,{slot})=0,{value}=INT({value}))),FALSE)")},
            # The cell is greyed when its parameter is not used, or is the one that is varied.
            {"name": f"ParameterEffectsParameterUnused{slot}",
             "refers": f"=OR({slot}>INDEX(ExplorerNP,ParameterEffectsRow),IFERROR({slot}=ParameterEffectsSlot,FALSE))"},
        ]
    formulas = [
        {"address": STATUS.replace("$", ""), "formula": status_formula(families)},
        {"address": MACROS, "formula": '=IF(ISERROR(K_STATS_NormalStandard_Cumulative(0)+0*COUNTA($D$6:$D$16)),"DISABLED","ENABLED")'},
        {"address": EFFECT_CELL, "formula": '=IFERROR(INDEX(ParameterEffectsEffects,ParameterEffectsRow,ParameterEffectsSlot),"")'},
        {"address": "F4", "formula": f'=IF(${MACROS[0]}${MACROS[1:]}="ENABLED","","MACROS ARE DISABLED: the results are not live")'},
        {"address": f"{X_COLUMN}{GRID_FIRST_ROW}", "formula": x_grid_formula()},
    ]
    for k in range(1, CURVES + 1):
        row = FIRST_CURVE_ROW + k - 1
        formulas += [
            {"address": f"{DENSITY_COLUMNS[k - 1]}{GRID_FIRST_ROW}", "formula": grid_formula(families, k, "density")},
            {"address": f"{CDF_COLUMNS[k - 1]}{GRID_FIRST_ROW}", "formula": grid_formula(families, k, "cdf")},
            {"address": f"{WINDOW_COLUMNS[0]}{row}", "formula": window_formula(families, k, "plot_lower")},
            {"address": f"{WINDOW_COLUMNS[1]}{row}", "formula": window_formula(families, k, "plot_upper")},
            {"address": f"{LEGEND_COLUMN}{row}", "formula": f'={VARY}&" = "&{CURVE_VALUES[k - 1]}'},
            {"address": f"C{row}", "formula": f'="Curve {k}: "&{VARY}'},
            {"address": f"D{row}", "formula": f"={CURVE_VALUES[k - 1]}"},
        ]
        formulas += stat_formulas(families, data, k)
    for slot in range(4):
        formulas.append({"address": f"C{7 + slot}",
                         "formula": f'=IFERROR(INDEX(ExplorerLabels,ParameterEffectsRow,{slot + 1}),"")'})
    texts = [
        {"address": "B1", "value": "Parameter Effects"},
        {"address": "B2", "value": "How one parameter at a time moves a distribution: location, scale, shape or degrees of freedom."},
        {"address": "C3", "value": "Choose a distribution and the parameter to vary; four curves take its four values and the others "
                                   "stay at their base values. Every curve is a call to the embedded VBA API (K_STATS_ functions). "
                                   "Macros are required."},
        {"address": "C11", "value": "Parameter to vary"},
        {"address": "C12", "value": "Curve 1 value"},
        {"address": "C13", "value": "Curve 2 value"},
        {"address": "C14", "value": "Curve 3 value"},
        {"address": "C15", "value": "Curve 4 value"},
        {"address": "C16", "value": "Number of decimals"},
        {"address": "C19", "value": "Input status"},
        {"address": "C20", "value": "Macros"},
        {"address": f"C{HEADER_ROW}", "value": "Curve"},
        {"address": f"D{HEADER_ROW}", "value": "Value"},
        {"address": f"F{HEADER_ROW}", "value": "Mean"},
        {"address": f"G{HEADER_ROW}", "value": "Std deviation"},
        {"address": f"H{HEADER_ROW}", "value": "Skewness"},
        {"address": f"I{HEADER_ROW}", "value": f"P(X > mean + {TAIL_SIGMAS} sd)"},
        {"address": "C26", "value": "What the parameter does"},
        {"address": "F5", "value": "Read the curves"},
        {"address": "F6", "value": "Yellow cells are editable. The base parameters are fixed; the parameter to vary takes the four values. "
                                    "A grey base cell is the varied one: its base value is not used."},
        {"address": "F8", "value": "Dispersion is the standard deviation: wider curves have lower peaks. Skewness is zero for a symmetric "
                                    "curve and positive when the right tail is the longer."},
        {"address": "F10", "value": f"Tails: the last column is P(X > mean + {TAIL_SIGMAS} sd) (0.00135 for the Normal). A larger value "
                                     "means a heavier right tail. A moment that does not exist reads undefined or infinite."},
        {"address": "F12", "value": "All curves share one x axis, the union of their plot windows, and one scale. A discrete family "
                                     "draws probability mass and the cumulative steps at the integers."},
        {"address": "F14", "value": "Educational demo, not an accuracy certification. The moments are the closed forms of the parameters, "
                                     "written in the sheet; the curves are the library's."},
        {"address": "F16", "value": "Needs macros and Excel 365 or 2024. If Macros says DISABLED, unblock the file and reopen."},
    ]
    return {
        "schema": "parameter-effects-plan-1",
        "sheet": SHEET,
        "names": names,
        "formulas": formulas,
        "texts": texts,
        "table": {"cells": table_cells(families, data), "rows": rows},
        "grid": {"first_row": GRID_FIRST_ROW, "last_row": last_row, "x_column": X_COLUMN,
                 "density_columns": DENSITY_COLUMNS, "cdf_columns": CDF_COLUMNS},
        "decimals": {"cell": DECIMALS, "default": 6, "max": E.MAX_DECIMALS,
                     "cells": f"F{FIRST_CURVE_ROW}:I{FIRST_CURVE_ROW + CURVES - 1}"},
        "messages": static_messages(defaults),
        "layout": {"family": "D6", "vary": VARY.replace("$", ""), "curve_values": [v.replace("$", "") for v in CURVE_VALUES],
                   "decimals": DECIMALS.replace("$", ""), "status": STATUS.replace("$", ""), "macros": MACROS,
                   "inputs": "D6:D16", "outputs": [f"{STATUS.replace('$', '')}:{MACROS}", f"C{FIRST_CURVE_ROW}:I{FIRST_CURVE_ROW + CURVES - 1}",
                                                    f"{EFFECT_CELL}:D32", f"{X_COLUMN}{GRID_FIRST_ROW}:{LEGEND_COLUMN}{last_row}"],
                   "header_row": HEADER_ROW, "first_curve_row": FIRST_CURVE_ROW, "effect": "C27:D32",
                   "hidden_from": "Q"},
        "validation": {
            "family_list": "=ExplorerFamily",
            "parameter": "=ParameterEffectsParameterOk{slot}",
            "vary_list": "=ParameterEffectsChoices",
            "curve": "=ParameterEffectsCurveOk{k}",
            "decimals_list": "=ExplorerDecimals",
        },
        "conditional": {
            "unused_parameter": "=ParameterEffectsParameterUnused{slot}",
            "status_bad": "=ParameterEffectsStatusBad",
        },
        "charts": [
            {"name": "ParameterEffectsDensity", "title": "Density f(x) of the four curves",
             "series": [{"name": f"{LEGEND_COLUMN}{FIRST_CURVE_ROW + k}", "x": "ParameterEffectsX",
                         "y": f"ParameterEffectsDensity{k + 1}"} for k in range(CURVES)]},
            {"name": "ParameterEffectsCdf", "title": "Cumulative F(x) of the four curves",
             "series": [{"name": f"{LEGEND_COLUMN}{FIRST_CURVE_ROW + k}", "x": "ParameterEffectsX",
                         "y": f"ParameterEffectsCdf{k + 1}"} for k in range(CURVES)]},
        ],
    }


def formula_lengths(result: dict[str, Any]) -> list[tuple[str, int]]:
    sizes = [(item["address"], len(item["formula"])) for item in result["formulas"]]
    sizes += [(item["name"], len(item["refers"])) for item in result["names"]]
    return sorted(sizes, key=lambda item: -item[1])


if __name__ == "__main__":
    for name, size in formula_lengths(plan())[:8]:
        print(size, name)
