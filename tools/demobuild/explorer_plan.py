"""The Explorer sheet as data: every formula, name, validation rule and text.

examples/demo/explorer-families.json is the one source. This module writes from
it the plan that tools/demobuild/Evolve-Explorer.ps1 applies in desktop Excel,
and the unit tests check the same plan against the library's public API. A cell
of the sheet is therefore never typed twice: the family table, the dispatch
formulas, the input checks and the plot window all come from the same rows.

Layout (the user's workbook, v3): in C:D the inputs D6:D16 (distribution, up
to four parameters, evaluation point, probability p, plot lower and upper
bounds, number of points, number of decimals), then the status D19, the
macros cell D20, the results D21:D25 (density, cumulative and survival at x,
quantile at p, right tail at the upper bound) and the support text C26. The
notes are in F5:I17; the x / density / survival / cumulative grid starts at
F20, under its headings in row 19, with one dynamic-array formula per column.
The family table sits in hidden columns from LZ.
"""
from __future__ import annotations

import json
import re
from typing import Any
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
FAMILIES_FILE = ROOT / "examples" / "demo" / "explorer-families.json"

SHEET = "Explorer"
FIRST_TABLE_COLUMN = 338            # LZ
GRID_FIRST_ROW = 20
PARAMETER_CELLS = ["$D$7", "$D$8", "$D$9", "$D$10"]
LETTERS = "abcd"
PROBABILITY = "$D$12"
LOWER, UPPER, POINTS, DECIMALS = "$D$13", "$D$14", "$D$15", "$D$16"
STATUS = "$D$19"                    # "OK", or what to check
MACROS, RESULTS, QUANTILE, TAIL, SUPPORT = "D20", "D21", "D24", "D25", "C26"
MAX_DECIMALS = 15
DECIMALS_COLUMN = "NB"              # 0 to 15, the list behind the decimals cell
UNUSED_LABEL = "(not used)"
UNUSED_MIN, UNUSED_MAX = -1e15, 1e15

# Table columns, as offsets from FIRST_TABLE_COLUMN.
COLUMNS = {
    "family": 0, "kind": 1, "count": 2,
    "label": 3, "min": 7, "max": 11, "integer": 15, "default": 19,
    "x_default": 23, "suggest_lower": 24, "suggest_upper": 25, "check": 26, "support": 27,
    # 28 is the list of decimals (DECIMALS_COLUMN); the messages of the parameter cells follow.
    "prompt": 29, "error_title": 33, "error": 37,
}
NAMES = {
    "ExplorerFamily": ("family", 1), "ExplorerKind": ("kind", 1), "ExplorerNP": ("count", 1),
    "ExplorerLabels": ("label", 4), "ExplorerMin": ("min", 4), "ExplorerMax": ("max", 4),
    "ExplorerInt": ("integer", 4), "ExplorerDefaults": ("default", 4), "ExplorerX0": ("x_default", 1),
    "ExplorerSuggestLower": ("suggest_lower", 1), "ExplorerSuggestUpper": ("suggest_upper", 1),
    "ExplorerCheck": ("check", 1), "ExplorerSupport": ("support", 1),
    "ExplorerPrompts": ("prompt", 4), "ExplorerErrorTitles": ("error_title", 4), "ExplorerErrors": ("error", 4),
}
# Excel's limits for the messages of a data validation rule.
TITLE_LIMIT, PROMPT_LIMIT, ERROR_LIMIT = 32, 255, 225
UNUSED_MESSAGES = ("Not used by this distribution: the value is ignored.", "Not used",
                   "This distribution has no parameter here: the value is ignored.")


def column_letters(number: int) -> str:
    letters = ""
    while number:
        number, rest = divmod(number - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def load() -> dict[str, Any]:
    data: dict[str, Any] = json.loads(FAMILIES_FILE.read_text(encoding="utf-8"))
    for family in data["families"]:
        count = len(family["parameters"])
        if not 1 <= count <= 4:
            raise ValueError(f'{family["name"]}: one to four parameters')
        if family["call"].split(",") != list(LETTERS[:count]):
            raise ValueError(f'{family["name"]}: call must list a.. for its {count} parameters')
    names = [family["name"] for family in data["families"]]
    if len(names) != len(set(names)):
        raise ValueError("duplicate family")
    return data


def cell(family_index: int, key: str, offset: int = 0) -> str:
    """Absolute address of a table cell: row 1 is the first family."""
    return f"${column_letters(FIRST_TABLE_COLUMN + COLUMNS[key] + offset)}${family_index + 1}"


def sheet_expression(expression: str) -> str:
    """An expression written in the parameters a, b, c, d, as the sheet's cell references."""
    return re.sub(r"\b([abcd])\b", lambda found: PARAMETER_CELLS[LETTERS.index(found.group(1))], expression)


def number(value: float) -> str:
    return repr(int(value)) if float(value).is_integer() else repr(value)


def shown(value: float) -> str:
    """A limit as the messages show it: 10,000 and 0.001, never 10000.0."""
    return f"{int(value):,}" if float(value).is_integer() else f"{value:g}"


def parameter_messages(parameter: dict[str, Any]) -> tuple[str, str, str]:
    """Input message, error title and error message of one parameter, in the K-SACCR style:
    what the value is and its unit, its limits, and what to enter instead."""
    kind = "a whole number" if parameter["integer"] else "a number"
    limits = f'from {shown(parameter["min"])} to {shown(parameter["max"])}'
    note = f' {parameter["note"]}' if parameter.get("note") else ""
    prompt = f'{parameter["help"]} From {limits.split(" ", 1)[1]}' + (", a whole number." if parameter["integer"] else ".") + note
    label = parameter["label"]
    title = "Invalid " + (label if label[:2].isupper() else label[0].lower() + label[1:])
    return prompt, title, f"Enter {kind} {limits}.{note}"


def static_messages(defaults: dict[str, Any]) -> dict[str, list[str]]:
    """Input title, input message, error title and error message of the fixed input cells."""
    bound = shown(defaults["bound_limit"])
    x_limit = shown(defaults["x_limit"])
    points = shown(defaults["max_points"])
    return {
        "D6": ["Distribution", "Choose one of the 16 distributions. Its parameters, probe point and plot window "
               "are loaded with it.", "Unknown distribution", "Choose a distribution from the list."],
        "D11": ["Evaluation point x / k", "Point where the density or mass, the CDF and P(X > x) are evaluated, "
                f"in units of x. A discrete family truncates it to a whole number. From -{x_limit} to {x_limit}.",
                "Invalid evaluation point", f"Enter a number from -{x_limit} to {x_limit}."],
        "D12": ["Probability p", "Probability for the quantile: the result is the smallest x with CDF(x) >= p. "
                "From 0 to 1; 0 and 1 give the ends of the support.",
                "Invalid probability", "Enter a probability from 0 to 1."],
        "D13": ["Plot lower bound", "Left end of the grid and the charts, in units of x. Reset when the "
                f"distribution or a parameter changes. From -{bound} to {bound}, below the upper bound.",
                "Invalid lower bound", f"Enter a number from -{bound} to {bound}."],
        "D14": ["Plot upper bound", "Right end of the grid and the charts, in units of x. Reset when the "
                f"distribution or a parameter changes. From -{bound} to {bound}, above the lower bound.",
                "Invalid upper bound", f"Enter a number from -{bound} to {bound}."],
        "D15": ["Number of points", f"Points of a continuous grid, evenly spaced between the bounds, from 2 to "
                f"{points}. A discrete family plots every integer between the bounds and ignores this value.",
                "Invalid number of points", f"Enter a whole number from 2 to {points}."],
        "D16": ["Number of decimals", "Decimals shown in the results and the grid, from 0 to 15. Display only: "
                "the values keep their full precision.",
                "Invalid number of decimals", "Choose a number of decimals from 0 to 15."],
    }


def call(family: dict[str, Any], operation: str, point: str) -> str:
    arguments = [point] + [PARAMETER_CELLS[LETTERS.index(letter)] for letter in family["call"].split(",")]
    return f'{family["functions"][operation]}({",".join(arguments)})'


def dispatch(families: list[dict[str, Any]], operation: str, point: str) -> str:
    """The selected family's call, as one SWITCH over the distribution cell."""
    branches = ",".join(f'"{family["name"]}",{call(family, operation, point)}' for family in families)
    return f"SWITCH($D$6,{branches},NA())"


def table_cells(families: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Every cell of the family table: constants, and the formulas that read the inputs."""
    cells: list[dict[str, Any]] = []

    def put(index: int, key: str, offset: int, value: Any = None, formula: str | None = None) -> None:
        address = cell(index, key, offset).replace("$", "")
        cells.append({"address": address, **({"value": value} if formula is None else {"formula": formula})})

    for index, family in enumerate(families):
        count = len(family["parameters"])
        put(index, "family", 0, family["name"])
        put(index, "kind", 0, family["kind"].capitalize())
        put(index, "count", 0, count)
        for slot in range(4):
            used = slot < count
            parameter: dict[str, Any] = family["parameters"][slot] if used else {}
            put(index, "label", slot, parameter["label"] if used else UNUSED_LABEL)
            put(index, "min", slot, parameter["min"] if used else UNUSED_MIN)
            put(index, "max", slot, parameter["max"] if used else UNUSED_MAX)
            put(index, "integer", slot, int(parameter["integer"]) if used else 0)
            put(index, "default", slot, parameter["default"] if used else 0)
            prompt, error_title, error = parameter_messages(parameter) if used else UNUSED_MESSAGES
            put(index, "prompt", slot, prompt)
            put(index, "error_title", slot, error_title)
            put(index, "error", slot, error)
        put(index, "x_default", 0, family["x_default"])
        put(index, "suggest_lower", 0, formula=f'=IFERROR({sheet_expression(family["plot_lower"])},"")')
        put(index, "suggest_upper", 0, formula=f'=IFERROR({sheet_expression(family["plot_upper"])},"")')
        put(index, "check", 0, formula=check_formula(index, family))
        put(index, "support", 0, family["support"])
    return cells


def check_formula(index: int, family: dict[str, Any]) -> str:
    """The first parameter that is not acceptable for this family, else "". Read at G5."""
    tests = []
    for slot, parameter in enumerate(family["parameters"]):
        reference = PARAMETER_CELLS[slot]
        parts = [f"ISNUMBER({reference})", f'{reference}>={cell(index, "min", slot)}',
                 f'{reference}<={cell(index, "max", slot)}']
        if parameter["integer"]:
            parts.append(f"{reference}=INT({reference})")
        tests.append((parameter["label"], f'IFERROR(AND({",".join(parts)}),FALSE)'))
    if family["constraint"] != "TRUE":
        tests.append((family["constraint_text"], f'IFERROR({sheet_expression(family["constraint"])},FALSE)'))
    formula = '""'
    for message, test in reversed(tests):
        formula = f'IF(NOT({test}),"{message}",{formula})'
    return "=" + formula


def x_grid_formula() -> str:
    return (
        f'=IF({STATUS}="OK",IF(INDEX(ExplorerKind,ExplorerRow)="Discrete",'
        f'LET(first,-INT(-{LOWER}),SEQUENCE(INT({UPPER})-first+1,,first)),'
        f'LET(i,SEQUENCE({POINTS},,0),{LOWER}+({UPPER}-{LOWER})*i/({POINTS}-1))),NA())'
    )


def grid_formula(families: list[dict[str, Any]], operation: str) -> str:
    return (f'=IF({STATUS}="OK",MAP(ExplorerX,LAMBDA(x,IFERROR({dispatch(families, operation, "x")},NA()))),NA())')


def results_formula(families: list[dict[str, Any]]) -> str:
    stack = ",".join(dispatch(families, operation, "$D$11") for operation in ("density", "cdf", "survival"))
    return f'=IF({STATUS}="OK",VSTACK({stack}),EXPAND(NA(),3,,NA()))'


def quantile_formula(families: list[dict[str, Any]]) -> str:
    """The quantile at p: only p is checked here, so a bad p leaves the other results alone."""
    return (f'=IF(AND({STATUS}="OK",ISNUMBER({PROBABILITY}),{PROBABILITY}>=0,{PROBABILITY}<=1),'
            f'{dispatch(families, "quantile", PROBABILITY)},NA())')


def status_formula(defaults: dict[str, Any]) -> str:
    x_limit = number(defaults["x_limit"])
    bound = number(defaults["bound_limit"])
    max_points = number(defaults["max_points"])
    max_discrete = number(defaults["max_discrete_points"] - 1)
    return (
        '=IFERROR(LET(message,INDEX(ExplorerCheck,ExplorerRow),'
        'IF(message<>"","Check: "&message,'
        f'IF(NOT(AND(ISNUMBER($D$11),ABS($D$11)<={x_limit})),"Check: evaluation point",'
        f'IF(NOT(AND(ISNUMBER({LOWER}),ISNUMBER({UPPER}),ABS({LOWER})<={bound},ABS({UPPER})<={bound})),"Check: plot window",'
        'IF(INDEX(ExplorerKind,ExplorerRow)="Discrete",'
        f'IF(AND({UPPER}>={LOWER},INT({UPPER})>=-INT(-{LOWER}),INT({UPPER})--INT(-{LOWER})<={max_discrete}),"OK","Check: plot window"),'
        f'IF({LOWER}>={UPPER},"Check: plot window",IF(NOT(AND(ISNUMBER({POINTS}),{POINTS}=INT({POINTS}),{POINTS}>=2,{POINTS}<={max_points})),'
        '"Check: number of points","OK"))))))),"Check: distribution")'
    )


def plan() -> dict[str, Any]:
    data = load()
    families, defaults = data["families"], data["defaults"]
    rows = len(families)
    names = []
    for name, (key, width) in NAMES.items():
        first = f"${column_letters(FIRST_TABLE_COLUMN + COLUMNS[key])}$1"
        last = f"${column_letters(FIRST_TABLE_COLUMN + COLUMNS[key] + width - 1)}${rows}"
        names.append({"name": name, "refers": f"={SHEET}!{first}:{last}"})
    names += [
        {"name": "ExplorerRow", "refers": f"=MATCH({SHEET}!$D$6,ExplorerFamily,0)"},
        {"name": "ExplorerX", "refers": f"={SHEET}!$F${GRID_FIRST_ROW}#"},
        {"name": "ExplorerPDF", "refers": f"={SHEET}!$G${GRID_FIRST_ROW}#"},
        {"name": "ExplorerSurvival", "refers": f"={SHEET}!$H${GRID_FIRST_ROW}#"},
        {"name": "ExplorerCDF", "refers": f"={SHEET}!$I${GRID_FIRST_ROW}#"},
        {"name": "ExplorerDecimals",
         "refers": f"={SHEET}!${DECIMALS_COLUMN}$1:${DECIMALS_COLUMN}${MAX_DECIMALS + 1}"},
    ]
    # The rules of validation and conditional formats are named formulas: a
    # formula passed to Excel through COM is read in the user's list separator
    # and function names, a name's RefersTo is not.
    for slot in range(1, 5):
        value = f"{SHEET}!$D${6 + slot}"
        names += [
            {"name": f"ExplorerParameterOk{slot}",
             "refers": (f"=IFERROR(AND(ISNUMBER({value}),{value}>=INDEX(ExplorerMin,ExplorerRow,{slot}),"
                        f"{value}<=INDEX(ExplorerMax,ExplorerRow,{slot}),"
                        f"OR(INDEX(ExplorerInt,ExplorerRow,{slot})=0,{value}=INT({value}))),FALSE)")},
            {"name": f"ExplorerParameterUnused{slot}", "refers": f"={slot}>INDEX(ExplorerNP,ExplorerRow)"},
        ]
    names += [
        {"name": "ExplorerStatusBad", "refers": f'={SHEET}!{STATUS}<>"OK"'},
    ]
    # The grid holds the larger plot window: continuous points or discrete integers.
    grid_rows = max(int(defaults["max_points"]), int(defaults["max_discrete_points"]))
    last_row = GRID_FIRST_ROW + grid_rows - 1
    formulas = [
        {"address": STATUS.replace("$", ""), "formula": status_formula(defaults)},
        {"address": MACROS, "formula": '=IF(ISERROR(K_STATS_NormalStandard_Cumulative(0)+0*COUNTA($D$6:$D$12)),"DISABLED","ENABLED")'},
        {"address": RESULTS, "formula": results_formula(families)},
        {"address": QUANTILE, "formula": quantile_formula(families)},
        {"address": TAIL, "formula": f'=IF({STATUS}="OK",{dispatch(families, "survival", UPPER)},NA())'},
        {"address": SUPPORT, "formula": '=IFERROR(INDEX(ExplorerSupport,ExplorerRow),"")'},
        {"address": "F4", "formula": f'=IF(${MACROS[0]}${MACROS[1:]}="ENABLED","","MACROS ARE DISABLED: the results are not live")'},
        {"address": f"F{GRID_FIRST_ROW}", "formula": x_grid_formula()},
        {"address": f"G{GRID_FIRST_ROW}", "formula": grid_formula(families, "density")},
        {"address": f"H{GRID_FIRST_ROW}", "formula": grid_formula(families, "survival")},
        {"address": f"I{GRID_FIRST_ROW}", "formula": grid_formula(families, "cdf")},
    ]
    for slot in range(4):
        formulas.append({"address": f"C{7 + slot}", "formula": f'=IFERROR(INDEX(ExplorerLabels,ExplorerRow,{slot + 1}),"")'})
    texts = [
        {"address": "C3", "value": "Choose a distribution and edit the highlighted inputs. Every result is a call to the embedded "
                                   "VBA API (K_STATS_ functions), spilled as a dynamic array. Macros are required."},
        {"address": "C11", "value": "Evaluation point x / k"},
        {"address": "C12", "value": "Probability p"},
        {"address": "C13", "value": "Plot lower bound"},
        {"address": "C14", "value": "Plot upper bound"},
        {"address": "C15", "value": "Number of points"},
        {"address": "C16", "value": "Number of decimals"},
        {"address": "C19", "value": "Input status"},
        {"address": "C20", "value": "Macros"},
        {"address": "C21", "value": "PDF / PMF at x"},
        {"address": "C22", "value": "CDF at x"},
        {"address": "C23", "value": "P(X > x)"},
        {"address": "C24", "value": "Quantile at p"},
        {"address": "C25", "value": "P(X > upper bound)"},
        {"address": "F5", "value": "Read the results"},
        {"address": "F6", "value": "Yellow cells are editable. A new distribution loads its parameters and plot window."},
        {"address": "F8", "value": "Select a result or grid cell to read its formula: each grid column is one spilled formula."},
        {"address": "F10", "value": "Continuous: Number of points between the bounds. Discrete: every integer between them."},
        {"address": "F12", "value": "Reset inputs on the Ribbon restores Normal(0,1) on this sheet, after confirmation."},
        {"address": "F14", "value": "Educational demo, not an accuracy certification. The quantile is the smallest x with CDF >= p."},
        {"address": "F16", "value": "Needs macros and Excel 365 or 2024. If Macros says DISABLED, unblock the file and reopen."},
        {"address": f"F{GRID_FIRST_ROW - 1}", "value": "x / integer k"},
        {"address": f"G{GRID_FIRST_ROW - 1}", "value": "PDF / PMF"},
        {"address": f"H{GRID_FIRST_ROW - 1}", "value": "P(X > x)"},
        {"address": f"I{GRID_FIRST_ROW - 1}", "value": "CDF"},
    ]
    return {
        "schema": "explorer-plan-1",
        "sheet": SHEET,
        "families": families,
        "defaults": defaults,
        "table": {"cells": table_cells(families) + [{"address": f"{DECIMALS_COLUMN}{n + 1}", "value": n}
                                                    for n in range(MAX_DECIMALS + 1)],
                  "rows": rows, "first_column": FIRST_TABLE_COLUMN},
        "names": names,
        "formulas": formulas,
        "texts": texts,
        "grid": {"first_row": GRID_FIRST_ROW, "last_row": last_row, "columns": ["F", "G", "H", "I"]},
        # Applied by ApplyDecimals in the Explorer module, which names the same cells.
        "decimals": {"cell": DECIMALS, "default": 6, "max": MAX_DECIMALS, "column": DECIMALS_COLUMN,
                     "cells": f"{RESULTS}:{TAIL},G{GRID_FIRST_ROW}:I{last_row}",
                     "x_cells": f"F{GRID_FIRST_ROW}:F{last_row}"},
        "probability": {"cell": PROBABILITY, "default": 0.95},
        "messages": static_messages(defaults),
        "layout": {"family": "D6", "x": "D11", "probability": PROBABILITY.replace("$", ""),
                   "lower": LOWER.replace("$", ""), "upper": UPPER.replace("$", ""),
                   "points": POINTS.replace("$", ""), "decimals": DECIMALS.replace("$", ""),
                   "status": STATUS.replace("$", ""), "macros": MACROS, "results": RESULTS,
                   "quantile": QUANTILE, "tail": TAIL, "support": SUPPORT,
                   "inputs": "D6:D16", "outputs": [f"{STATUS.replace('$', '')}:{TAIL}"]},
        "validation": {
            "family_list": "=ExplorerFamily",
            "parameter": "=ExplorerParameterOk{slot}",
            "x_limit": defaults["x_limit"], "bound_limit": defaults["bound_limit"],
            "points": [2, defaults["max_points"]],
            "decimals_list": "=ExplorerDecimals",
            "probability": [0, 1],
        },
        "conditional": {
            "unused_parameter": "=ExplorerParameterUnused{slot}",
            "status_bad": "=ExplorerStatusBad",
            "grid_inside": f"=ROW()-{GRID_FIRST_ROW - 1}<=ROWS(ExplorerX)",
        },
    }


if __name__ == "__main__":
    print(json.dumps(plan(), indent=1)[:2000])
