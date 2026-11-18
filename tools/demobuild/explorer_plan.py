"""The Explorer sheet as data: every formula, name, validation rule and text.

examples/demo/explorer-families.json is the one source. This module writes from
it the plan that tools/demobuild/Evolve-Explorer.ps1 applies in desktop Excel,
and the unit tests check the same plan against the library's public API. A cell
of the sheet is therefore never typed twice: the family table, the dispatch
formulas, the input checks and the plot window all come from the same rows.

Layout (the user's workbook, v1): inputs D6:D11 (distribution, up to four
parameters, evaluation point), plot window G10:G12, results G7:G9 and G13, the
x / density / cumulative grid in F19:H19 with one dynamic-array formula per
column. The family table sits in hidden columns from MA.
"""
from __future__ import annotations

import json
import re
from typing import Any
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
FAMILIES_FILE = ROOT / "examples" / "demo" / "explorer-families.json"

SHEET = "Explorer"
FIRST_TABLE_COLUMN = 339            # MA
GRID_FIRST_ROW = 19
GRID_MAX_ROWS = 1000                # the largest plot window: points and discrete counts
PARAMETER_CELLS = ["$D$7", "$D$8", "$D$9", "$D$10"]
LETTERS = "abcd"
UNUSED_LABEL = "(not used)"
UNUSED_MIN, UNUSED_MAX = -1e15, 1e15

# Table columns, as offsets from FIRST_TABLE_COLUMN.
COLUMNS = {
    "family": 0, "kind": 1, "count": 2,
    "label": 3, "min": 7, "max": 11, "integer": 15, "default": 19,
    "x_default": 23, "suggest_lower": 24, "suggest_upper": 25, "check": 26, "support": 27,
}
NAMES = {
    "ExplorerFamily": ("family", 1), "ExplorerKind": ("kind", 1), "ExplorerNP": ("count", 1),
    "ExplorerLabels": ("label", 4), "ExplorerMin": ("min", 4), "ExplorerMax": ("max", 4),
    "ExplorerInt": ("integer", 4), "ExplorerDefaults": ("default", 4), "ExplorerX0": ("x_default", 1),
    "ExplorerSuggestLower": ("suggest_lower", 1), "ExplorerSuggestUpper": ("suggest_upper", 1),
    "ExplorerCheck": ("check", 1), "ExplorerSupport": ("support", 1),
}


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
        '=IF($G$5="OK",IF(INDEX(ExplorerKind,ExplorerRow)="Discrete",'
        'LET(first,-INT(-$G$10),SEQUENCE(INT($G$11)-first+1,,first)),'
        'LET(i,SEQUENCE($G$12,,0),$G$10+($G$11-$G$10)*i/($G$12-1))),NA())'
    )


def grid_formula(families: list[dict[str, Any]], operation: str) -> str:
    return (f'=IF($G$5="OK",MAP(ExplorerX,LAMBDA(x,IFERROR({dispatch(families, operation, "x")},NA()))),NA())')


def results_formula(families: list[dict[str, Any]]) -> str:
    stack = ",".join(dispatch(families, operation, "$D$11") for operation in ("density", "cdf", "survival"))
    return f'=IF($G$5="OK",VSTACK({stack}),EXPAND(NA(),3,,NA()))'


def status_formula(defaults: dict[str, Any]) -> str:
    x_limit = number(defaults["x_limit"])
    bound = number(defaults["bound_limit"])
    max_points = number(defaults["max_points"])
    max_discrete = number(defaults["max_discrete_points"] - 1)
    return (
        '=IFERROR(LET(message,INDEX(ExplorerCheck,ExplorerRow),'
        'IF(message<>"","Check: "&message,'
        f'IF(NOT(AND(ISNUMBER($D$11),ABS($D$11)<={x_limit})),"Check: evaluation point",'
        f'IF(NOT(AND(ISNUMBER($G$10),ISNUMBER($G$11),ABS($G$10)<={bound},ABS($G$11)<={bound})),"Check: plot window",'
        'IF(INDEX(ExplorerKind,ExplorerRow)="Discrete",'
        f'IF(AND($G$11>=$G$10,INT($G$11)>=-INT(-$G$10),INT($G$11)--INT(-$G$10)<={max_discrete}),"OK","Check: plot window"),'
        f'IF($G$10>=$G$11,"Check: plot window",IF(NOT(AND(ISNUMBER($G$12),$G$12=INT($G$12),$G$12>=2,$G$12<={max_points})),'
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
        {"name": "ExplorerCDF", "refers": f"={SHEET}!$H${GRID_FIRST_ROW}#"},
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
        {"name": "ExplorerStatusBad", "refers": f'={SHEET}!$G$5<>"OK"'},
        {"name": "ExplorerGridDiscrete", "refers": '=INDEX(ExplorerKind,ExplorerRow)="Discrete"'},
    ]
    last_row = GRID_FIRST_ROW + GRID_MAX_ROWS - 1
    formulas = [
        {"address": "G5", "formula": status_formula(defaults)},
        {"address": "G6", "formula": '=IF(ISERROR(K_STATS_NormalStandard_Cumulative(0)+0*COUNTA($D$6:$D$11)),"DISABLED","ENABLED")'},
        {"address": "G7", "formula": results_formula(families)},
        {"address": "G13", "formula": f'=IF($G$5="OK",{dispatch(families, "survival", "$G$11")},NA())'},
        {"address": "F14", "formula": '=IFERROR(INDEX(ExplorerSupport,ExplorerRow),"")'},
        {"address": f"F{GRID_FIRST_ROW}", "formula": x_grid_formula()},
        {"address": f"G{GRID_FIRST_ROW}", "formula": grid_formula(families, "density")},
        {"address": f"H{GRID_FIRST_ROW}", "formula": grid_formula(families, "cdf")},
    ]
    for slot in range(4):
        formulas.append({"address": f"C{7 + slot}", "formula": f'=IFERROR(INDEX(ExplorerLabels,ExplorerRow,{slot + 1}),"")'})
    texts = [
        {"address": "C3", "value": "Choose a distribution and edit the highlighted inputs. Every result is a call to the embedded "
                                   "VBA API (K_STATS_ functions), spilled as a dynamic array. Macros are required."},
        {"address": "F7", "value": "PDF / PMF at x"},
        {"address": "F8", "value": "CDF at x"},
        {"address": "F9", "value": "P(X > x)"},
        {"address": "F10", "value": "Plot lower bound"},
        {"address": "F11", "value": "Plot upper bound"},
        {"address": "F12", "value": "Number of points"},
        {"address": "F13", "value": "P(X > upper bound)"},
        {"address": "C19", "value": "Yellow cells are editable. A new distribution loads its parameters and plot window."},
        {"address": "C21", "value": "Select a result or grid cell to read its formula: each grid column is one spilled formula."},
        {"address": "C23", "value": "Continuous: Number of points between the bounds. Discrete: every integer between them."},
        {"address": "C25", "value": "Reset Explorer inputs on the Ribbon restores Normal(0,1), after confirmation."},
        {"address": "C27", "value": "Educational demo, not an accuracy certification. Quantiles are not offered here."},
        {"address": "C29", "value": "Needs macros and Excel 365 or 2024. If Macros says DISABLED, unblock the file and reopen."},
    ]
    return {
        "schema": "explorer-plan-1",
        "sheet": SHEET,
        "families": families,
        "defaults": defaults,
        "table": {"cells": table_cells(families), "rows": rows, "first_column": FIRST_TABLE_COLUMN},
        "names": names,
        "formulas": formulas,
        "texts": texts,
        "grid": {"first_row": GRID_FIRST_ROW, "last_row": last_row, "columns": ["F", "G", "H"]},
        "validation": {
            "family_list": "=ExplorerFamily",
            "parameter": "=ExplorerParameterOk{slot}",
            "x_limit": defaults["x_limit"], "bound_limit": defaults["bound_limit"],
            "points": [2, defaults["max_points"]],
        },
        "conditional": {
            "unused_parameter": "=ExplorerParameterUnused{slot}",
            "status_bad": "=ExplorerStatusBad",
            "grid_discrete": "=ExplorerGridDiscrete",
            "grid_inside": f"=ROW()-{GRID_FIRST_ROW - 1}<=ROWS(ExplorerX)",
        },
    }


if __name__ == "__main__":
    print(json.dumps(plan(), indent=1)[:2000])
