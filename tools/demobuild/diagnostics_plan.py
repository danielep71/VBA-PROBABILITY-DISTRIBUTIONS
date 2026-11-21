"""The Diagnostics sheet (errors and diagnostics) as data: every formula, name, validation rule and text.

Same method as risk_plan.py: this module writes the plan that tools/demobuild/Evolve-Diagnostics.ps1 applies in
desktop Excel. examples/demo/diagnostics-cases.json (made by examples/demo/diagnostics_references.py) holds the
catalogue of functions, the case table with the outcome the documentation gives each call, and the trials that
ci/Test-Diagnostics.ps1 types into the try-it area. Every value on the sheet is a live call to the library; the
Status text comes from the library through DiagnosticsStatus, a function of the sheet's VBA module that makes
the same call with a Status argument and returns what the library wrote there.

Layout, in the style of the Explorer:

    C6:D12    the try-it inputs: the function (a list), its arguments 1 to 4 (any value, the point of the sheet:
              the label of each names the parameter), and the range over which the chart sweeps argument 1
    D19:D20   the input status and the macros
    C21:I33   the results: the call with its values, the value returned (and in F:I its formula), the kind of
              that value, the Status the library wrote, what they mean and what to do, the domain of the
              function, the arguments Excel converted before the call, the sweep, the case table in one line,
              what the sheet does not show, the Excel that computed it and the provenance
    K5:O35    one chart: the function across the sweep of argument 1, a gap in the line at each error
    C40:O70   the case table: per call, the call as written, what it returns, what the documentation says it
              returns, a check (As documented or UNEXPECTED), the Status and the cause with its correction
    AE:AX     hidden: an empty cell, the index of the function, the spilled sweep and the catalogue
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
CASES_FILE = ROOT / "examples" / "demo" / "diagnostics-cases.json"

SHEET = "Diagnostics"
FUNCTION_CELL = "D6"
ARG_CELLS = ("D7", "D8", "D9", "D10")
SWEEP_CELLS = ("D11", "D12")
STATUS = "$D$19"
MACROS = "D20"
SWEEP_LIMIT = 1e15
LITERAL_LIMIT = 255
RESULT_ROWS = {
    "call": 21, "returned": 22, "kind": 23, "status": 24, "meaning": 25, "action": 26, "domain": 27,
    "converted": 28, "sweep": 29, "cases": 30, "not_shown": 31, "excel_build": 32, "provenance": 33,
}
HEADER_ROW = 40
FIRST_ROW = 41
# The case table: the call over D:G (E is the Explorer's narrow spacer), the status over L:M, the cause and
# its correction over N:O.
TABLE = {"title": "C", "call": "D", "returned": "H", "expected": "I", "check": "K", "status": "L", "cause": "N"}
TABLE_MERGES = ("D{row}:G{row}", "L{row}:M{row}", "N{row}:O{row}")
EMPTY_CELL = "AE1"
INDEX_CELL = "AE6"
ARITY_CELL = "AE7"
HELPERS = {"x": "AF", "raw": "AG", "value": "AH"}
CATALOGUE_FIRST = 6
CATALOGUE: dict[str, str] = {"name": "AL", "arity": "AM", "sweep_from": "AV", "sweep_to": "AW", "domain": "AX"}
CATALOGUE_PARAMS = ("AN", "AO", "AP", "AQ")
CATALOGUE_EXAMPLES = ("AR", "AS", "AT", "AU")
GROUPS = {"valid": "Valid", "limit": "Natural limit", "underflow": "Outside the Double range",
          "invalid": "Invalid on purpose", "excel": "Excel input"}
ERROR_NAMES = ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A")


def absolute(address: str) -> str:
    column = address.rstrip("0123456789")
    return f"${column}${address[len(column):]}"


def ref(address: str) -> str:
    return f"{SHEET}!{absolute(address)}"


def load() -> dict[str, Any]:
    data: dict[str, Any] = json.loads(CASES_FILE.read_text(encoding="utf-8"))
    names = [item["name"] for item in data["functions"]]
    if len(names) != len(set(names)):
        raise ValueError("duplicate function")
    if data["default_function"] not in names:
        raise ValueError("the default function is not in the catalogue")
    for item in data["functions"]:
        if not 1 <= len(item["params"]) <= len(ARG_CELLS) or len(item["example"]) != len(item["params"]):
            raise ValueError(f"{item['name']}: one to four parameters, one example value each")
    for case in data["cases"] + data["trials"]:
        if case["function"] not in names:
            raise ValueError(f"{case['function']} is not in the catalogue")
    if data["sweep_points"] < 2:
        raise ValueError("the sweep needs two points")
    return data


def literal(text: str) -> str:
    """A text as formula constants of at most LITERAL_LIMIT characters, joined with &, cut at spaces."""
    pieces, rest = [], text.replace('"', '""')
    while len(rest) > LITERAL_LIMIT:
        cut = rest.rfind(" ", 0, LITERAL_LIMIT)
        pieces.append(rest[:cut])
        rest = rest[cut:]
    pieces.append(rest)
    return "&".join(f'"{piece}"' for piece in pieces)


def number_text(value: float) -> str:
    """A number as a formula constant that gives back the same Double."""
    if float(value).is_integer() and abs(value) < 1e15:
        return str(int(value))
    return repr(float(value)).upper()


def argument_text(value: Any) -> str:
    """An argument of a case as it is written in the formula."""
    if isinstance(value, dict) and value.get("empty"):
        return "DiagEmpty"
    if isinstance(value, dict):
        return '"' + str(value["text"]).replace('"', '""') + '"'
    return number_text(value)


def call_text(function: str, args: list[Any]) -> str:
    return f"{function}({','.join(argument_text(value) for value in args)})"


def error_name(cell: str) -> str:
    """The name of the worksheet error a cell or a LET name holds, as a text."""
    names = ",".join(f'"{name}"' for name in ERROR_NAMES)
    return f'IFERROR(CHOOSE(ERROR.TYPE({cell}),{names}),"another error")'


def kind_of(cell: str) -> str:
    """The kind of a value: "a number", or the name of the worksheet error."""
    return f'IF(ISERROR({cell}),{error_name(cell)},"a number")'


def dispatch(functions: list[dict[str, Any]], arguments: tuple[str, ...]) -> str:
    """CHOOSE over the catalogue: each function called with as many of the arguments as it takes."""
    calls = [f"{item['name']}({','.join(arguments[:len(item['params'])])})" for item in functions]
    return f"CHOOSE(DiagIndex,{','.join(calls)})"


def ok() -> str:
    return f'{STATUS}="OK"'


def result(key: str) -> str:
    return f"D{RESULT_ROWS[key]}"


def check_names() -> list[dict[str, str]]:
    """One TRUE/FALSE name per checked input: the data validation and the status read the same rules. The
    arguments are not checked: what the library does with any value is what the sheet shows."""
    low, high = (ref(cell) for cell in SWEEP_CELLS)
    bound = f"{SWEEP_LIMIT:.0E}"
    return [
        {"name": "DiagFunctionOk", "refers": f"=ISNUMBER(MATCH({ref(FUNCTION_CELL)},DiagFunctionList,0))"},
        {"name": "DiagSweepFromOk", "refers": f"=IFERROR(AND(ISNUMBER({low}),ABS({low})<={bound}),FALSE)"},
        {"name": "DiagSweepToOk", "refers": f"=IFERROR(AND(ISNUMBER({high}),ABS({high})<={bound}),FALSE)"},
        {"name": "DiagSweepOrderOk", "refers": f"=IFERROR({low}<{high},FALSE)"},
    ]


CHECK_ORDER = (("DiagFunctionOk", "choose a function from the list"),
               ("DiagSweepFromOk", "sweep from (a number, at most 1E15 in size)"),
               ("DiagSweepToOk", "sweep to (a number, at most 1E15 in size)"),
               ("DiagSweepOrderOk", "sweep from must be below sweep to"))


def status_cell_formula() -> str:
    formula = '"OK"'
    for name, label in reversed(CHECK_ORDER):
        formula = f'IF(NOT({name}),"Check: {label}",{formula})'
    return "=" + formula


def parameter(slot: int) -> str:
    return f"INDEX(DiagParams,DiagIndex,{slot})"


def label_formula(slot: int) -> str:
    return (f'=IF(DiagFunctionOk,IF(DiagArity>={slot},"Argument {slot}: "&{parameter(slot)},'
            f'"Argument {slot} (not used)"),"Argument {slot}")')


def shown_argument(cell: str) -> str:
    """An argument as the call with its values shows it: a text in quotes, an empty cell named."""
    return f'IF(ISBLANK({cell}),"[empty cell]",IF(ISTEXT({cell}),""""&{cell}&"""",{cell}&""))'


MEANING = {
    "number": "A number: the call succeeded. A 0 or a 1 at the edge of a support is the natural limit, not an error.",
    "num": "#NUM!: the library understood the request and declined it. The Status above names the argument or the "
           "condition.",
    "excel": "#VALUE! from Excel: an argument is text or an error value, so the function never ran.",
    "value": "#VALUE!: an unexpected failure inside the library. Not expected for numeric input.",
    "other": "Another error: an argument holds it and Excel passes it on.",
}
ACTION = {
    "number": "Nothing to correct.",
    "num": "Correct the argument the Status names, within the domain below. Do not hide the error with IFERROR.",
    "excel": "Enter a number in place of the text or the error value.",
    "value": "Report it with the formula, the arguments and the library version (manual, chapter 9).",
    "other": "Correct the cell the error comes from.",
}
NOT_SHOWN = (
    "Non-convergence: an iterative calculation that does not settle returns #NUM! with a Status that says so; no "
    "input of the library's survey of bad input produces it, so no case shows it. Not shown either: "
    "K_STATS_NormalStandard_InverseCumulativeFast, the one function that validates nothing (a helper for "
    "simulation loops, documented so); the differences between functions recorded as not settled in #418; more "
    "than one wrong argument at a time; Excel builds other than the one named below."
)


def by_kind(texts: dict[str, str]) -> str:
    kind, status = result("kind"), result("status")
    return (f'=IF({ok()},IF({kind}="a number",{literal(texts["number"])},IF({kind}="#NUM!",{literal(texts["num"])},'
            f'IF(AND({kind}="#VALUE!",LEFT({status},10)="Not called"),{literal(texts["excel"])},'
            f'IF({kind}="#VALUE!",{literal(texts["value"])},{literal(texts["other"])})))),"")')


def converted_formula() -> str:
    """The arguments Excel converts before the call: an empty cell becomes 0, a text stops the call."""
    parts = []
    for slot, cell in enumerate(ARG_CELLS, 1):
        name = absolute(cell)
        used = f"DiagArity>={slot}"
        parts.append(f'IF(AND({used},ISBLANK({name})),"Argument {slot} ("&{parameter(slot)}&") is empty: Excel '
                     f'passes it as 0. ","")')
        parts.append(f'IF(AND({used},ISTEXT({name})),"Argument {slot} ("&{parameter(slot)}&") is text: unless VBA '
                     f'reads it as a number, the function never runs and Excel returns #VALUE!. ","")')
    joined = "&".join(parts)
    return f'=IF({ok()},LET(t,{joined},IF(t="","None: every argument is a number.",TRIM(t))),"")'


def sweep_formula(points: int) -> str:
    low, high = (absolute(cell) for cell in SWEEP_CELLS)
    return (f'=IF({ok()},LET(n,ROWS(DiagSweepRaw),e,SUM(--ISTEXT(DiagSweepRaw)),'
            f'm,SUM(--(DiagSweepRaw="#NUM!")),'
            f'"Over "&n&" values of "&{parameter(1)}&" from "&{low}&" to "&{high}&": "&(n-e)&" numbers, "&m&'
            f'" #NUM!"&IF(e>m," and "&(e-m)&" other errors","")&IF(e>0,"; the chart leaves a gap at each '
            f'error.",".")),"")')


def case_formulas(cases: list[dict[str, Any]]) -> list[dict[str, str]]:
    out = []
    last = FIRST_ROW + len(cases) - 1
    for offset, case in enumerate(cases):
        row = FIRST_ROW + offset
        call = call_text(case["function"], case["args"])
        returned, expected = f"{TABLE['returned']}{row}", f"{TABLE['expected']}{row}"
        status_args = ",".join(argument_text(value) for value in case["args"])
        if case["expected"]["kind"] == "number":
            check = (f'=IF(ISERROR({returned}),"UNEXPECTED",IF(ABS({returned}-{expected})<=1E-12*ABS({expected}),'
                     f'"As documented","UNEXPECTED"))')
            shown = "=" + number_text(case["expected"]["value"])
        else:
            check = (f'=IF(ISERROR({returned}),IF({kind_of(returned)}={expected},"As documented","UNEXPECTED"),'
                     f'"UNEXPECTED")')
            shown = f'="{case["expected"]["kind"]}"'
        out += [
            {"address": f"{TABLE['call']}{row}", "formula": '="=' + call.replace('"', '""') + '"'},
            {"address": returned, "formula": "=" + call},
            {"address": expected, "formula": shown},
            {"address": f"{TABLE['check']}{row}", "formula": check},
            {"address": f"{TABLE['status']}{row}",
             "formula": f'=LET(s,DiagnosticsStatus("{case["function"]}",{status_args}),IF(s="","(empty)",s))'},
        ]
    checks = f"{TABLE['check']}{FIRST_ROW}:{TABLE['check']}{last}"
    out.append({"address": result("cases"),
                "formula": (f'="The case table below: "&COUNTIF({checks},"As documented")&" of {len(cases)} calls as '
                            f'documented, "&COUNTIF({checks},"UNEXPECTED")&" unexpected."')})
    return out


def try_formulas(data: dict[str, Any]) -> list[dict[str, str]]:
    functions = data["functions"]
    args = tuple(absolute(cell) for cell in ARG_CELLS)
    low, high = (absolute(cell) for cell in SWEEP_CELLS)
    points = data["sweep_points"]
    shown = [shown_argument(cell) for cell in args]
    # The arguments are joined by the list separator a user of this Excel would type: a semicolon where the
    # decimal separator is a comma, so that 0,95 reads as one number.
    values = "CHOOSE(DiagArity," + ",".join('&sep&'.join(shown[:count]) for count in range(1, 5)) + ")"
    refs = 'CHOOSE(DiagArity,"D7","D7,D8","D7,D8,D9","D7,D8,D9,D10")'
    row = str(FIRST_ROW)
    sweep_args = ("xv",) + args[1:]
    returned = result("returned")
    return [
        {"address": INDEX_CELL, "formula": f"=MATCH({absolute(FUNCTION_CELL)},DiagFunctionList,0)"},
        {"address": ARITY_CELL, "formula": "=IFERROR(INDEX(DiagArityList,DiagIndex),0)"},
        *({"address": f"C{7 + slot - 1}", "formula": label_formula(slot)} for slot in range(1, 5)),
        {"address": result("call"),
         "formula": (f'=IF({ok()},LET(sep,IF(MID(1/2&"",2,1)=",",";",","),'
                     f'"="&{absolute(FUNCTION_CELL)}&"("&{values}&")"),"")')},
        {"address": returned, "formula": f"=IF({ok()},{dispatch(functions, args)},NA())"},
        {"address": f"F{RESULT_ROWS['returned']}", "formula": f'=IF({ok()},"="&{absolute(FUNCTION_CELL)}&"("&{refs}&")","")'},
        {"address": result("kind"), "formula": f'=IF({ok()},{kind_of(returned)},"")'},
        {"address": result("status"),
         "formula": (f'=IF({ok()},LET(s,DiagnosticsStatus({absolute(FUNCTION_CELL)},{",".join(args)}),'
                     f'IF(s="","(empty: the library wrote no failure)",s)),"")')},
        {"address": result("meaning"), "formula": by_kind(MEANING)},
        {"address": result("action"), "formula": by_kind(ACTION)},
        {"address": result("domain"), "formula": f'=IF({ok()},INDEX(DiagDomains,DiagIndex),"")'},
        {"address": result("converted"), "formula": converted_formula()},
        {"address": result("sweep"), "formula": sweep_formula(points)},
        {"address": f"{HELPERS['x']}{row}",
         "formula": f"=IF({ok()},{low}+({high}-{low})*SEQUENCE({points},1,0,1)/{points - 1},NA())"},
        # The LAMBDA never returns an error value: an error of the library comes back as its name, a text, as
        # every LAMBDA of the other sheets catches its errors. The chart's helpers then tell numbers from names.
        {"address": f"{HELPERS['raw']}{row}",
         "formula": (f"=IF({ok()},MAP(DiagSweepX,LAMBDA(xv,LET(v,{dispatch(functions, sweep_args)},"
                     f"IF(ISNUMBER(v),v,{error_name('v')})))),NA())")},
        {"address": f"{HELPERS['value']}{row}", "formula": "=IF(ISNUMBER(DiagSweepRaw),DiagSweepRaw,NA())"},
    ]


def catalogue_texts(functions: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """The catalogue in hidden columns: one row per function."""
    out = []
    for offset, item in enumerate(functions):
        row = CATALOGUE_FIRST + offset
        out.append({"address": f"{CATALOGUE['name']}{row}", "value": item["name"]})
        out.append({"address": f"{CATALOGUE['arity']}{row}", "value": len(item["params"])})
        for slot, name in enumerate(item["params"]):
            out.append({"address": f"{CATALOGUE_PARAMS[slot]}{row}", "value": name})
            out.append({"address": f"{CATALOGUE_EXAMPLES[slot]}{row}", "value": item["example"][slot]})
        out.append({"address": f"{CATALOGUE['sweep_from']}{row}", "value": item["sweep"][0]})
        out.append({"address": f"{CATALOGUE['sweep_to']}{row}", "value": item["sweep"][1]})
        out.append({"address": f"{CATALOGUE['domain']}{row}", "value": item["domain"]})
    return out


def defaults(data: dict[str, Any]) -> dict[str, Any]:
    item = next(entry for entry in data["functions"] if entry["name"] == data["default_function"])
    example = list(item["example"]) + [None] * (len(ARG_CELLS) - len(item["example"]))
    return {"function": item["name"], "args": example, "sweep": list(item["sweep"])}


def static_messages() -> dict[str, list[str]]:
    any_value = ("Any value: a number, a text, an empty cell. The sheet calls the function with it and shows what "
                 "comes back. The label on the left names the parameter.")
    bound = f"{SWEEP_LIMIT:.0E}"
    return {
        FUNCTION_CELL: ["Function", "A public K_STATS_ function. Choosing one loads its example arguments and sweep.",
                        "Not in the list", "Choose a function from the list."],
        **{cell: [f"Argument {slot}", any_value, "", ""] for slot, cell in enumerate(ARG_CELLS, 1)},
        SWEEP_CELLS[0]: ["Sweep from", f"First value of argument 1 on the chart, a number at most {bound} in size.",
                         "Invalid sweep", f"Enter a number at most {bound} in size."],
        SWEEP_CELLS[1]: ["Sweep to", f"Last value of argument 1 on the chart, above Sweep from and at most {bound} "
                         "in size.", "Invalid sweep", f"Enter a number at most {bound} in size."],
    }


def plan() -> dict[str, Any]:
    data = load()
    functions, cases = data["functions"], data["cases"]
    last = FIRST_ROW + len(cases) - 1
    catalogue_last = CATALOGUE_FIRST + len(functions) - 1

    def column_range(column: str) -> str:
        return f"={SHEET}!${column}${CATALOGUE_FIRST}:${column}${catalogue_last}"

    names = [
        {"name": "DiagFunctionList", "refers": column_range(CATALOGUE["name"])},
        {"name": "DiagArityList", "refers": column_range(CATALOGUE["arity"])},
        {"name": "DiagParams", "refers": (f"={SHEET}!${CATALOGUE_PARAMS[0]}${CATALOGUE_FIRST}:"
                                          f"${CATALOGUE_PARAMS[-1]}${catalogue_last}")},
        {"name": "DiagDomains", "refers": column_range(CATALOGUE["domain"])},
        {"name": "DiagEmpty", "refers": "=" + ref(EMPTY_CELL)},
        {"name": "DiagIndex", "refers": "=" + ref(INDEX_CELL)},
        {"name": "DiagArity", "refers": "=" + ref(ARITY_CELL)},
        *check_names(),
        {"name": "DiagStatusBad", "refers": f'={ref(STATUS.replace("$", ""))}<>"OK"'},
    ]
    spills = {"DiagSweepX": HELPERS["x"], "DiagSweepRaw": HELPERS["raw"], "DiagSweepValue": HELPERS["value"]}
    names += [{"name": name, "refers": f"={SHEET}!${letter}${FIRST_ROW}#"} for name, letter in spills.items()]
    formulas = [
        {"address": STATUS.replace("$", ""), "formula": status_cell_formula()},
        {"address": MACROS,
         "formula": '=IF(ISERROR(K_STATS_NormalStandard_Cumulative(0)+0*COUNTA($D$6:$D$12)),"DISABLED","ENABLED")'},
        {"address": "F4", "formula": f'=IF(${MACROS[0]}${MACROS[1:]}="ENABLED","","MACROS ARE DISABLED: the results are not live")'},
        *try_formulas(data),
        *case_formulas(cases),
    ]
    environment = data["environment"]
    provenance = (f"Every value is a live call to the library. The expected outcomes come from the documentation "
                  f"(manual chapters 6, 7 and 9, docs/PUBLIC_API.md) and the numbers from mpmath "
                  f"{environment['mpmath']} (Python {environment['python']}), stored by "
                  "examples/demo/diagnostics_references.py; ci/Test-Diagnostics.ps1 checks them in Excel.")
    labels = {
        "call": "The call with its values", "returned": "Returned", "kind": "Kind of result",
        "status": "Status (written by the library)", "meaning": "What it means", "action": "What to do",
        "domain": "Domain of the function", "converted": "Arguments Excel converted", "sweep": "Sweep of argument 1",
        "cases": "Case table", "not_shown": "Not shown", "excel_build": "Computed by", "provenance": "Provenance",
    }
    texts = [
        {"address": "B1", "value": "Errors and Diagnostics"},
        {"address": "B2", "value": "What each function returns on valid input, at the edge of a support and on input "
                                   "it refuses, and how to read the reason."},
        {"address": "C3", "value": "Choose a function and type any arguments in the yellow cells, invalid ones "
                                   "included. Every result is a call to the embedded VBA API (K_STATS_ functions). "
                                   "Macros are required."},
        {"address": "C6", "value": "Function"},
        {"address": "C11", "value": "Sweep argument 1 from"},
        {"address": "C12", "value": "Sweep argument 1 to"},
        {"address": "C19", "value": "Input status"},
        {"address": "C20", "value": "Macros"},
        *({"address": f"C{RESULT_ROWS[key]}", "value": label} for key, label in labels.items()),
        {"address": result("not_shown"), "value": NOT_SHOWN},
        {"address": result("provenance"), "value": provenance},
        {"address": "F5", "value": "Read the results"},
        {"address": "F6", "value": "A number means success. #NUM! means the library declined the request; the Status "
                                   "names the reason. #VALUE! means Excel could not pass an argument."},
        {"address": "F8", "value": "A worksheet cannot read Status: the sheet reads it through DiagnosticsStatus, a "
                                   "macro of this workbook that makes the same call from VBA."},
        {"address": "F10", "value": "Outside a continuous support the result is 0 or 1, not an error; a discrete "
                                    "count outside its support is #NUM!."},
        {"address": "F12", "value": "The case table checks each call against the outcome the documentation gives it: "
                                    "As documented, or UNEXPECTED."},
        {"address": "F14", "value": "Educational demo, not an accuracy certification. The library's guarantees are "
                                    "in docs/."},
        {"address": "F16", "value": "Needs macros and Excel 365 or 2024. If Macros says DISABLED, unblock the file and "
                                    "reopen."},
        {"address": f"C{HEADER_ROW - 1}", "value": "The case table: what the library returns, and why"},
        {"address": f"{TABLE['title']}{HEADER_ROW}", "value": "Case"},
        {"address": f"{TABLE['call']}{HEADER_ROW}", "value": "The call as written"},
        {"address": f"{TABLE['returned']}{HEADER_ROW}", "value": "Returned"},
        {"address": f"{TABLE['expected']}{HEADER_ROW}", "value": "Documented outcome"},
        {"address": f"{TABLE['check']}{HEADER_ROW}", "value": "Check"},
        {"address": f"{TABLE['status']}{HEADER_ROW}", "value": "Status (from the library)"},
        {"address": f"{TABLE['cause']}{HEADER_ROW}", "value": "Cause and correction"},
        *({"address": f"{TABLE['title']}{FIRST_ROW + offset}", "value": f"{GROUPS[case['group']]}: {case['title']}"}
          for offset, case in enumerate(cases)),
        *({"address": f"{TABLE['cause']}{FIRST_ROW + offset}", "value": f"{case['cause']} {case['action']}"}
          for offset, case in enumerate(cases)),
        {"address": f"{HELPERS['x']}{HEADER_ROW}", "value": "Argument 1"},
        {"address": f"{HELPERS['raw']}{HEADER_ROW}", "value": "Returned"},
        {"address": f"{HELPERS['value']}{HEADER_ROW}", "value": "Number"},
        {"address": f"{CATALOGUE['name']}{CATALOGUE_FIRST - 1}", "value": "Function"},
        *catalogue_texts(functions),
    ]
    return {
        "schema": "diagnostics-plan-1",
        "sheet": SHEET,
        "names": names,
        "formulas": formulas,
        "texts": texts,
        "defaults": defaults(data),
        "messages": static_messages(),
        "layout": {"inputs": "D6:D12", "function": FUNCTION_CELL, "arguments": list(ARG_CELLS),
                   "sweep": list(SWEEP_CELLS), "status": STATUS.replace("$", ""), "macros": MACROS,
                   "result_rows": RESULT_ROWS, "header_row": HEADER_ROW, "first_row": FIRST_ROW, "last_row": last,
                   "table": TABLE, "table_merges": list(TABLE_MERGES), "formula_merge": "F{row}:I{row}",
                   "text_merge": "D{row}:I{row}", "hidden_from": "Q", "spacers": ["C13:D16"],
                   "outputs": ["D19:D20", "D21:I33", f"C{HEADER_ROW}:O{last}"],
                   "catalogue": f"{CATALOGUE['name']}{CATALOGUE_FIRST}:{CATALOGUE['domain']}{catalogue_last}",
                   "empty_cell": EMPTY_CELL, "chart_area": "K5:O35", "sweep_points": data["sweep_points"]},
        "validation": {"function_list": "=DiagFunctionList",
                       "custom": {SWEEP_CELLS[0]: "=DiagSweepFromOk", SWEEP_CELLS[1]: "=DiagSweepToOk"},
                       "input_only": list(ARG_CELLS)},
        "conditional": {"status_bad": "=DiagStatusBad",
                        "check_rules": [{"text": "As documented", "font": 0x006100, "fill": 0xCEEFC6},
                                        {"text": "UNEXPECTED", "font": 0x0000C0, "fill": 0xCEC7FF}]},
        "charts": [
            # One series, a line: an error leaves a gap. A second series of markers at each error, drawn over the
            # line in another chart type, made Excel stop (access violation) in other sheets' chart macros.
            {"name": "DiagSweep", "title": "The function across the sweep of argument 1; a gap at each error",
             "series": [{"name": "Returned value", "x": "DiagSweepX", "y": "DiagSweepValue"}]},
        ],
    }


def formula_lengths(result_plan: dict[str, Any]) -> list[tuple[str, int]]:
    """Length of every formula and name, to check them against Excel's limit of 8192 characters."""
    sizes = [(item["address"], len(item["formula"])) for item in result_plan["formulas"]]
    sizes += [(item["name"], len(item["refers"])) for item in result_plan["names"]]
    return sorted(sizes, key=lambda item: -item[1])


if __name__ == "__main__":
    for name, size in formula_lengths(plan())[:6]:
        print(size, name)
