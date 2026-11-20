"""The Tails sheet (Extreme Tails and Numerical Limits) as data: every formula, name, validation rule and text.

Same method as comparison_plan.py: examples/demo/tails-cases.json (made by examples/demo/tails_references.py)
is the one source of the scenarios, and this module writes from it the plan that
tools/demobuild/Evolve-Tails.ps1 applies in desktop Excel.

A scenario is a family at fixed parameters and a ladder of points from the body of the distribution to
far beyond the smallest positive number an Excel cell holds (2.2250738585072014E-308). The sheet shows,
for the scenario chosen in D6, what two things do to a small probability:

  * cancellation: 1 - CDF becomes exactly 0, or loses its digits, while the tail is still representable;
    the direct survival function keeps them;
  * underflow: the true tail is below the range of a Double, so every method returns 0. For the discrete
    families the library's LogPMF goes past that wall; LN(PMF) does not.

Layout, in the style of the Explorer:

    C6:D8     the inputs: the scenario and the two grade thresholds (correct digits, the study's 12 and 6)
    D19:D20   the input status and the macros
    C21:I35   the scenario: parameters, the two formulas as written, the reference and its provenance,
              the release guarantee, where the alternative stops being usable, where the direct path does,
              the representation limit, what the statuses mean, the Excel that computed the sheet
    K5:O35    two charts: the correct digits against the depth of the tail, and log10 of the values
              against it, with the limit of the Double range drawn
    C37:M..   the ladder: point, reference, the two values, their correct digits and statuses, the depth
    R:W       hidden: every point of every scenario (scenario, x, reference text, reference value, depth,
              the reference with ten digits, which is what the ladder shows)
    X:AE      hidden: the scenarios (name, kind, title, parameters, the two formulas as text, guarantee, method)
    AG:AJ     hidden: spilled helpers for the second chart; AL:AO the two lines of the Double limit

The reference is the mpmath value at the point, never 1 - CDF. A reference below 2.2250738585072014E-308
is stored as text and its value cell holds 0: the cell cannot hold it, and the sheet says so ("zero:
underflow") instead of grading. Correct digits are -log10 of the relative difference, from 0 to 16.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
CASES_FILE = ROOT / "examples" / "demo" / "tails-cases.json"

SHEET = "Tails"
SCENARIO, REFERENCE_GRADE, REPORT_GRADE = "$D$6", "$D$7", "$D$8"
STATUS = "$D$19"
MACROS = "D20"
HEADER_ROW = 37
FIRST_ROW = 38
FULL_DIGITS = 16
SMALLEST_NORMAL = 2.2250738585072014e-308
DETAIL_ROWS = {
    "title": 21, "parameters": 22, "direct": 23, "alternative": 24, "reference": 25, "guarantee": 26,
    "alternative_usable": 27, "direct_usable": 28, "limit": 29, "count": 30, "statuses": 31,
    "excel_build": 34, "provenance": 35,
}
STATUSES_MERGE = "D31:I33"
# Hidden storage: one row per point, one row per scenario, spilled helpers, the two lines of the limit.
STORE = {"scenario": "R", "x": "S", "reference_text": "T", "reference_value": "U", "depth": "V", "reference_short": "W"}
SCENARIOS = {"name": "X", "kind": "Y", "title": "Z", "parameters": "AA", "direct": "AB", "alternative": "AC",
             "contract": "AD", "method": "AE"}
HELPERS = {"reference_value": "AG", "direct_log": "AH", "alternative_log": "AI", "reference_log": "AJ"}
LIMIT = {"wall_x": "AL", "wall_y": "AM", "line_x": "AN", "line_y": "AO"}
STATUS_WORDS = {
    "reference": "reference grade", "report": "report grade", "inadequate": "inadequate", "wrong": "wrong",
    "cancellation": "zero: cancellation", "underflow": "zero: underflow", "mass": "PMF underflow",
    "unexpected": "zero: unexpected", "below": "nonzero below range", "error": "error",
}
# BGR, as Excel's Color property reads them: red for a wrong value or a cancellation, amber for an underflow.
STATUS_COLORS = [
    ("wrong", 0x0000C0, 0xCEC7FF), ("zero: cancellation", 0x0000C0, 0xCEC7FF), ("zero: unexpected", 0x0000C0, 0xCEC7FF),
    ("error", 0x0000C0, 0xCEC7FF),
    ("zero: underflow", 0x00579C, 0x9CEBFF), ("PMF underflow", 0x00579C, 0x9CEBFF),
    ("nonzero below range", 0x00579C, 0x9CEBFF),
]


def load() -> dict[str, Any]:
    data: dict[str, Any] = json.loads(CASES_FILE.read_text(encoding="utf-8"))
    names = [scenario["name"] for scenario in data["scenarios"]]
    if len(names) != len(set(names)):
        raise ValueError("duplicate scenario")
    if data["default_scenario"] not in names:
        raise ValueError("the default scenario is not a scenario")
    return data


def ref(address: str) -> str:
    return f"{SHEET}!{address}"


def counts(data: dict[str, Any]) -> tuple[int, int]:
    """The rows of the ladder (the longest scenario) and the rows of the hidden storage."""
    return (max(len(item["points"]) for item in data["scenarios"]),
            sum(len(item["points"]) for item in data["scenarios"]))


def last_row(data: dict[str, Any]) -> int:
    return FIRST_ROW + counts(data)[0] - 1


def store_last_row(data: dict[str, Any]) -> int:
    return FIRST_ROW + counts(data)[1] - 1


def call(template: str) -> str:
    return template.format(x="x")


def column(letter: str, data: dict[str, Any], rows: int) -> str:
    return f"${letter}${FIRST_ROW}:${letter}${FIRST_ROW + rows - 1}"


def digits_formula(value: str, truth: str) -> str:
    """Correct digits of a value against the reference, 0 to 16; #N/A where there is nothing to grade
    (not a number, or a reference that no Double can hold)."""
    # The relative difference is computed, never a comparison with =: Excel treats two numbers equal to
    # 15 digits as equal, which would show 16 digits for a value that has 14.6.
    return (f"IF(OR(NOT(ISNUMBER({value})),{truth}=0),NA(),LET(rel,ABS({value}-{truth})/ABS({truth}),"
            f"IF(rel=0,{FULL_DIGITS},MIN({FULL_DIGITS},MAX(0,-LOG10(rel))))))")


def status_formula(alternative: bool) -> str:
    """The status of a value: a grade, or what happened to it when it is not gradeable."""
    zero = STATUS_WORDS["cancellation"] if alternative else STATUS_WORDS["unexpected"]
    return (f'LET(digs,IFERROR(got,-1),IF(ISTEXT(val),IF(val="underflow","{STATUS_WORDS["mass"]}","{STATUS_WORDS["error"]}"),'
            f'IF(truth=0,IF(val=0,"{STATUS_WORDS["underflow"]}","{STATUS_WORDS["below"]}"),'
            f'IF(val=0,"{zero}",IF(digs<=0,"{STATUS_WORDS["wrong"]}",'
            f'IF(digs>={REFERENCE_GRADE},"{STATUS_WORDS["reference"]}",'
            f'IF(digs>={REPORT_GRADE},"{STATUS_WORDS["report"]}","{STATUS_WORDS["inadequate"]}")))))))')


def switch(data: dict[str, Any], key: str, failure: dict[str, str]) -> str:
    """The selected scenario's call, as one SWITCH over the scenario cell."""
    branches = ",".join(
        f'"{item["name"]}",IFERROR({call(item[key])},"{failure[item["kind"]]}")' for item in data["scenarios"])
    return f"SWITCH({SCENARIO},{branches},NA())"


def spill(expression: str) -> str:
    return f'=IF({STATUS}="OK",{expression},NA())'


def ladder_formulas(data: dict[str, Any]) -> list[dict[str, str]]:
    scenario = f"TailsStoreScenario={SCENARIO}"
    direct = switch(data, "direct", {"tail": "error", "log": "error"})
    alternative = switch(data, "alternative", {"tail": "error", "log": "underflow"})
    row = str(FIRST_ROW)

    def pick(name: str) -> str:
        return f"FILTER({name},{scenario})"

    return [
        {"address": f"C{row}", "formula": spill(pick("TailsStoreX"))},
        {"address": f"D{row}", "formula": spill(pick("TailsStoreReferenceShort"))},
        {"address": f"F{row}", "formula": spill(f"MAP(TailsX,LAMBDA(x,{direct}))")},
        {"address": f"G{row}", "formula": spill(f"MAP(TailsX,LAMBDA(x,{alternative}))")},
        {"address": f"H{row}", "formula": spill(
            f"MAP(TailsDirect,TailsReferenceValue,LAMBDA(val,truth,{digits_formula('val', 'truth')}))")},
        {"address": f"I{row}", "formula": spill(
            f"MAP(TailsAlternative,TailsReferenceValue,LAMBDA(val,truth,{digits_formula('val', 'truth')}))")},
        {"address": f"K{row}", "formula": spill(
            f"MAP(TailsDirect,TailsReferenceValue,TailsDirectDigits,LAMBDA(val,truth,got,{status_formula(False)}))")},
        {"address": f"L{row}", "formula": spill(
            f"MAP(TailsAlternative,TailsReferenceValue,TailsAlternativeDigits,LAMBDA(val,truth,got,{status_formula(True)}))")},
        {"address": f"M{row}", "formula": spill(pick("TailsStoreDepth"))},
        {"address": f"{HELPERS['reference_value']}{row}", "formula": spill(pick("TailsStoreReferenceValue"))},
        {"address": f"{HELPERS['direct_log']}{row}", "formula": spill(
            'MAP(TailsDirect,LAMBDA(val,IF(ISNUMBER(val),IF(TailsKind="log",val/LN(10),IF(val>0,LOG10(val),NA())),NA())))')},
        {"address": f"{HELPERS['alternative_log']}{row}", "formula": spill(
            'MAP(TailsAlternative,LAMBDA(val,IF(ISNUMBER(val),IF(TailsKind="log",val/LN(10),IF(val>0,LOG10(val),NA())),NA())))')},
        {"address": f"{HELPERS['reference_log']}{row}", "formula": spill("-TailsDepth")},
    ]


def detail_formulas(data: dict[str, Any]) -> list[dict[str, str]]:
    def pick(name: str) -> str:
        return f"INDEX({name},TailsRow)"

    environment = data["environment"]
    reference = (f'mpmath {environment["mpmath"]} at 60 digits, stored with {data["reference"]["digits"]}: "'
                 f'&{pick("TailsScenarioMethod")}&". A tail is never referenced by 1 - CDF."')
    guarantee = ('"Release guarantee (docs/user-manual, section 10.3): "&' + pick("TailsScenarioContract")
                 + '&" The far-tail points of this ladder are a demonstration: they are measured on this run '
                   'and do not extend it."')
    usable = ('LET(i,XMATCH(TRUE,IFERROR(TailsAlternativeDigits,0)<' + REPORT_GRADE + '),"from the point "&INDEX(TailsX,i)'
              '&" (depth "&ROUND(INDEX(TailsDepth,i),1)&", about 1E-"&ROUND(INDEX(TailsDepth,i),0)&"): "&INDEX(TailsAlternativeStatus,i))')
    deepest = ('LET(ok,IFERROR(TailsDirectDigits,0)>=' + REFERENCE_GRADE + ',i,XMATCH(TRUE,ok,0,-1),'
               '"the deepest point at reference grade is "&INDEX(TailsX,i)&" (depth "&ROUND(INDEX(TailsDepth,i),1)'
               '&"), then: "&IFERROR(INDEX(TailsDirectStatus,i+1),"no further point"))')
    limit = ('IF(TailsKind="log",LET(i,XMATCH("' + STATUS_WORDS["mass"] + '",TailsAlternativeStatus),"from the point "&INDEX(TailsX,i)'
             '&": the mass, about 1E-"&ROUND(INDEX(TailsDepth,i),0)&", is below 2.2E-308: LN(PMF) fails and LogPMF stays finite ("'
             '&ROUND(INDEX(TailsDirect,i),1)&")."),'
             'LET(i,XMATCH(0,TailsReferenceValue),"from the point "&INDEX(TailsX,i)&": the reference, about 1E-"'
             '&ROUND(INDEX(TailsDepth,i),0)&", is below 2.2E-308, the smallest positive number an Excel cell holds. '
             'Direct path: "&INDEX(TailsDirectStatus,i)&"; alternative: "&INDEX(TailsAlternativeStatus,i)&"."))')
    count = ('ROWS(TailsX)&" points: "&SUM(--(IFERROR(TailsDirectDigits,0)>=' + REFERENCE_GRADE + '))&" at reference grade for the direct '
             'path and "&SUM(--(IFERROR(TailsAlternativeDigits,0)>=' + REFERENCE_GRADE + '))&" for the alternative path."')
    at = {key: f"D{number}" for key, number in DETAIL_ROWS.items()}
    return [
        {"address": at["title"], "formula": f'=IFERROR({pick("TailsScenarioTitle")},"")'},
        {"address": at["parameters"], "formula": f'=IFERROR({pick("TailsScenarioParameters")},"")'},
        {"address": at["direct"], "formula": f'=IFERROR({pick("TailsScenarioDirect")},"")'},
        {"address": at["alternative"], "formula": f'=IFERROR({pick("TailsScenarioAlternative")},"")'},
        {"address": at["reference"], "formula": f'=IFERROR("{reference},"")'},
        {"address": at["guarantee"], "formula": f'=IFERROR({guarantee},"")'},
        {"address": at["alternative_usable"],
         "formula": f'=IFERROR({usable},"never below the report grade on this ladder")'},
        {"address": at["direct_usable"], "formula": f'=IFERROR({deepest},"nowhere on this ladder")'},
        {"address": at["limit"], "formula": f'=IFERROR({limit},"no point of this ladder is below 2.2E-308")'},
        {"address": at["count"], "formula": f'=IFERROR({count},"")'},
    ]


def status_cell_formula() -> str:
    return (f'=IF(ISNA(TailsRow),"Check: scenario",'
            f'IF(NOT(AND(ISNUMBER({REFERENCE_GRADE}),ISNUMBER({REPORT_GRADE}))),"Check: grade thresholds",'
            f'IF(OR({REPORT_GRADE}<1,{REFERENCE_GRADE}>{FULL_DIGITS},{REPORT_GRADE}>{REFERENCE_GRADE}),'
            f'"Check: grade thresholds","OK")))')


def static_messages(data: dict[str, Any]) -> dict[str, list[str]]:
    count = len(data["scenarios"])
    return {
        "D6": ["Scenario", f"Choose one of the {count} scenarios: a family at fixed parameters and a ladder of points from the "
               "body of the distribution to beyond the range of a Double. The log scenarios compare LogPMF with LN(PMF).",
               "Unknown scenario", "Choose a scenario from the list."],
        "D7": ["Reference grade from", "Correct significant digits from which a value is reference grade "
               f"(the study uses 12). From 1 to {FULL_DIGITS}, not below the report threshold.",
               "Invalid threshold", f"Enter a number of digits from 1 to {FULL_DIGITS}, not below D8."],
        "D8": ["Report grade from", "Correct significant digits from which a value is fit to report "
               f"(the study uses 6). From 1 to {FULL_DIGITS}, not above the reference threshold.",
               "Invalid threshold", f"Enter a number of digits from 1 to {FULL_DIGITS}, not above D7."],
    }


def columns_to_write(data: dict[str, Any]) -> list[dict[str, Any]]:
    """The hidden storage, one column at a time: the points of every scenario and the table of scenarios."""
    points = [(item, point) for item in data["scenarios"] for point in item["points"]]
    end = store_last_row(data)
    columns = [
        {"address": f"{STORE['scenario']}{FIRST_ROW}:{STORE['scenario']}{end}", "kind": "text",
         "values": [item["name"] for item, _ in points]},
        {"address": f"{STORE['x']}{FIRST_ROW}:{STORE['x']}{end}", "kind": "number",
         "values": [float(point["x"]) for _, point in points]},
        {"address": f"{STORE['reference_text']}{FIRST_ROW}:{STORE['reference_text']}{end}", "kind": "text",
         "values": [point["reference"] for _, point in points]},
        {"address": f"{STORE['reference_value']}{FIRST_ROW}:{STORE['reference_value']}{end}", "kind": "number",
         "values": [float(point["reference_value"]) for _, point in points]},
        {"address": f"{STORE['depth']}{FIRST_ROW}:{STORE['depth']}{end}", "kind": "number",
         "values": [float(point["depth"]) for _, point in points]},
        {"address": f"{STORE['reference_short']}{FIRST_ROW}:{STORE['reference_short']}{end}", "kind": "text",
         "values": [point["reference_short"] for _, point in points]},
    ]
    scenarios = data["scenarios"]
    finish = FIRST_ROW + len(scenarios) - 1
    shown = {
        "name": [item["name"] for item in scenarios], "kind": [item["kind"] for item in scenarios],
        "title": [item["title"] for item in scenarios], "parameters": [item["parameters"] for item in scenarios],
        "direct": ["=" + call(item["direct"]) for item in scenarios],
        "alternative": ["=" + call(item["alternative"]) for item in scenarios],
        "contract": [item["contract"] for item in scenarios], "method": [item["method"] for item in scenarios],
    }
    for key, letter in SCENARIOS.items():
        columns.append({"address": f"{letter}{FIRST_ROW}:{letter}{finish}", "kind": "text", "values": shown[key]})
    return columns


def plan() -> dict[str, Any]:
    data = load()
    rows, store_rows = counts(data)
    end = last_row(data)
    scenario_count = len(data["scenarios"])
    names = [
        {"name": "TailsStoreScenario", "refers": "=" + ref(column(STORE["scenario"], data, store_rows))},
        {"name": "TailsStoreX", "refers": "=" + ref(column(STORE["x"], data, store_rows))},
        {"name": "TailsStoreReferenceText", "refers": "=" + ref(column(STORE["reference_text"], data, store_rows))},
        {"name": "TailsStoreReferenceValue", "refers": "=" + ref(column(STORE["reference_value"], data, store_rows))},
        {"name": "TailsStoreDepth", "refers": "=" + ref(column(STORE["depth"], data, store_rows))},
        {"name": "TailsStoreReferenceShort", "refers": "=" + ref(column(STORE["reference_short"], data, store_rows))},
        {"name": "TailsScenarioNames", "refers": "=" + ref(column(SCENARIOS["name"], data, scenario_count))},
        {"name": "TailsScenarioKind", "refers": "=" + ref(column(SCENARIOS["kind"], data, scenario_count))},
        {"name": "TailsScenarioTitle", "refers": "=" + ref(column(SCENARIOS["title"], data, scenario_count))},
        {"name": "TailsScenarioParameters", "refers": "=" + ref(column(SCENARIOS["parameters"], data, scenario_count))},
        {"name": "TailsScenarioDirect", "refers": "=" + ref(column(SCENARIOS["direct"], data, scenario_count))},
        {"name": "TailsScenarioAlternative", "refers": "=" + ref(column(SCENARIOS["alternative"], data, scenario_count))},
        {"name": "TailsScenarioContract", "refers": "=" + ref(column(SCENARIOS["contract"], data, scenario_count))},
        {"name": "TailsScenarioMethod", "refers": "=" + ref(column(SCENARIOS["method"], data, scenario_count))},
        {"name": "TailsRow", "refers": f"=IFERROR(MATCH({ref(SCENARIO)},TailsScenarioNames,0),NA())"},
        {"name": "TailsKind", "refers": "=INDEX(TailsScenarioKind,TailsRow)"},
        {"name": "TailsStatusBad", "refers": f'={ref(STATUS)}<>"OK"'},
    ]
    spills = {"TailsX": "C", "TailsReference": "D", "TailsDirect": "F", "TailsAlternative": "G",
              "TailsDirectDigits": "H", "TailsAlternativeDigits": "I", "TailsDirectStatus": "K",
              "TailsAlternativeStatus": "L", "TailsDepth": "M", "TailsReferenceValue": HELPERS["reference_value"],
              "TailsDirectLog": HELPERS["direct_log"], "TailsAlternativeLog": HELPERS["alternative_log"],
              "TailsReferenceLog": HELPERS["reference_log"]}
    names += [{"name": name, "refers": "=" + ref(f"${letter}${FIRST_ROW}#")} for name, letter in spills.items()]
    names += [
        {"name": "TailsWallX", "refers": "=" + ref(f"${LIMIT['wall_x']}${FIRST_ROW}:${LIMIT['wall_x']}${FIRST_ROW + 1}")},
        {"name": "TailsWallY", "refers": "=" + ref(f"${LIMIT['wall_y']}${FIRST_ROW}:${LIMIT['wall_y']}${FIRST_ROW + 1}")},
        {"name": "TailsLineX", "refers": "=" + ref(f"${LIMIT['line_x']}${FIRST_ROW}:${LIMIT['line_x']}${FIRST_ROW + 1}")},
        {"name": "TailsLineY", "refers": "=" + ref(f"${LIMIT['line_y']}${FIRST_ROW}:${LIMIT['line_y']}${FIRST_ROW + 1}")},
    ]
    depth_of_limit = float(-__import__("math").log10(SMALLEST_NORMAL))
    formulas = [
        {"address": STATUS.replace("$", ""), "formula": status_cell_formula()},
        {"address": MACROS, "formula": '=IF(ISERROR(K_STATS_NormalStandard_Cumulative(0)+0*COUNTA($D$6:$D$8)),"DISABLED","ENABLED")'},
        {"address": "F4", "formula": f'=IF(${MACROS[0]}${MACROS[1:]}="ENABLED","","MACROS ARE DISABLED: the results are not live")'},
        *detail_formulas(data),
        *ladder_formulas(data),
        {"address": f"{LIMIT['line_x']}{FIRST_ROW + 1}", "formula": "=IFERROR(MAX(TailsDepth),0)"},
    ]
    environment = data["environment"]
    provenance = (f'Reference: {data["reference"]["method"]} (mpmath {environment["mpmath"]}, Python {environment["python"]}), '
                  "stored by examples/demo/tails_references.py. The library and Excel values are computed live.")
    texts = [
        {"address": "B1", "value": "Extreme Tails and Numerical Limits"},
        {"address": "B2", "value": "Direct survival against 1 - CDF, tiny probabilities, log-mass and the limit of the Double range."},
        {"address": "C3", "value": "Choose a scenario to read the library's direct result and the alternative, point by point, against a "
                                   "high-precision reference. Every value is a call to the embedded VBA API (K_STATS_ functions). Macros are required."},
        {"address": "C6", "value": "Scenario"},
        {"address": "C7", "value": "Reference grade from (digits)"},
        {"address": "C8", "value": "Report grade from (digits)"},
        {"address": "C19", "value": "Input status"},
        {"address": "C20", "value": "Macros"},
        {"address": "C21", "value": "Scenario shown"},
        {"address": "C22", "value": "Parameters and units"},
        {"address": "C23", "value": "Direct path"},
        {"address": "C24", "value": "Alternative path"},
        {"address": "C25", "value": "Reference"},
        {"address": "C26", "value": "Library guarantee"},
        {"address": "C27", "value": "Alternative stops being usable"},
        {"address": "C28", "value": "Direct path, reference grade"},
        {"address": "C29", "value": "Representation limit"},
        {"address": "C30", "value": "Points at reference grade"},
        {"address": "C31", "value": "What the statuses mean"},
        {"address": "D31", "value": "Zero: cancellation - 1 - CDF is exactly 0 although the true tail can be stored (about 1E-16 and below). "
                                    "Zero: underflow - the true value is below 2.2E-308, so every method returns 0. PMF underflow - "
                                    "LN of a mass that underflowed to 0; LogPMF is finite. Wrong, inadequate, report and reference grade "
                                    "follow the thresholds in D7 and D8."},
        {"address": "C34", "value": "Computed by"},
        {"address": "C35", "value": "Provenance"},
        {"address": "D35", "value": provenance},
        {"address": "F5", "value": "Read the results"},
        {"address": "F6", "value": "Yellow cells are editable. Choose a scenario: the table lists its points from the body to beyond "
                                   "the Double range, with the depth of each, -log10 of the probability."},
        {"address": "F8", "value": "Cancellation is not underflow. 1 - CDF dies near 1E-16 while the tail is still a normal number; the "
                                   "direct path goes on until the true value is below 2.2E-308, where nothing can be stored."},
        {"address": "F10", "value": "Each value is graded against the mpmath reference only; a tail is never referenced by 1 - CDF. "
                                    f"Correct digits are -log10 of the relative difference, at most {FULL_DIGITS}."},
        {"address": "F12", "value": "Only the discrete families have a log function (LogPMF). There is no log-survival or log-CDF: "
                                    "a continuous tail below 2.2E-308 is not available from the library."},
        {"address": "F14", "value": "A point far in the tail is a demonstration, not a certified guarantee: the library's contracts are in docs/. "
                                    "Where the sheet shows an error or a zero, that is what this Excel returned."},
        {"address": "F16", "value": "Needs macros and Excel 365 or 2024. If Macros says DISABLED, unblock the file and reopen."},
        {"address": f"C{HEADER_ROW - 1}", "value": "The ladder against the reference"},
        {"address": f"C{HEADER_ROW}", "value": "Point (x, z, t or k)"},
        {"address": f"D{HEADER_ROW}", "value": "Reference (mpmath)"},
        {"address": f"F{HEADER_ROW}", "value": "Direct (library)"},
        {"address": f"G{HEADER_ROW}", "value": "1 - CDF or LN(PMF)"},
        {"address": f"H{HEADER_ROW}", "value": "Direct digits"},
        {"address": f"I{HEADER_ROW}", "value": "Alt. digits"},
        {"address": f"K{HEADER_ROW}", "value": "Direct status"},
        {"address": f"L{HEADER_ROW}", "value": "Alternative status"},
        {"address": f"M{HEADER_ROW}", "value": "Depth, -log10"},
        {"address": f"{HELPERS['reference_value']}{HEADER_ROW}", "value": "Reference value"},
        {"address": f"{HELPERS['direct_log']}{HEADER_ROW}", "value": "Direct (library)"},
        {"address": f"{HELPERS['alternative_log']}{HEADER_ROW}", "value": "1 - CDF or LN(PMF)"},
        {"address": f"{HELPERS['reference_log']}{HEADER_ROW}", "value": "Reference (mpmath)"},
        {"address": f"{LIMIT['wall_x']}{HEADER_ROW}", "value": "Smallest number an Excel cell holds (2.2E-308)"},
        {"address": f"{LIMIT['line_x']}{HEADER_ROW}", "value": "Smallest number an Excel cell holds (2.2E-308) "},
    ]
    numbers = [
        {"address": f"{LIMIT['wall_x']}{FIRST_ROW}", "value": depth_of_limit},
        {"address": f"{LIMIT['wall_x']}{FIRST_ROW + 1}", "value": depth_of_limit},
        {"address": f"{LIMIT['wall_y']}{FIRST_ROW}", "value": 0.0},
        {"address": f"{LIMIT['wall_y']}{FIRST_ROW + 1}", "value": float(FULL_DIGITS)},
        {"address": f"{LIMIT['line_x']}{FIRST_ROW}", "value": 0.0},
        {"address": f"{LIMIT['line_y']}{FIRST_ROW}", "value": -depth_of_limit},
        {"address": f"{LIMIT['line_y']}{FIRST_ROW + 1}", "value": -depth_of_limit},
    ]
    detail = [f"D{number}" for number in DETAIL_ROWS.values()]
    return {
        "schema": "tails-plan-1",
        "sheet": SHEET,
        "names": names,
        "columns": columns_to_write(data),
        "numbers": numbers,
        "formulas": formulas,
        "texts": texts,
        "defaults": {"scenario": data["default_scenario"], "reference_grade": data["grades"]["reference"],
                     "report_grade": data["grades"]["report"]},
        "messages": static_messages(data),
        "layout": {"scenario": "D6", "reference_grade": "D7", "report_grade": "D8",
                   "status": STATUS.replace("$", ""), "macros": MACROS, "inputs": "D6:D8",
                   "detail_rows": DETAIL_ROWS, "detail_merge": "D{row}:I{row}", "statuses_merge": STATUSES_MERGE,
                   "header_row": HEADER_ROW, "first_row": FIRST_ROW, "last_row": end,
                   "store_last_row": store_last_row(data), "rows": rows,
                   "table": f"C{HEADER_ROW}:M{end}", "values": f"F{FIRST_ROW}:G{end}",
                   "digits": f"H{FIRST_ROW}:I{end}", "statuses": f"K{FIRST_ROW}:L{end}", "depth": f"M{FIRST_ROW}:M{end}",
                   "points": f"C{FIRST_ROW}:C{end}", "reference_column": f"D{FIRST_ROW}:D{end}",
                   "text_columns": [f"{STORE['scenario']}{FIRST_ROW}:{STORE['scenario']}{store_last_row(data)}",
                                    f"{STORE['reference_text']}{FIRST_ROW}:{STORE['reference_text']}{store_last_row(data)}",
                                    f"{STORE['reference_short']}{FIRST_ROW}:{STORE['reference_short']}{store_last_row(data)}",
                                    f"{SCENARIOS['name']}{FIRST_ROW}:{SCENARIOS['alternative']}{FIRST_ROW + scenario_count - 1}",
                                    f"{SCENARIOS['contract']}{FIRST_ROW}:{SCENARIOS['method']}{FIRST_ROW + scenario_count - 1}"],
                   "hidden_from": "Q",
                   "outputs": ["D19:D20", "D21:I35", f"C{HEADER_ROW}:M{end}"], "detail": detail,
                   "spacers": ["C9:D10", "C11:D12", "C13:D14", "C15:D16"], "chart_area": "K5:O35"},
        "validation": {"scenario_list": "=TailsScenarioNames", "digits": [1, FULL_DIGITS]},
        "conditional": {"status_bad": "=TailsStatusBad",
                        "grid_inside": f"=ROW()-{HEADER_ROW}<=ROWS(TailsX)",
                        "status_rules": [{"text": text, "font": font, "fill": fill} for text, font, fill in STATUS_COLORS]},
        "charts": [
            {"name": "TailsDigits", "title": "Correct digits against the depth of the tail",
             "x_title": "Depth of the tail: -log10 of the probability", "y_title": "Correct digits",
             "x": "TailsDepth", "series": [
                 {"name": f"{HELPERS['direct_log']}{HEADER_ROW}", "x": "TailsDepth", "y": "TailsDirectDigits"},
                 {"name": f"{HELPERS['alternative_log']}{HEADER_ROW}", "x": "TailsDepth", "y": "TailsAlternativeDigits"},
                 {"name": f"{LIMIT['wall_x']}{HEADER_ROW}", "x": "TailsWallX", "y": "TailsWallY"}]},
            {"name": "TailsValues", "title": "log10 of the values against the depth of the tail",
             "x_title": "Depth of the tail: -log10 of the probability", "y_title": "log10 of the value",
             "x": "TailsDepth", "series": [
                 {"name": f"{HELPERS['reference_log']}{HEADER_ROW}", "x": "TailsDepth", "y": "TailsReferenceLog"},
                 {"name": f"{HELPERS['direct_log']}{HEADER_ROW}", "x": "TailsDepth", "y": "TailsDirectLog"},
                 {"name": f"{HELPERS['alternative_log']}{HEADER_ROW}", "x": "TailsDepth", "y": "TailsAlternativeLog"},
                 {"name": f"{LIMIT['line_x']}{HEADER_ROW}", "x": "TailsLineX", "y": "TailsLineY"}]},
        ],
        "full_digits": FULL_DIGITS,
        "smallest_normal": SMALLEST_NORMAL,
        "helper_reference_value": HELPERS["reference_value"],
    }


def formula_lengths(result: dict[str, Any]) -> list[tuple[str, int]]:
    """Length of every formula and name, to check them against Excel's limit of 8192 characters."""
    sizes = [(item["address"], len(item["formula"])) for item in result["formulas"]]
    sizes += [(item["name"], len(item["refers"])) for item in result["names"]]
    return sorted(sizes, key=lambda item: -item[1])


if __name__ == "__main__":
    for name, size in formula_lengths(plan())[:6]:
        print(size, name)
