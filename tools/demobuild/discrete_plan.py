"""The Discrete sheet (Binomial and Poisson counts) as data: every formula, name, validation rule and text.

Same method as tails_plan.py: this module writes the plan that tools/demobuild/Evolve-Discrete.ps1
applies in desktop Excel. examples/demo/discrete-cases.json (made by examples/demo/discrete_references.py)
holds the default inputs and the cases that ci/Test-Discrete.ps1 checks against mpmath; the sheet itself
stores no reference, every value on it is a live call to the library.

The sheet shows one model at a time, chosen in D6:

  Binomial  n independent trials, each a success with the same probability p: X counts the successes.
  Poisson   events at a constant rate per unit of exposure, independent of each other: X counts the events
            over the exposure, with mean m = rate x exposure.

The default is a simplified count of defaults: 100 loans that default independently with probability
2% each (Binomial), and the same portfolio as 0.02 defaults per loan-year over 100 loan-years (Poisson).

Layout, in the style of the Explorer:

    C6:D12    the inputs: model, trials n, probability p, rate, exposure, threshold k, probability q
    D19:D20   the input status and the macros
    C21:I35   the results: mean, standard deviation, P(X = k), the four readings of a threshold (at most,
              fewer than, at least, more than), the quantile Q(q) with F(Q - 1) and F(Q), the check of
              F(Q - 1) < q <= F(Q); in F:I the formula of each result as written; then the assumptions of
              the model, the Excel that computed the sheet and the provenance
    K5:O35    two charts: the mass P(X = k) as columns with the counts k and above drawn in front, and the
              cumulative function F(k) as columns with the quantile's column drawn in front and q as a line
    C37:I..   the table: every count of a window around the mean, its mass and the four readings
    AG:AI     hidden: spilled helpers of the charts; AK the list of models

A threshold is a whole number: "fewer than k" is P(X <= k - 1) and "at least k" is P(X > k - 1), each
with its own call (a survival value is never 1 - CDF). At k = 0 the first is 0 and the second 1, without a
call: the library refuses a count below the support. The quantile is the smallest k with F(k) >= q; there
is in general no k with F(k) = q, so the sheet shows the two cumulative values that frame q.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
CASES_FILE = ROOT / "examples" / "demo" / "discrete-cases.json"

SHEET = "Discrete"
MODEL, TRIALS, PROBABILITY, RATE, EXPOSURE, THRESHOLD, LEVEL = "$D$6", "$D$7", "$D$8", "$D$9", "$D$10", "$D$11", "$D$12"
STATUS = "$D$19"
MACROS = "D20"
HEADER_ROW = 37
FIRST_ROW = 38
MODELS = ("Binomial", "Poisson")
LIMITS = {"trials": 10000, "rate": 1000, "exposure": 1000000, "mean_min": 0.001, "mean_max": 5000,
          "threshold": 1000000}
SPREAD = 5          # the table runs from mean - 5 sd to mean + 5 sd, widened to hold k and Q(q)
MAX_ROWS = 1000
RESULT_ROWS = {
    "mean": 21, "sd": 22, "mass": 23, "at_most": 24, "fewer_than": 25, "at_least": 26, "more_than": 27,
    "quantile": 28, "before_quantile": 29, "at_quantile": 30, "check": 31, "assumptions": 32,
    "excel_build": 34, "provenance": 35,
}
ASSUMPTIONS_MERGE = "D32:I33"
# The table: one column per reading. E is the narrow spacer column of the Explorer's layout.
TABLE = {"k": "C", "mass": "D", "at_most": "F", "fewer_than": "G", "at_least": "H", "more_than": "I"}
HELPERS = {"at_least_mass": "AG", "quantile_bar": "AH", "level": "AI"}
MODEL_LIST = "AK"
INPUT_ROWS = {"model": 6, "trials": 7, "probability": 8, "rate": 9, "exposure": 10, "threshold": 11, "level": 12}


def load() -> dict[str, Any]:
    data: dict[str, Any] = json.loads(CASES_FILE.read_text(encoding="utf-8"))
    names = [case["name"] for case in data["cases"]]
    if len(names) != len(set(names)):
        raise ValueError("duplicate case")
    if data["default_case"] not in names:
        raise ValueError("the default case is not a case")
    if data["window"]["spread"] != SPREAD or data["window"]["max_rows"] != MAX_ROWS:
        raise ValueError("the window of the references is not the sheet's")
    return data


def defaults(data: dict[str, Any]) -> dict[str, Any]:
    case = next(item for item in data["cases"] if item["name"] == data["default_case"])
    return {"model": case["model"], **case["inputs"]}


def ref(address: str) -> str:
    return f"{SHEET}!{address}"


def mean_text() -> str:
    return f"{RATE}*{EXPOSURE}"


def call(operation: str, point: str) -> str:
    """The selected model's call of one operation at one point."""
    binomial = f"K_STATS_Binomial_{operation}({point},{TRIALS},{PROBABILITY})"
    poisson = f"K_STATS_Poisson_{operation}({point},{mean_text()})"
    return f'IF({MODEL}="Binomial",{binomial},{poisson})'


def moment(operation: str) -> str:
    return (f'IF({MODEL}="Binomial",K_STATS_Binomial_{operation}({TRIALS},{PROBABILITY}),'
            f"K_STATS_Poisson_{operation}({mean_text()}))")


def quantile_call(level: str) -> str:
    return (f'IF({MODEL}="Binomial",K_STATS_Binomial_InverseCumulative({level},{TRIALS},{PROBABILITY}),'
            f"K_STATS_Poisson_InverseCumulative({level},{mean_text()}))")


def shown_call(operation: str, point: str) -> str:
    """The text of a call as the user would type it, for the formula column."""
    return (f'IF({MODEL}="Binomial","=K_STATS_Binomial_{operation}({point},D7,D8)",'
            f'"=K_STATS_Poisson_{operation}({point},D9*D10)")')


def check_names() -> list[dict[str, str]]:
    """One TRUE/FALSE name per input: the data validation and the status read the same rules."""
    trials, probability, rate, exposure, threshold, level = (ref(cell) for cell in (TRIALS, PROBABILITY, RATE, EXPOSURE, THRESHOLD, LEVEL))
    whole = "AND(ISNUMBER({c}),{c}=INT({c}))"
    return [
        {"name": "DiscreteModelOk", "refers": f'=OR({ref(MODEL)}="Binomial",{ref(MODEL)}="Poisson")'},
        {"name": "DiscreteTrialsOk",
         "refers": f"=IFERROR(AND({whole.format(c=trials)},{trials}>=1,{trials}<={LIMITS['trials']}),FALSE)"},
        {"name": "DiscreteProbabilityOk", "refers": f"=IFERROR(AND(ISNUMBER({probability}),{probability}>0,{probability}<1),FALSE)"},
        {"name": "DiscreteRateOk", "refers": f"=IFERROR(AND(ISNUMBER({rate}),{rate}>0,{rate}<={LIMITS['rate']}),FALSE)"},
        {"name": "DiscreteExposureOk",
         "refers": f"=IFERROR(AND(ISNUMBER({exposure}),{exposure}>0,{exposure}<={LIMITS['exposure']}),FALSE)"},
        {"name": "DiscreteMeanOk",
         "refers": f"=IFERROR(AND({rate}*{exposure}>={LIMITS['mean_min']},{rate}*{exposure}<={LIMITS['mean_max']}),FALSE)"},
        {"name": "DiscreteThresholdOk",
         "refers": (f"=IFERROR(AND({whole.format(c=threshold)},{threshold}>=0,{threshold}<={LIMITS['threshold']},"
                    f'OR({ref(MODEL)}<>"Binomial",{threshold}<={trials})),FALSE)')},
        {"name": "DiscreteLevelOk", "refers": f"=IFERROR(AND(ISNUMBER({level}),{level}>0,{level}<1),FALSE)"},
    ]


def window_names() -> list[dict[str, str]]:
    """The counts of the table. The names read the inputs, not the result cells, which depend on the status."""
    model, trials, probability, threshold, level = (ref(cell) for cell in (MODEL, TRIALS, PROBABILITY, THRESHOLD, LEVEL))
    lam = f"{ref(RATE)}*{ref(EXPOSURE)}"
    quantile = (f'IFERROR(IF({model}="Binomial",K_STATS_Binomial_InverseCumulative({level},{trials},{probability}),'
                f"K_STATS_Poisson_InverseCumulative({level},{lam})),{threshold})")
    return [
        {"name": "DiscreteCentre", "refers": f'=IF({model}="Binomial",{trials}*{probability},{lam})'},
        {"name": "DiscreteSpread",
         "refers": f'=IF({model}="Binomial",SQRT({trials}*{probability}*(1-{probability})),SQRT({lam}))'},
        {"name": "DiscreteQuantileSeen", "refers": "=" + quantile},
        {"name": "DiscreteLower",
         "refers": f"=MAX(0,MIN({threshold},DiscreteQuantileSeen,INT(DiscreteCentre-{SPREAD}*DiscreteSpread)))"},
        {"name": "DiscreteUpper",
         "refers": (f"=LET(top,MAX({threshold},DiscreteQuantileSeen,-INT(-(DiscreteCentre+{SPREAD}*DiscreteSpread))),"
                    f'IF({model}="Binomial",MIN({trials},top),top))')},
        {"name": "DiscreteTooWide", "refers": f"=IFERROR(DiscreteUpper-DiscreteLower+1>{MAX_ROWS},TRUE)"},
    ]


def status_cell_formula() -> str:
    parameters = (f'IF({MODEL}="Binomial",'
                  f'IF(NOT(DiscreteTrialsOk),"trials n",IF(NOT(DiscreteProbabilityOk),"probability of success p","")),'
                  f'IF(NOT(DiscreteRateOk),"event rate",IF(NOT(DiscreteExposureOk),"exposure",'
                  f'IF(NOT(DiscreteMeanOk),"rate x exposure",""))))')
    return (f'=IF(NOT(DiscreteModelOk),"Check: model",LET(parameter,{parameters},'
            f'IF(parameter<>"","Check: "&parameter,'
            f'IF(NOT(DiscreteThresholdOk),"Check: threshold k",'
            f'IF(NOT(DiscreteLevelOk),"Check: probability q",'
            f'IF(DiscreteTooWide,"Check: the table would pass {MAX_ROWS} rows","OK"))))))')


def ok() -> str:
    return f'{STATUS}="OK"'


def result_formulas() -> list[dict[str, str]]:
    at = {key: f"D{row}" for key, row in RESULT_ROWS.items()}
    k, quantile = THRESHOLD, f"$D${RESULT_ROWS['quantile']}"
    return [
        {"address": at["mean"], "formula": f"=IF({ok()},{moment('Mean')},NA())"},
        {"address": at["sd"], "formula": f"=IF({ok()},{moment('StdDev')},NA())"},
        {"address": at["mass"], "formula": f"=IF({ok()},{call('PMF', k)},NA())"},
        {"address": at["at_most"], "formula": f"=IF({ok()},{call('Cumulative', k)},NA())"},
        {"address": at["fewer_than"], "formula": f"=IF({ok()},IF({k}=0,0,{call('Cumulative', k + '-1')}),NA())"},
        {"address": at["at_least"], "formula": f"=IF({ok()},IF({k}=0,1,{call('Survival', k + '-1')}),NA())"},
        {"address": at["more_than"], "formula": f"=IF({ok()},{call('Survival', k)},NA())"},
        {"address": at["quantile"], "formula": f"=IF({ok()},{quantile_call(LEVEL)},NA())"},
        {"address": at["before_quantile"],
         "formula": f"=IFERROR(IF({quantile}=0,0,{call('Cumulative', quantile + '-1')}),NA())"},
        {"address": at["at_quantile"], "formula": f"=IFERROR({call('Cumulative', quantile)},NA())"},
        {"address": at["check"], "formula": check_formula()},
        {"address": at["assumptions"], "formula": assumptions_formula()},
    ]


def check_formula() -> str:
    """The quantile's two inequalities with the numbers in them, or what is wrong."""
    q, before, after = (f"$D${RESULT_ROWS[key]}" for key in ("quantile", "before_quantile", "at_quantile"))
    holds = (f'"F("&({q}-1)&") = "&{before}&" < q = "&{LEVEL}&" <= F("&{q}&") = "&{after}'
             f'&": Q = "&{q}&" is the smallest k with F(k) >= q."')
    return (f'=IFERROR(IF(AND({before}<{LEVEL},{LEVEL}<={after}),{holds},'
            f'"CHECK: F(Q - 1) < q <= F(Q) does not hold"),"")')


ASSUMPTIONS = {
    "Binomial": ("Binomial: a fixed number n of trials, independent of each other, each a success with the same "
                 "probability p; X counts the successes, from 0 to n. Mean n p, variance n p (1 - p). The default "
                 "example: 100 loans that default independently with probability 2% each; X is the number of "
                 "defaults. Real defaults are correlated: independence is a simplification."),
    "Poisson": ("Poisson: events one at a time, independent of each other, at a constant rate per unit of "
                "exposure (time, loan-years, policies); X counts the events over the exposure, from 0 upwards. "
                "Mean m = rate x exposure, which is also the variance. The default example: 0.02 defaults per "
                "loan-year over 100 loan-years, a mean of 2, the Poisson counterpart of the Binomial example."),
}


def assumptions_formula() -> str:
    return f'=IF({MODEL}="Binomial","{ASSUMPTIONS["Binomial"]}",IF({MODEL}="Poisson","{ASSUMPTIONS["Poisson"]}",""))'


def shown_formulas() -> list[dict[str, str]]:
    """The formula of each result, as text in F:I of its row."""
    rows = {
        "mean": f'IF({MODEL}="Binomial","=K_STATS_Binomial_Mean(D7,D8)","=K_STATS_Poisson_Mean(D9*D10)")',
        "sd": f'IF({MODEL}="Binomial","=K_STATS_Binomial_StdDev(D7,D8)","=K_STATS_Poisson_StdDev(D9*D10)")',
        "mass": shown_call("PMF", "D11"),
        "at_most": shown_call("Cumulative", "D11"),
        "fewer_than": f'"P(X <= k - 1): "&{shown_call("Cumulative", "D11-1")}&", and 0 when k = 0"',
        "at_least": f'"P(X > k - 1): "&{shown_call("Survival", "D11-1")}&", and 1 when k = 0"',
        "more_than": f'{shown_call("Survival", "D11")}&", not 1 - P(X <= k)"',
        "quantile": (f'IF({MODEL}="Binomial","=K_STATS_Binomial_InverseCumulative(D12,D7,D8)",'
                     f'"=K_STATS_Poisson_InverseCumulative(D12,D9*D10)")&": the smallest k with F(k) >= q"'),
        "before_quantile": f'{shown_call("Cumulative", "D28-1")}&", and 0 when Q = 0"',
        "at_quantile": shown_call("Cumulative", "D28"),
    }
    return [{"address": f"F{RESULT_ROWS[key]}", "formula": f"=IFERROR({text},\"\")"} for key, text in rows.items()]


def table_formulas() -> list[dict[str, str]]:
    row = str(FIRST_ROW)

    def spill(expression: str) -> str:
        return f"=IF({ok()},{expression},NA())"

    def each(operation: str) -> str:
        return f"MAP(DiscreteK,LAMBDA(kval,{call(operation, 'kval')}))"

    return [
        {"address": f"{TABLE['k']}{row}", "formula": spill("SEQUENCE(DiscreteUpper-DiscreteLower+1,,DiscreteLower)")},
        {"address": f"{TABLE['mass']}{row}", "formula": spill(each("PMF"))},
        {"address": f"{TABLE['at_most']}{row}", "formula": spill(each("Cumulative"))},
        {"address": f"{TABLE['more_than']}{row}", "formula": spill(each("Survival"))},
        # P(X < k) is the row above's P(X <= k), and P(X >= k) the row above's P(X > k): one call for the
        # first row of the window, none for the others.
        {"address": f"{TABLE['fewer_than']}{row}", "formula": spill(
            f"MAP(DiscreteK,LAMBDA(kval,IF(kval=0,0,IF(kval=DiscreteLower,{call('Cumulative', 'kval-1')},"
            f"INDEX(DiscreteAtMost,kval-DiscreteLower)))))")},
        {"address": f"{TABLE['at_least']}{row}", "formula": spill(
            f"MAP(DiscreteK,LAMBDA(kval,IF(kval=0,1,IF(kval=DiscreteLower,{call('Survival', 'kval-1')},"
            f"INDEX(DiscreteMoreThan,kval-DiscreteLower)))))")},
        {"address": f"{HELPERS['at_least_mass']}{row}", "formula": spill(f"IF(DiscreteK>={THRESHOLD},DiscreteMass,0)")},
        {"address": f"{HELPERS['quantile_bar']}{row}",
         "formula": spill(f"IF(DiscreteK=$D${RESULT_ROWS['quantile']},DiscreteAtMost,0)")},
        {"address": f"{HELPERS['level']}{row}", "formula": spill(f"DiscreteK*0+{LEVEL}")},
    ]


def static_messages() -> dict[str, list[str]]:
    return {
        "D6": ["Model", "Binomial: n independent trials with the same probability p. Poisson: events at a constant "
               "rate per unit of exposure. Each model reads its own two inputs; the others are greyed.",
               "Unknown model", "Choose Binomial or Poisson."],
        "D7": ["Trials n (Binomial)", f"Number of independent trials, for example loans in a portfolio. A whole "
               f"number from 1 to {LIMITS['trials']:,}.", "Invalid number of trials",
               f"Enter a whole number from 1 to {LIMITS['trials']:,}."],
        "D8": ["Probability of success p", "Probability that one trial is a success (a loan defaults), the same "
               "for every trial. Strictly between 0 and 1.", "Invalid probability", "Enter a number strictly between 0 and 1."],
        "D9": ["Event rate (Poisson)", f"Expected events per unit of exposure, for example defaults per loan-year. "
               f"Above 0 and at most {LIMITS['rate']:,}.", "Invalid rate",
               f"Enter a number above 0 and at most {LIMITS['rate']:,}."],
        "D10": ["Exposure (Poisson)", "Units of exposure observed, for example loan-years. The Poisson mean is rate "
                f"x exposure, from {LIMITS['mean_min']} to {LIMITS['mean_max']:,}.", "Invalid exposure",
                f"Enter a number above 0 and at most {LIMITS['exposure']:,}; rate x exposure from "
                f"{LIMITS['mean_min']} to {LIMITS['mean_max']:,}."],
        "D11": ["Threshold k", "A count: a whole number from 0 (at most n for the Binomial). Read at most k, fewer "
                "than k, at least k and more than k.", "Invalid threshold",
                "Enter a whole number from 0, not above n for the Binomial."],
        "D12": ["Probability q", "Probability for the quantile: the smallest k with F(k) >= q. Strictly between "
                "0 and 1.", "Invalid probability", "Enter a number strictly between 0 and 1."],
    }


def plan() -> dict[str, Any]:
    data = load()
    chosen = defaults(data)
    end = FIRST_ROW + MAX_ROWS - 1
    names = [
        {"name": "DiscreteModels", "refers": "=" + ref(f"${MODEL_LIST}${FIRST_ROW}:${MODEL_LIST}${FIRST_ROW + len(MODELS) - 1}")},
        *check_names(),
        *window_names(),
        {"name": "DiscreteStatusBad", "refers": f'={ref(STATUS)}<>"OK"'},
        {"name": "DiscreteBinomialUnused", "refers": f'={ref(MODEL)}="Poisson"'},
        {"name": "DiscretePoissonUnused", "refers": f'={ref(MODEL)}="Binomial"'},
    ]
    spills = {"DiscreteK": TABLE["k"], "DiscreteMass": TABLE["mass"], "DiscreteAtMost": TABLE["at_most"],
              "DiscreteFewerThan": TABLE["fewer_than"], "DiscreteAtLeast": TABLE["at_least"],
              "DiscreteMoreThan": TABLE["more_than"], "DiscreteAtLeastMass": HELPERS["at_least_mass"],
              "DiscreteQuantileBar": HELPERS["quantile_bar"], "DiscreteLevel": HELPERS["level"]}
    names += [{"name": name, "refers": "=" + ref(f"${letter}${FIRST_ROW}#")} for name, letter in spills.items()]
    formulas = [
        {"address": STATUS.replace("$", ""), "formula": status_cell_formula()},
        {"address": MACROS, "formula": '=IF(ISERROR(K_STATS_NormalStandard_Cumulative(0)+0*COUNTA($D$6:$D$12)),"DISABLED","ENABLED")'},
        {"address": "F4", "formula": f'=IF(${MACROS[0]}${MACROS[1:]}="ENABLED","","MACROS ARE DISABLED: the results are not live")'},
        *result_formulas(),
        *shown_formulas(),
        *table_formulas(),
    ]
    environment = data["environment"]
    provenance = (f"Every value is a live call to the library. ci/Test-Discrete.ps1 checks the sheet against "
                  f"{len(data['cases'])} cases computed with mpmath {environment['mpmath']} (Python {environment['python']}), "
                  "stored by examples/demo/discrete_references.py.")
    labels = {
        "mean": "Mean", "sd": "Standard deviation", "mass": "P(X = k)", "at_most": "At most k: P(X <= k)",
        "fewer_than": "Fewer than k: P(X < k)", "at_least": "At least k: P(X >= k)", "more_than": "More than k: P(X > k)",
        "quantile": "Quantile Q(q)", "before_quantile": "F(Q - 1)", "at_quantile": "F(Q)", "check": "Quantile check",
        "assumptions": "Assumptions", "excel_build": "Computed by", "provenance": "Provenance",
    }
    texts = [
        {"address": "B1", "value": "Discrete Distributions"},
        {"address": "B2", "value": "Binomial and Poisson counts: thresholds read four ways, and quantiles that are steps."},
        {"address": "C3", "value": "Choose a model and edit the yellow inputs. Every result is a call to the embedded VBA API "
                                   "(K_STATS_ functions). Macros are required."},
        {"address": "C6", "value": "Model"},
        {"address": "C7", "value": "Trials n (Binomial)"},
        {"address": "C8", "value": "Probability of success p"},
        {"address": "C9", "value": "Event rate per unit (Poisson)"},
        {"address": "C10", "value": "Exposure, units (Poisson)"},
        {"address": "C11", "value": "Threshold k"},
        {"address": "C12", "value": "Probability q"},
        {"address": "C19", "value": "Input status"},
        {"address": "C20", "value": "Macros"},
        *({"address": f"C{RESULT_ROWS[key]}", "value": label} for key, label in labels.items()),
        {"address": f"D{RESULT_ROWS['provenance']}", "value": provenance},
        {"address": "F5", "value": "Read the results"},
        {"address": "F6", "value": "Yellow cells are editable. The inputs of the other model are greyed and ignored. "
                                   "The defaults count the defaults of 100 independent loans."},
        {"address": "F8", "value": "A count is a whole number, so a threshold has four readings: at most k includes k, fewer "
                                   "than k is at most k - 1, at least k includes k, more than k is at least k + 1."},
        {"address": "F10", "value": "P(X > k) is the library's survival function, not 1 - P(X <= k): the subtraction "
                                    "loses a small upper tail."},
        {"address": "F12", "value": "The quantile is a step: Q(q) is the smallest k with F(k) >= q, so F(Q - 1) < q <= F(Q). "
                                    "F(Q) is above q unless q falls exactly on a step."},
        {"address": "F14", "value": "Educational demo, not an accuracy certification. The library's guarantees are in docs/."},
        {"address": "F16", "value": "Needs macros and Excel 365 or 2024. If Macros says DISABLED, unblock the file and reopen."},
        {"address": f"C{HEADER_ROW - 1}", "value": "Every count of the window around the mean"},
        {"address": f"{TABLE['k']}{HEADER_ROW}", "value": "Count k"},
        {"address": f"{TABLE['mass']}{HEADER_ROW}", "value": "P(X = k)"},
        {"address": f"{TABLE['at_most']}{HEADER_ROW}", "value": "At most k: P(X <= k)"},
        {"address": f"{TABLE['fewer_than']}{HEADER_ROW}", "value": "Fewer than k: P(X < k)"},
        {"address": f"{TABLE['at_least']}{HEADER_ROW}", "value": "At least k: P(X >= k)"},
        {"address": f"{TABLE['more_than']}{HEADER_ROW}", "value": "More than k: P(X > k)"},
        {"address": f"{HELPERS['at_least_mass']}{HEADER_ROW}", "value": "At least k"},
        {"address": f"{HELPERS['quantile_bar']}{HEADER_ROW}", "value": "F(Q), the quantile"},
        {"address": f"{HELPERS['level']}{HEADER_ROW}", "value": "Probability q"},
        *({"address": f"{MODEL_LIST}{FIRST_ROW + index}", "value": model} for index, model in enumerate(MODELS)),
    ]
    detail = [f"D{row}" for row in RESULT_ROWS.values()]
    return {
        "schema": "discrete-plan-1",
        "sheet": SHEET,
        "names": names,
        "formulas": formulas,
        "texts": texts,
        "defaults": chosen,
        "messages": static_messages(),
        "layout": {"inputs": "D6:D12", "input_rows": INPUT_ROWS, "status": STATUS.replace("$", ""), "macros": MACROS,
                   "result_rows": RESULT_ROWS, "assumptions_merge": ASSUMPTIONS_MERGE, "formula_merge": "F{row}:I{row}",
                   "text_merge": "D{row}:I{row}", "header_row": HEADER_ROW, "first_row": FIRST_ROW, "last_row": end,
                   "table": f"C{HEADER_ROW}:I{end}", "counts": f"C{FIRST_ROW}:C{end}",
                   "probabilities": f"D{FIRST_ROW}:I{end}", "results": f"D{RESULT_ROWS['mean']}:D{RESULT_ROWS['at_quantile']}",
                   "binomial_inputs": "D7:D8", "poisson_inputs": "D9:D10",
                   "hidden_from": "Q", "spacers": ["C13:D14", "C15:D16"],
                   "outputs": ["D19:D20", "D21:I35", f"C{HEADER_ROW}:I{end}"], "detail": detail, "chart_area": "K5:O35"},
        "validation": {"model_list": "=DiscreteModels",
                       "custom": {"D7": "=DiscreteTrialsOk", "D8": "=DiscreteProbabilityOk", "D9": "=DiscreteRateOk",
                                  "D10": "=DiscreteExposureOk", "D11": "=DiscreteThresholdOk", "D12": "=DiscreteLevelOk"}},
        "conditional": {"status_bad": "=DiscreteStatusBad", "binomial_unused": "=DiscreteBinomialUnused",
                        "poisson_unused": "=DiscretePoissonUnused",
                        "grid_inside": f"=ROW()-{HEADER_ROW}<=ROWS(DiscreteK)",
                        "grid_outside": f"=ROW()-{HEADER_ROW}>ROWS(DiscreteK)",
                        "threshold_cell": f"={THRESHOLD}", "quantile_cell": f"=$D${RESULT_ROWS['quantile']}"},
        "charts": [
            {"name": "DiscreteMass", "title": "Probability mass P(X = k), with P(X >= k) drawn in front",
             "series": [{"name": "P(X = k)", "x": "DiscreteK", "y": "DiscreteMass"},
                        {"name": "At least k", "x": "DiscreteK", "y": "DiscreteAtLeastMass"}]},
            {"name": "DiscreteCumulative", "title": "Cumulative F(k) = P(X <= k), with the quantile and q",
             "series": [{"name": "F(k)", "x": "DiscreteK", "y": "DiscreteAtMost"},
                        {"name": "F(Q), the quantile", "x": "DiscreteK", "y": "DiscreteQuantileBar"},
                        {"name": "Probability q", "x": "DiscreteK", "y": "DiscreteLevel"}]},
        ],
        "limits": LIMITS,
        "max_rows": MAX_ROWS,
    }


def formula_lengths(result: dict[str, Any]) -> list[tuple[str, int]]:
    """Length of every formula and name, to check them against Excel's limit of 8192 characters."""
    sizes = [(item["address"], len(item["formula"])) for item in result["formulas"]]
    sizes += [(item["name"], len(item["refers"])) for item in result["names"]]
    return sorted(sizes, key=lambda item: -item[1])


if __name__ == "__main__":
    for name, size in formula_lengths(plan())[:6]:
        print(size, name)
