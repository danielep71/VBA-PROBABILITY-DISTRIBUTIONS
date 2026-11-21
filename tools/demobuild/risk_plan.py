"""The Risk sheet (value at risk, expected shortfall, tails and default counts) as data: every formula, name,
validation rule and text.

Same method as discrete_plan.py: this module writes the plan that tools/demobuild/Evolve-Risk.ps1 applies in
desktop Excel. examples/demo/risk-cases.json (made by examples/demo/risk_references.py) holds the default
inputs and the cases that ci/Test-Risk.ps1 checks against mpmath; the sheet itself stores no reference, every
value on it is a live call to the library or arithmetic on such calls.

The sheet works in losses: L = -V R for a position of value V and a return R over the horizon, positive when
money is lost, in the currency of V. Daily returns have mean mu and volatility sigma and the days are
independent, so the loss over h days has mean m = -V mu h and standard deviation sd = V sigma sqrt(h). Two
models of the loss share them: a Normal, and a Student t with nu > 2 degrees of freedom scaled by
s = sd sqrt((nu - 2) / nu) so its variance is sd^2. For a confidence level a, the value at risk (VaR) is the
a-quantile of the loss and the expected shortfall (ES) the mean loss beyond it. Beside them, the exceedance
probability of a loss threshold, and the count of defaults of n independent loans (a Binomial): its
a-quantile and the loss it implies.

Layout, in the style of the Explorer:

    C6:D15    the inputs: value V, mean daily return, daily volatility, horizon in days, confidence level,
              Student-t degrees of freedom, loss threshold, number of loans, default probability, loss per
              default
    D19:D20   the input status and the macros
    C21:I38   the results: mean and standard deviation of the loss, Normal VaR and ES, the Student-t scale,
              VaR and ES, P(L > threshold) under both, the expected defaults, the default-count quantile, its
              loss and P(defaults > quantile); in F:I the formula of each result as written; then a reading
              of the two models, the assumptions, the Excel that computed the sheet and the provenance
    K5:O35    two charts: the two loss densities with a vertical line at each VaR, and the two exceedance
              probabilities P(L > x) on a log scale with a line at 1 - a
    C40:I141  the table: 101 losses from mean - 4 sd to mean + top sd, the two densities, the two exceedance
              probabilities and their ratio
    AF:AJ     hidden: the z of each row and the spilled helpers of the charts
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
CASES_FILE = ROOT / "examples" / "demo" / "risk-cases.json"

SHEET = "Risk"
INPUTS = {"value": "D6", "mean_return": "D7", "volatility": "D8", "horizon": "D9", "confidence": "D10",
          "df": "D11", "threshold": "D12", "loans": "D13", "default_probability": "D14", "loss_per_default": "D15"}
STATUS = "$D$19"
MACROS = "D20"
HEADER_ROW = 40
FIRST_ROW = 41
ROWS = 101
FIRST_Z = -4
MIN_TOP = 6
LIMITS = {"value": 1e12, "mean_return": 0.1, "volatility": 1, "horizon": 250, "confidence_min": 0.5,
          "confidence_max": 0.9999, "df": 1000, "threshold": 1e13, "loans": 10000, "loss_per_default": 1e12}
RESULT_ROWS = {
    "mean_loss": 21, "sd_loss": 22, "normal_var": 23, "normal_es": 24, "t_scale": 25, "t_var": 26, "t_es": 27,
    "normal_exceed": 28, "t_exceed": 29, "expected_defaults": 30, "default_quantile": 31, "default_var": 32,
    "default_exceed": 33, "reading": 34, "assumptions": 35, "excel_build": 37, "provenance": 38,
}
VALUE_KEYS = ("mean_loss", "sd_loss", "normal_var", "normal_es", "t_scale", "t_var", "t_es", "normal_exceed",
              "t_exceed", "expected_defaults", "default_quantile", "default_var", "default_exceed")
ASSUMPTIONS_MERGE = "D35:I36"
# The table: E is the narrow spacer column of the Explorer's layout.
TABLE = {"loss": "C", "normal_density": "D", "t_density": "F", "normal_exceed": "G", "t_exceed": "H", "ratio": "I"}
HELPERS = {"z": "AF", "normal_var_x": "AG", "var_height": "AH", "t_var_x": "AI", "tail_level": "AJ"}


def cell(key: str) -> str:
    return INPUTS[key]


def result(key: str) -> str:
    return f"D{RESULT_ROWS[key]}"


def absolute(address: str) -> str:
    column = address.rstrip("0123456789")
    return f"${column}${address[len(column):]}"


def ref(address: str) -> str:
    return f"{SHEET}!{absolute(address)}"


def load() -> dict[str, Any]:
    data: dict[str, Any] = json.loads(CASES_FILE.read_text(encoding="utf-8"))
    names = [case["name"] for case in data["cases"]]
    if len(names) != len(set(names)):
        raise ValueError("duplicate case")
    if data["default_case"] not in names:
        raise ValueError("the default case is not a case")
    if data["table"] != {"rows": ROWS, "first_z": FIRST_Z, "min_top": MIN_TOP}:
        raise ValueError("the table of the references is not the sheet's")
    return data


def defaults(data: dict[str, Any]) -> dict[str, Any]:
    case = next(item for item in data["cases"] if item["name"] == data["default_case"])
    return dict(case["inputs"])


def expressions() -> dict[str, str]:
    """Each result as the formula a user would type, with plain cell references: the sheet's cell is
    IF(status OK, this, NA()) and the text beside it is this formula."""
    value, mu, sigma, horizon, level, nu, threshold, loans, probability, per_default = (
        cell(key) for key in INPUTS)
    mean, sd, scale = result("mean_loss"), result("sd_loss"), result("t_scale")
    quantile_t = f"K_STATS_StudentT_InverseCumulative({level},{nu})"
    return {
        "mean_loss": f"-{value}*{mu}*{horizon}",
        "sd_loss": f"{value}*{sigma}*SQRT({horizon})",
        "normal_var": f"K_STATS_Normal_InverseCumulative({level},{mean},{sd})",
        "normal_es": (f"{mean}+{sd}*K_STATS_NormalStandard_Density(K_STATS_NormalStandard_InverseCumulative({level}))"
                      f"/(1-{level})"),
        "t_scale": f"{sd}*SQRT(({nu}-2)/{nu})",
        "t_var": f"{mean}+{scale}*{quantile_t}",
        "t_es": (f"{mean}+{scale}*({nu}+{quantile_t}^2)/({nu}-1)*K_STATS_StudentT_Density({quantile_t},{nu})"
                 f"/(1-{level})"),
        "normal_exceed": f"K_STATS_Normal_Survival({threshold},{mean},{sd})",
        "t_exceed": f"K_STATS_StudentT_Survival(({threshold}-{mean})/{scale},{nu})",
        "expected_defaults": f"K_STATS_Binomial_Mean({loans},{probability})",
        "default_quantile": f"K_STATS_Binomial_InverseCumulative({level},{loans},{probability})",
        "default_var": f"{result('default_quantile')}*{per_default}",
        "default_exceed": f"K_STATS_Binomial_Survival({result('default_quantile')},{loans},{probability})",
    }


def check_names() -> list[dict[str, str]]:
    """One TRUE/FALSE name per input: the data validation and the status read the same rules."""
    c = {key: ref(address) for key, address in INPUTS.items()}
    number = "ISNUMBER({c})"
    whole = "AND(ISNUMBER({c}),{c}=INT({c}))"
    return [
        {"name": "RiskValueOk", "refers": f"=IFERROR(AND({number.format(c=c['value'])},{c['value']}>0,{c['value']}<={LIMITS['value']:.0E}),FALSE)"},
        {"name": "RiskMeanReturnOk",
         "refers": f"=IFERROR(AND({number.format(c=c['mean_return'])},ABS({c['mean_return']})<={LIMITS['mean_return']}),FALSE)"},
        {"name": "RiskVolatilityOk",
         "refers": f"=IFERROR(AND({number.format(c=c['volatility'])},{c['volatility']}>0,{c['volatility']}<={LIMITS['volatility']}),FALSE)"},
        {"name": "RiskHorizonOk",
         "refers": f"=IFERROR(AND({whole.format(c=c['horizon'])},{c['horizon']}>=1,{c['horizon']}<={LIMITS['horizon']}),FALSE)"},
        {"name": "RiskConfidenceOk",
         "refers": (f"=IFERROR(AND({number.format(c=c['confidence'])},{c['confidence']}>={LIMITS['confidence_min']},"
                    f"{c['confidence']}<={LIMITS['confidence_max']}),FALSE)")},
        {"name": "RiskDfOk", "refers": f"=IFERROR(AND({number.format(c=c['df'])},{c['df']}>2,{c['df']}<={LIMITS['df']}),FALSE)"},
        {"name": "RiskThresholdOk",
         "refers": f"=IFERROR(AND({number.format(c=c['threshold'])},ABS({c['threshold']})<={LIMITS['threshold']:.0E}),FALSE)"},
        {"name": "RiskLoansOk",
         "refers": f"=IFERROR(AND({whole.format(c=c['loans'])},{c['loans']}>=1,{c['loans']}<={LIMITS['loans']}),FALSE)"},
        {"name": "RiskDefaultProbabilityOk",
         "refers": (f"=IFERROR(AND({number.format(c=c['default_probability'])},{c['default_probability']}>0,"
                    f"{c['default_probability']}<1),FALSE)")},
        {"name": "RiskLossPerDefaultOk",
         "refers": (f"=IFERROR(AND({number.format(c=c['loss_per_default'])},{c['loss_per_default']}>=0,"
                    f"{c['loss_per_default']}<={LIMITS['loss_per_default']:.0E}),FALSE)")},
    ]


CHECK_ORDER = (("RiskValueOk", "position value V"), ("RiskMeanReturnOk", "mean daily return"),
               ("RiskVolatilityOk", "daily volatility"), ("RiskHorizonOk", "horizon in days"),
               ("RiskConfidenceOk", "confidence level"),
               ("RiskDfOk", "degrees of freedom (above 2, or the variance is not finite)"),
               ("RiskThresholdOk", "loss threshold"), ("RiskLoansOk", "number of loans"),
               ("RiskDefaultProbabilityOk", "default probability"), ("RiskLossPerDefaultOk", "loss per default"))


def status_cell_formula() -> str:
    formula = '"OK"'
    for name, label in reversed(CHECK_ORDER):
        formula = f'IF(NOT({name}),"Check: {label}",{formula})'
    return "=" + formula


def ok() -> str:
    return f'{STATUS}="OK"'


def result_formulas() -> list[dict[str, str]]:
    text = expressions()
    return [{"address": result(key), "formula": f"=IF({ok()},{text[key]},NA())"} for key in VALUE_KEYS]


def shown_formulas() -> list[dict[str, str]]:
    """The formula of each result, as text in F:I of its row."""
    text = expressions()
    return [{"address": f"F{RESULT_ROWS[key]}", "formula": '="=' + text[key] + '"'} for key in VALUE_KEYS]


def reading_formula() -> str:
    """The two models side by side, in words with the numbers in them."""
    n_var, n_es, t_var, t_es = (f"$D${RESULT_ROWS[key]}" for key in ("normal_var", "normal_es", "t_var", "t_es"))
    exceed = f"$D${RESULT_ROWS['default_exceed']}"
    level = absolute(cell("confidence"))
    return (f'=IFERROR(IF({ok()},"At the same mean and variance, the Student t gives a VaR "&ROUND({t_var}/{n_var},3)'
            f'&" times the Normal VaR and an ES "&ROUND({t_es}/{n_es},3)&" times the Normal ES. ES is at least the '
            f'VaR in both models: "&IF(AND({n_es}>={n_var},{t_es}>={t_var}),"yes","NO")&". P(defaults > Q) = "'
            f'&ROUND({exceed},6)&IF({exceed}<=1-{level}," <= 1 - confidence."," CHECK: above 1 - confidence."),""),"")')


ASSUMPTIONS = (
    "Losses are positive amounts in the currency of V: L = -V R. Daily returns have the same mean and volatility "
    "and are independent, so over h days the mean grows with h and the volatility with sqrt(h). The Normal and "
    "the Student t share that mean and variance: the t is scaled by s = sd sqrt((nu - 2) / nu), which needs "
    "nu > 2 (with nu <= 2 the variance is not finite). VaR is the confidence-level quantile of the loss; ES is "
    "the mean loss beyond it, at least the VaR. The defaults are independent with the same probability: real "
    "defaults are correlated, so the count's tail is thinner here than in a portfolio. Synthetic numbers."
)


def assumptions_formula() -> str:
    return f'="{ASSUMPTIONS}"'


def table_formulas() -> list[dict[str, str]]:
    row = str(FIRST_ROW)
    mean, sd, scale = (absolute(result(key)) for key in ("mean_loss", "sd_loss", "t_scale"))
    nu, level = absolute(cell("df")), absolute(cell("confidence"))
    n_es, t_es = absolute(result("normal_es")), absolute(result("t_es"))

    def spill(expression: str) -> str:
        return f"=IF({ok()},{expression},NA())"

    def each(call: str) -> str:
        return f"MAP(RiskLoss,LAMBDA(xv,{call}))"

    top = f"MAX({MIN_TOP},INT((MAX({n_es},{t_es})-{mean})/{sd})+1)"
    return [
        {"address": f"{HELPERS['z']}{row}",
         "formula": spill(f"SEQUENCE({ROWS},1,0,1)*({top}-({FIRST_Z}))/{ROWS - 1}+({FIRST_Z})")},
        {"address": f"{TABLE['loss']}{row}", "formula": spill(f"{mean}+{sd}*RiskZ")},
        {"address": f"{TABLE['normal_density']}{row}", "formula": spill(each(f"K_STATS_Normal_Density(xv,{mean},{sd})"))},
        {"address": f"{TABLE['t_density']}{row}",
         "formula": spill(each(f"K_STATS_StudentT_Density((xv-{mean})/{scale},{nu})/{scale}"))},
        {"address": f"{TABLE['normal_exceed']}{row}", "formula": spill(each(f"K_STATS_Normal_Survival(xv,{mean},{sd})"))},
        {"address": f"{TABLE['t_exceed']}{row}",
         "formula": spill(each(f"K_STATS_StudentT_Survival((xv-{mean})/{scale},{nu})"))},
        {"address": f"{TABLE['ratio']}{row}", "formula": spill("RiskStudentExceed/RiskNormalExceed")},
        {"address": f"{HELPERS['normal_var_x']}{row}",
         "formula": spill(f"VSTACK({absolute(result('normal_var'))},{absolute(result('normal_var'))})")},
        {"address": f"{HELPERS['var_height']}{row}", "formula": spill("VSTACK(0,MAX(RiskNormalDensity,RiskStudentDensity))")},
        {"address": f"{HELPERS['t_var_x']}{row}",
         "formula": spill(f"VSTACK({absolute(result('t_var'))},{absolute(result('t_var'))})")},
        {"address": f"{HELPERS['tail_level']}{row}", "formula": spill(f"RiskZ*0+1-{level}")},
    ]


def static_messages() -> dict[str, list[str]]:
    lim = LIMITS
    return {
        "D6": ["Position value V", "Value of the position today, in any currency; every loss is in the same "
               "currency. Above 0 and at most 1E12.", "Invalid value", "Enter a number above 0 and at most 1E12."],
        "D7": ["Mean daily return", f"Expected return per day as a fraction (0.0003 = 0.03% a day). From -{lim['mean_return']} "
               f"to {lim['mean_return']}.", "Invalid mean return",
               f"Enter a number from -{lim['mean_return']} to {lim['mean_return']}."],
        "D8": ["Daily volatility", "Standard deviation of the daily return as a fraction (0.01 = 1% a day). Above 0 "
               f"and at most {lim['volatility']}.", "Invalid volatility", f"Enter a number above 0 and at most {lim['volatility']}."],
        "D9": ["Horizon in days", f"Number of days the loss covers: a whole number from 1 to {lim['horizon']}. The "
               "volatility grows with its square root.", "Invalid horizon", f"Enter a whole number from 1 to {lim['horizon']}."],
        "D10": ["Confidence level", f"Probability that the loss stays at or below the VaR: from {lim['confidence_min']} "
                f"to {lim['confidence_max']} (0.99 = 99%).", "Invalid confidence level",
                f"Enter a number from {lim['confidence_min']} to {lim['confidence_max']}."],
        "D11": ["Degrees of freedom nu", f"Student t compared with the Normal at the same variance: above 2 (with 2 or "
                f"less the variance is not finite) and at most {lim['df']}. Fewer degrees, heavier tails.",
                "Invalid degrees of freedom", f"Enter a number above 2 and at most {lim['df']}."],
        "D12": ["Loss threshold", "A loss amount in the currency of V: the sheet gives the probability of a larger "
                "loss under both models. A negative threshold is a gain.", "Invalid threshold",
                "Enter a number between -1E13 and 1E13."],
        "D13": ["Number of loans", f"Loans that default independently: a whole number from 1 to {lim['loans']:,}.",
                "Invalid number of loans", f"Enter a whole number from 1 to {lim['loans']:,}."],
        "D14": ["Default probability", "Probability that one loan defaults over the period, the same for every loan. "
                "Strictly between 0 and 1.", "Invalid probability", "Enter a number strictly between 0 and 1."],
        "D15": ["Loss per default", "Loss on one defaulted loan, in the currency of V (exposure times loss given "
                "default). From 0 to 1E12.", "Invalid loss per default", "Enter a number from 0 to 1E12."],
    }


def plan() -> dict[str, Any]:
    data = load()
    chosen = defaults(data)
    end = FIRST_ROW + ROWS - 1
    names = [*check_names(), {"name": "RiskStatusBad", "refers": f'={ref(STATUS.replace("$", ""))}<>"OK"'}]
    spills = {"RiskZ": HELPERS["z"], "RiskLoss": TABLE["loss"], "RiskNormalDensity": TABLE["normal_density"],
              "RiskStudentDensity": TABLE["t_density"], "RiskNormalExceed": TABLE["normal_exceed"],
              "RiskStudentExceed": TABLE["t_exceed"], "RiskRatio": TABLE["ratio"],
              "RiskNormalVaRX": HELPERS["normal_var_x"], "RiskVaRHeight": HELPERS["var_height"],
              "RiskStudentVaRX": HELPERS["t_var_x"], "RiskTailLevel": HELPERS["tail_level"]}
    names += [{"name": name, "refers": f"={SHEET}!${letter}${FIRST_ROW}#"} for name, letter in spills.items()]
    formulas = [
        {"address": STATUS.replace("$", ""), "formula": status_cell_formula()},
        {"address": MACROS, "formula": '=IF(ISERROR(K_STATS_NormalStandard_Cumulative(0)+0*COUNTA($D$6:$D$15)),"DISABLED","ENABLED")'},
        {"address": "F4", "formula": f'=IF(${MACROS[0]}${MACROS[1:]}="ENABLED","","MACROS ARE DISABLED: the results are not live")'},
        *result_formulas(),
        *shown_formulas(),
        {"address": result("reading"), "formula": reading_formula()},
        {"address": result("assumptions"), "formula": assumptions_formula()},
        *table_formulas(),
    ]
    environment = data["environment"]
    provenance = (f"Every value is a live call to the library. ci/Test-Risk.ps1 checks the sheet against "
                  f"{len(data['cases'])} cases computed with mpmath {environment['mpmath']} (Python {environment['python']}), "
                  "stored by examples/demo/risk_references.py.")
    labels = {
        "mean_loss": "Mean loss m over the horizon", "sd_loss": "Loss standard deviation sd",
        "normal_var": "Normal VaR", "normal_es": "Normal ES", "t_scale": "Student-t scale s",
        "t_var": "Student-t VaR", "t_es": "Student-t ES", "normal_exceed": "Normal P(L > threshold)",
        "t_exceed": "Student-t P(L > threshold)", "expected_defaults": "Expected defaults n p",
        "default_quantile": "Default count VaR Q", "default_var": "Default loss VaR",
        "default_exceed": "P(defaults > Q)", "reading": "Reading", "assumptions": "Assumptions",
        "excel_build": "Computed by", "provenance": "Provenance",
    }
    texts = [
        {"address": "B1", "value": "Risk Applications"},
        {"address": "B2", "value": "Value at risk and expected shortfall, Normal against Student t at the same variance, "
                                   "tail probabilities and default counts."},
        {"address": "C3", "value": "Edit the yellow inputs. Every result is a call to the embedded VBA API "
                                   "(K_STATS_ functions). Macros are required."},
        {"address": "C6", "value": "Position value V"},
        {"address": "C7", "value": "Mean daily return"},
        {"address": "C8", "value": "Daily volatility"},
        {"address": "C9", "value": "Horizon in days h"},
        {"address": "C10", "value": "Confidence level"},
        {"address": "C11", "value": "Student-t degrees of freedom"},
        {"address": "C12", "value": "Loss threshold"},
        {"address": "C13", "value": "Number of loans n"},
        {"address": "C14", "value": "Default probability p"},
        {"address": "C15", "value": "Loss per default"},
        {"address": "C19", "value": "Input status"},
        {"address": "C20", "value": "Macros"},
        *({"address": f"C{RESULT_ROWS[key]}", "value": label} for key, label in labels.items()),
        {"address": f"D{RESULT_ROWS['provenance']}", "value": provenance},
        {"address": "F5", "value": "Read the results"},
        {"address": "F6", "value": "Losses are positive amounts in the currency of V. The horizon scales the mean by h "
                                   "and the volatility by sqrt(h)."},
        {"address": "F8", "value": "VaR is the loss not exceeded with the confidence level; ES is the mean loss beyond "
                                   "the VaR, so ES >= VaR."},
        {"address": "F10", "value": "The Student t has the Normal's mean and variance (nu > 2): its tails are heavier, so "
                                    "far enough out its VaR, ES and P(L > x) are larger."},
        {"address": "F12", "value": "The default count is a Binomial of independent loans: its VaR Q is a whole number "
                                    "of defaults, the smallest with P(defaults <= Q) >= confidence."},
        {"address": "F14", "value": "Educational demo, not an accuracy certification or a risk model. The library's "
                                    "guarantees are in docs/."},
        {"address": "F16", "value": "Needs macros and Excel 365 or 2024. If Macros says DISABLED, unblock the file and reopen."},
        {"address": f"C{HEADER_ROW - 1}", "value": "The loss distribution under both models"},
        {"address": f"{TABLE['loss']}{HEADER_ROW}", "value": "Loss x"},
        {"address": f"{TABLE['normal_density']}{HEADER_ROW}", "value": "Normal density"},
        {"address": f"{TABLE['t_density']}{HEADER_ROW}", "value": "Student-t density"},
        {"address": f"{TABLE['normal_exceed']}{HEADER_ROW}", "value": "Normal P(L > x)"},
        {"address": f"{TABLE['t_exceed']}{HEADER_ROW}", "value": "Student-t P(L > x)"},
        {"address": f"{TABLE['ratio']}{HEADER_ROW}", "value": "Ratio t / Normal"},
        {"address": f"{HELPERS['z']}{HEADER_ROW}", "value": "z (sd from the mean)"},
        {"address": f"{HELPERS['normal_var_x']}{HEADER_ROW}", "value": "Normal VaR"},
        {"address": f"{HELPERS['var_height']}{HEADER_ROW}", "value": "VaR line height"},
        {"address": f"{HELPERS['t_var_x']}{HEADER_ROW}", "value": "Student-t VaR"},
        {"address": f"{HELPERS['tail_level']}{HEADER_ROW}", "value": "1 - confidence"},
    ]
    detail = [f"D{row}" for row in RESULT_ROWS.values()]
    return {
        "schema": "risk-plan-1",
        "sheet": SHEET,
        "names": names,
        "formulas": formulas,
        "texts": texts,
        "defaults": chosen,
        "messages": static_messages(),
        "layout": {"inputs": "D6:D15", "input_cells": INPUTS, "status": STATUS.replace("$", ""), "macros": MACROS,
                   "result_rows": RESULT_ROWS, "value_keys": list(VALUE_KEYS), "assumptions_merge": ASSUMPTIONS_MERGE,
                   "formula_merge": "F{row}:I{row}", "text_merge": "D{row}:I{row}", "header_row": HEADER_ROW,
                   "first_row": FIRST_ROW, "last_row": end, "rows": ROWS, "table": f"C{HEADER_ROW}:I{end}",
                   "losses": f"C{FIRST_ROW}:C{end}", "probabilities": f"D{FIRST_ROW}:I{end}",
                   "results": f"D{RESULT_ROWS['mean_loss']}:D{RESULT_ROWS['default_exceed']}",
                   "hidden_from": "Q", "spacers": ["C16:D16"],
                   "outputs": ["D19:D20", "D21:I38", f"C{HEADER_ROW}:I{end}"], "detail": detail, "chart_area": "K5:O35"},
        "validation": {"custom": {address: f"={name}" for address, (name, _) in zip(INPUTS.values(), CHECK_ORDER)}},
        "conditional": {"status_bad": "=RiskStatusBad"},
        "charts": [
            {"name": "RiskDensity", "title": "Loss density: Normal and Student t at the same variance, with each VaR",
             "series": [{"name": "Normal", "x": "RiskLoss", "y": "RiskNormalDensity"},
                        {"name": "Student t", "x": "RiskLoss", "y": "RiskStudentDensity"},
                        {"name": "Normal VaR", "x": "RiskNormalVaRX", "y": "RiskVaRHeight"},
                        {"name": "Student-t VaR", "x": "RiskStudentVaRX", "y": "RiskVaRHeight"}]},
            {"name": "RiskTail", "title": "Exceedance probability P(L > x), log scale",
             "series": [{"name": "Normal P(L > x)", "x": "RiskLoss", "y": "RiskNormalExceed"},
                        {"name": "Student-t P(L > x)", "x": "RiskLoss", "y": "RiskStudentExceed"},
                        {"name": "1 - confidence", "x": "RiskLoss", "y": "RiskTailLevel"}]},
        ],
        "limits": LIMITS,
    }


def formula_lengths(result_plan: dict[str, Any]) -> list[tuple[str, int]]:
    """Length of every formula and name, to check them against Excel's limit of 8192 characters."""
    sizes = [(item["address"], len(item["formula"])) for item in result_plan["formulas"]]
    sizes += [(item["name"], len(item["refers"])) for item in result_plan["names"]]
    return sorted(sizes, key=lambda item: -item[1])


if __name__ == "__main__":
    for name, size in formula_lengths(plan())[:6]:
        print(size, name)
