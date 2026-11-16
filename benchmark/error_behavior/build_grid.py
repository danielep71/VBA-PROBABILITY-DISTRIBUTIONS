"""
#331: the calls that show what every public function returns on bad input.

One file, error_behavior_grid.csv. For each declaration of
docs/PUBLIC_API.txt there is one valid call, and then that call with one
argument at a time replaced by each class of input below:

  nan, pinf, ninf   not a number and the two infinities, which only VBA code
                    can pass: a worksheet cell cannot hold them;
  negative, zero    -1 and 0;
  huge              1E308;
  below, above, one -0.1, 1.1 and 1, for an argument that is a probability;
  text, empty       a text and an empty cell, as a worksheet formula passes
                    them.

Many of these are valid for some arguments: zero is a valid X, and -1 a
valid mean. The grid records what is returned and judges nothing.

Columns (6): point_id, function, class, argument, arguments, observed.
arguments holds the call's arguments separated by ";". Excel fills observed
with M_STATS_PROBDIST_ERB.bas: the 16 hexadecimal digits of a number, a
worksheet error such as #NUM!, RAISED when the call raised a VBA error, or
TYPE and a number when it returned something else.

Run:
    python3 build_grid.py --write   (only when the file is absent)
    python3 build_grid.py --check   (the committed file is the design)
    python3 build_grid.py --dispatch   (the generated cases of the exporter)
    python3 build_grid.py --battery    (the generated region of the VBA test module)
"""
import argparse
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
FILE = os.path.join(HERE, "error_behavior_grid.csv")
API = os.path.join(ROOT, "docs", "PUBLIC_API.txt")
HEADER = ["point_id", "function", "class", "argument", "arguments", "observed"]
KEY_COLUMNS = 5

# A valid value for each parameter name of the public surface.
VALID = {
    "X": "3", "Z": "0.5", "LowerZ": "-1", "UpperZ": "1", "Probability": "0.3",
    "Alpha": "2", "Beta": "3", "Shape": "2", "ScaleParam": "1.5", "Lambda": "0.7",
    "Mean": "2", "StdDev": "1.5", "MeanLog": "0.5", "StdDevLog": "0.8",
    "DegreesFreedom": "5", "DegreesFreedom1": "5", "DegreesFreedom2": "10",
    "LowerBound": "0", "UpperBound": "10",
    "Trials": "10", "ProbSuccess": "0.3", "NumberSuccesses": "3", "NumberFailures": "3",
    "NumberEvents": "3", "SampleSuccesses": "2", "SampleSize": "10",
    "PopulationSuccesses": "8", "PopulationSize": "40",
}
# Where the general value is outside a family's support.
VALID_IN = {("Beta", "X"): "0.4"}
PROBABILITIES = ("Probability", "ProbSuccess")
CLASSES = (("nan", "NAN"), ("pinf", "PINF"), ("ninf", "NINF"), ("negative", "-1"),
           ("zero", "0"), ("huge", "1E308"), ("text", "TEXT"), ("empty", "EMPTY"))
PROBABILITY_CLASSES = (("below", "-0.1"), ("above", "1.1"), ("one", "1"))


def declarations(path=API):
    """[(name, [parameter names without Status])] in the order of the baseline."""
    out = []
    with open(path, encoding="utf-8") as handle:
        for line in handle:
            if line.startswith("#") or "\t" not in line:
                continue
            name, declaration = line.rstrip("\n").split("\t")[2:4]
            inside = re.search(r"\((.*)\)", declaration).group(1)
            names = [re.match(r"(?:Optional )?By(?:Val|Ref) (\w+) As", part).group(1)
                     for part in inside.split(", ") if part]
            out.append((name, [n for n in names if n != "Status"]))
    return out


def rows():
    out = []
    for name, parameters in declarations():
        family = name.split("_")[2]
        valid = [VALID_IN.get((family, p), VALID[p]) for p in parameters]
        out.append((name, "valid", "", valid))
        for index, parameter in enumerate(parameters):
            classes = CLASSES + (PROBABILITY_CLASSES if parameter in PROBABILITIES else ())
            for label, token in classes:
                changed = list(valid)
                changed[index] = token
                out.append((name, label, parameter, changed))
    return out


def dispatch_lines():
    """The cases of DirectCall in M_STATS_PROBDIST_ERB.bas, one per declaration."""
    out = []
    for name, parameters in declarations():
        arguments = ", ".join(f"A({index})" for index in range(len(parameters)))
        out.append(f'        Case "{name}"')
        out.append(f"            DirectCall = {name}({arguments})")
    return out


# The one public function that does not reject a non-finite argument, by its
# documented contract: a Double-to-Double helper for simulation loops.
NOT_VALIDATING = ("K_STATS_NormalStandard_InverseCumulativeFast",)
BATTERY_BEGIN = "' BEGIN generated public API error contract (#331)"
BATTERY_END = "' END generated public API error contract (#331)"


def vba_number(text):
    return text if "." in text else text + "#"


def battery_lines():
    """The generated region of tests/M_STATS_PROBDIST_TEST.bas."""
    declared = declarations()
    out = [BATTERY_BEGIN,
           "Private Sub Test_Core_ApiErrorContract()",
           "'",
           "'" + "=" * 78,
           "' Test_Core_ApiErrorContract",
           "'" + "-" * 78,
           "' PURPOSE",
           "'   One row for every public function (#331): a valid call returns no error,",
           "'   and each argument in turn set to not-a-number, +infinity and -infinity",
           "'   returns #NUM!. A function that raised instead of returning fails its row.",
           "'",
           "' EXCEPTION",
           "'   K_STATS_NormalStandard_InverseCumulativeFast returns a Double and",
           "'   validates nothing, by its documented contract; only its valid call is",
           "'   checked.",
           "'",
           "' GENERATED",
           "'   benchmark/error_behavior/build_grid.py --battery, from",
           "'   docs/PUBLIC_API.txt. Do not edit: test_survey.py fails when the two",
           "'   differ, so a new public function without a row fails the Python suite.",
           "'" + "=" * 78,
           "'",
           '    Debug.Print "-- Public API error contract (valid call; each argument non-finite is #NUM)"',
           ""]
    for name, parameters in declared:
        family = name.split("_")[2]
        valid = [VALID_IN.get((family, p), VALID[p]) for p in parameters]
        valid += ["0"] * (4 - len(valid))
        rejects = "False" if name in NOT_VALIDATING else "True"
        out.append(f'    ApiContractRow "{name}", {len(parameters)}, {rejects}, _')
        out.append("        " + ", ".join(vba_number(v) for v in valid))
    out += ["End Sub", "", "",
            "Private Function ApiContractCall(ByVal FunctionName As String, ByRef A() As Double) As Variant",
            "'",
            "'" + "=" * 78,
            "' ApiContractCall",
            "'" + "-" * 78,
            "' PURPOSE",
            "'   Calls one public function by name with the first of the four Doubles of",
            "'   A that it takes. A call that raises returns the text RAISED and the error",
            "'   number, which no assertion of the battery accepts.",
            "'",
            "' GENERATED",
            "'   benchmark/error_behavior/build_grid.py --battery. Do not edit.",
            "'" + "=" * 78,
            "'",
            "    On Error GoTo Raised",
            "",
            "    Select Case FunctionName"]
    for name, parameters in declared:
        arguments = ", ".join(f"A({index})" for index in range(len(parameters)))
        out.append(f'        Case "{name}"')
        out.append(f"            ApiContractCall = {name}({arguments})")
    out += ["        Case Else",
            '            ApiContractCall = "UNKNOWN"',
            "    End Select",
            "    Exit Function",
            "",
            "Raised:",
            '    ApiContractCall = "RAISED " & CStr(Err.Number)',
            "End Function",
            BATTERY_END]
    return out


def render():
    lines = [",".join(HEADER)]
    for number, (name, label, parameter, arguments) in enumerate(rows()):
        lines.append(",".join([f"e{number:04d}", name, label, parameter, ";".join(arguments), ""]))
    return "\n".join(lines) + "\n"


def identity(text):
    return [line.split(",")[:KEY_COLUMNS] for line in text.splitlines() if line.strip()]


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    mode = ap.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true")
    mode.add_argument("--check", action="store_true")
    mode.add_argument("--battery", action="store_true",
                      help="print the generated region of the VBA test module")
    mode.add_argument("--dispatch", action="store_true",
                      help="print the cases of DirectCall for the exporter")
    a = ap.parse_args(argv)
    if a.battery:
        print("\n".join(battery_lines()))
        return 0
    if a.dispatch:
        print("\n".join(dispatch_lines()))
        return 0
    text = render()
    if a.check:
        with open(FILE, encoding="utf-8") as handle:
            if identity(handle.read()) != identity(text):
                print("FAIL: the identity columns of the file differ from a fresh build")
                return 1
    elif os.path.exists(FILE):
        print(f"refusing: {FILE} exists; it may hold observations. Use --check.")
        return 1
    else:
        with open(FILE, "w", encoding="utf-8", newline="") as handle:
            handle.write(text)
    print(f"{'PASS' if a.check else 'wrote'}: {len(identity(text)) - 1} rows")
    return 0


if __name__ == "__main__":
    sys.exit(main())
