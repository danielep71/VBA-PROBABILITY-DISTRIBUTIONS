"""
Compare Excel's native statistical functions and this library against the same
50-digit references.

TWO SEPARATE QUESTIONS, TWO SEPARATE COLUMNS
    "Which is more accurate" and "is this fit for use" are different questions,
    and answering them in one column misleads. This script reports both:

      Closer to truth  - a plain comparison. Often true and meaningless: being
                         closer by 1E-14 versus 1E-16 is a real fact about the
                         two implementations and no basis for any conclusion.

      Fit for use      - an absolute judgment against the stated thresholds
                         below. This is the column that should drive decisions.

    An earlier version printed a single verdict and a win/loss tally. It gave
    "1.8E-14 vs 1.8E-16" the same weight as "no correct digits vs exact", and so
    reported a roughly even contest when the real picture is "never meaningfully
    behind, occasionally far ahead".

THE THRESHOLDS, AND WHY THEY SIT WHERE THEY DO
    REFERENCE_DIGITS = 12
        Safe as a BUILDING BLOCK. A value this accurate can be fed into further
        computation - iteration, root-finding, accumulation over many calls -
        without its error becoming visible in the result. This is the standard
        the library's own accuracy contracts are written to, and it is
        deliberately far stricter than any single reported number needs.

    REPORT_DIGITS = 6
        Safe as a DIRECTLY REPORTED statistic. No published p-value, critical
        value or test statistic carries more than about six significant figures,
        so an error below this cannot change a reported result.

    Below REPORT_DIGITS the value could change what a user reports. At zero
    correct digits it is not an approximation of the answer at all, and is
    reported as wrong rather than as inaccurate.

    These numbers are a judgment. They are stated here so a reader can disagree
    with them explicitly instead of having to infer them from a verdict.
"""
import csv
import math
from decimal import Decimal

# A Double carries about 15-17 significant decimal digits, so anything at or
# beyond this is reported as "full" rather than a spuriously precise count.
FULL_PRECISION = 15.0
# Safe as a building block for further computation. See the module docstring.
REFERENCE_DIGITS = 12.0
# Safe as a directly reported statistic. See the module docstring.
REPORT_DIGITS = 6.0


def observed(text):
    text = text.strip()
    if not text or text.upper() in ("ERROR", "NONE"):
        return None
    try:
        return float(sum(Decimal(part) for part in text.split(";")))
    except Exception:
        return None


def no_equivalent(row):
    """True when Excel has no function for this case at all."""
    return (row["excel_formula"].strip() == "-"
            or row["observed_excel"].strip().upper() == "NONE")


def digits(value, ref):
    """Correct significant digits, i.e. -log10 of the relative error."""
    if value is None:
        return None
    if ref == 0:
        return FULL_PRECISION if value == 0 else 0.0
    rel = abs(value - ref) / abs(ref)
    if rel == 0:
        return 99.0
    if rel >= 1:
        return 0.0
    return -math.log10(rel)


def fmt_digits(d):
    if d is None:
        return "n/a"
    if d >= FULL_PRECISION:
        return "full"
    return f"{d:.0f}"


def grade(d, absent=False):
    """Absolute fitness, independent of what the other implementation did."""
    if absent:
        return "no function"
    if d is None:
        return "unavailable"
    if d <= 0:
        return "wrong"
    if d >= REFERENCE_DIGITS:
        return "reference"
    if d >= REPORT_DIGITS:
        return "report"
    return "inadequate"


def closer(dx, dk, absent):
    """Plain comparison, with no claim that the difference matters."""
    if absent:
        return "-"
    if dx is None and dk is None:
        return "-"
    if dx is None:
        return "library"
    if dk is None:
        return "Excel"
    if abs(dx - dk) < 1.0:
        return "tie"
    return "Excel" if dx > dk else "library"


def fitness(gx, gk):
    """One phrase describing whether the difference can affect a user."""
    if gx == "no function" and gk in ("reference", "report"):
        return "**Excel has no such function**"
    if gx in ("wrong", "unavailable") and gk in ("reference", "report"):
        return "**only the library gives a usable answer**"
    if gk in ("wrong", "unavailable") and gx in ("reference", "report"):
        return "**only Excel gives a usable answer**"
    if gx == "reference" and gk == "reference":
        return "both reference grade"
    if gx in ("reference", "report") and gk in ("reference", "report"):
        return "both usable (Excel " + gx + ", library " + gk + ")"
    return "Excel " + gx + ", library " + gk


def main():
    rows = list(csv.DictReader(open("excel_comparison_grid.csv", encoding="utf-8"),
                               delimiter="|"))
    out = []
    for r in rows:
        ref = float(r["reference"])
        absent = no_equivalent(r)
        dx = digits(observed(r["observed_excel"]), ref)
        dk = digits(observed(r["observed_kstats"]), ref)
        gx, gk = grade(dx, absent), grade(dk)
        out.append((r["label"], r["excel_formula"], dx, dk, gx, gk,
                    closer(dx, dk, absent), fitness(gx, gk)))

    print("case".ljust(30) + "Excel".rjust(7) + "lib".rjust(6) + "  "
          + "closer".ljust(9) + "fit for use")
    for label, _, dx, dk, gx, gk, cl, fit in out:
        ex = "none" if gx == "no function" else fmt_digits(dx)
        print(label.ljust(30) + ex.rjust(7) + fmt_digits(dk).rjust(6)
              + "  " + cl.ljust(9) + fit.replace("**", ""))

    material = [o for o in out if "only" in o[7] or "no such function" in o[7]]
    print("")
    print("  cases where one side cannot serve and the other can: "
          + str(len(material)))
    for o in material:
        print("    " + o[0].ljust(30) + o[7].replace("**", ""))

    below = [o for o in out if o[5] not in ("reference",)]
    if below:
        print("")
        print("  library below reference grade ("
              + str(int(REFERENCE_DIGITS)) + " digits): " + str(len(below)))
        for o in below:
            print("    " + o[0].ljust(30) + fmt_digits(o[3])
                  + " digits (" + o[5] + ")")

    print("")
    print("--- markdown for the README ---")
    print("")
    print("| Case | Excel formula | Excel | This library | Closer | Fit for use |")
    print("| --- | --- | ---: | ---: | :---: | --- |")
    for label, xl, dx, dk, gx, gk, cl, fit in out:
        ex = "none" if gx == "no function" else fmt_digits(dx)
        formula = "*(none)*" if xl.strip() == "-" else "`" + xl + "`"
        print("| " + label + " | " + formula + " | " + ex + " | "
              + fmt_digits(dk) + " | " + cl + " | " + fit + " |")


if __name__ == "__main__":
    main()
