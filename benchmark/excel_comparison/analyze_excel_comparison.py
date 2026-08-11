"""
Compare Excel's native statistical functions and this library against the same
50-digit references, and emit a markdown table for the README.

Run after Export_ExcelComparison has filled the two observed columns.
"""
import csv
from decimal import Decimal


def observed(text):
    text = text.strip()
    if not text or text.upper() == "ERROR":
        return None
    try:
        return float(sum(Decimal(part) for part in text.split(";")))
    except Exception:
        return None


def rel(v, ref):
    if v is None:
        return None
    if ref == 0:
        return 0.0 if v == 0 else float("inf")
    return abs(v - ref) / abs(ref)


def fmt(e):
    if e is None:
        return "#NUM!"
    if e == 0:
        return "exact"
    return f"{e:.1e}"


def main():
    rows = list(csv.DictReader(open("excel_comparison_grid.csv", encoding="utf-8"),
                               delimiter="|"))
    print(f"{'case':<30}{'Excel':>12}{'this library':>15}   verdict")
    wins = ties = losses = 0
    table = []
    for r in rows:
        ref = float(r["reference"])
        ex = rel(observed(r["observed_excel"]), ref)
        ks = rel(observed(r["observed_kstats"]), ref)
        if ex is None and ks is not None:
            verdict = "library returns a value, Excel errors"; wins += 1
        elif ex is None and ks is None:
            verdict = "both error"; ties += 1
        elif ks is None:
            verdict = "EXCEL BETTER (library errors)"; losses += 1
        elif ks < ex / 10:
            verdict = "library more accurate"; wins += 1
        elif ex < ks / 10:
            verdict = "EXCEL MORE ACCURATE"; losses += 1
        else:
            verdict = "comparable"; ties += 1
        print(f"{r['label']:<30}{fmt(ex):>12}{fmt(ks):>15}   {verdict}")
        table.append((r["label"], r["excel_formula"], fmt(ex), fmt(ks), verdict))

    print(f"\n  library better: {wins}   comparable: {ties}   Excel better: {losses}")
    if losses:
        print("  NOTE: at least one point favours Excel. That belongs in the README too.")

    print("\n--- markdown for the README ---\n")
    print("| Case | Excel formula | Excel rel. error | This library | ")
    print("| --- | --- | --- | --- |")
    for label, xl, ex, ks, _ in table:
        print(f"| {label} | `{xl}` | {ex} | {ks} |")


if __name__ == "__main__":
    main()
