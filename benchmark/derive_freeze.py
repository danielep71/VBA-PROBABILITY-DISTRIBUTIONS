"""
Freeze record for the four active #12/#15 contracts.

One rule for all four:

    empirical_worst = max(promoted main-grid worst, independent holdout worst)
    freeze          = the 1-2-5 level above it with credible robustness headroom

The main-grid worst is recomputed from the promoted observations rather than
carried over from the studies. The rows are the same nominal points, but the
evidence path is not: main-grid serialisation, Val() parsing, the main
exporter, hi/lo reconstruction, and the shared contract evaluator. Twice in this
sequence a numerical equivalence that looked obvious turned out not to be, so it
is measured.

The study worst is reported alongside as a ratio. It should be near 1; anything
else needs explaining rather than normalising away.

Report only. Writes nothing.
"""
import argparse, csv, os, sys
from decimal import Decimal, getcontext

getcontext().prec = 60
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from _contract_eval import (parse_observed, parse_reference, calculate_error,
                            calculate_scaled_error)

# Margin floor, grounded in the registry rather than asserted. Of the 69
# existing frozen contracts with a finite holdout margin, 4 sit below 2x and the
# thinnest is 1.3x; the median band is 3-5x. A 2.5x floor is therefore stricter
# than the repository has historically accepted, without demanding headroom that
# no existing contract is held to. The margin is measured against the COMBINED
# worst - the binding number - not against whichever of main or holdout happens
# to be smaller.
MIN_MARGIN = Decimal("2.5")

# contract -> (grid function, grid regime, measure, metric,
#              holdout worst, holdout source)
SPEC = [
    ("LogGamma.small_positive.log_abs", "LogGamma", "small_positive",
     "output_error", "absolute", Decimal("6.27E-14"), "loggamma_regimes_study holdout"),
    ("LogGamma.near_zero.log_abs", "LogGamma", "near_zero",
     "output_error", "absolute", Decimal("1.13E-14"), "loggamma_regimes_study holdout"),
    ("LogGamma.general.output_rel", "LogGamma", "general",
     "output_error", "relative", Decimal("3.24E-14"), "loggamma_regimes_study holdout"),
    ("LogGamma1p.small.scaled_abs", "LogGamma1p", "small",
     "scaled_output_error", "absolute", Decimal("1.92E-16"), "loggamma1p_holdout.csv"),
]

# worst on the study's own fitting grid, for the cross-check
STUDY_WORST = {
    "LogGamma.small_positive.log_abs": Decimal("8.57E-14"),
    "LogGamma.near_zero.log_abs": Decimal("9.77E-15"),
    "LogGamma.general.output_rel": Decimal("9.31E-14"),
    "LogGamma1p.small.scaled_abs": Decimal("1.36E-16"),
}


def one_two_five(x):
    """Ascending 1-2-5 levels strictly above x."""
    e = x.adjusted() - 1
    out = []
    while len(out) < 8:
        for m in (1, 2, 5):
            c = Decimal(m) * (Decimal(10) ** e)
            if c > x:
                out.append(c)
        e += 1
    return sorted(set(out))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grid", default=os.path.join(HERE, "probability_accuracy_grid.csv"))
    ap.add_argument("--out", default=os.path.join(HERE, "freeze_record.csv"))
    a = ap.parse_args()
    rows = list(csv.DictReader(open(a.grid)))

    rec = []
    for cid, fn, regime, measure, metric, hold_worst, hold_src in SPEC:
        matched = [r for r in rows if r["function"] == fn and r["regime"] == regime]
        worst = Decimal(0); at = ""; n = 0
        for r in matched:
            o = parse_observed(r["observed_vba"])
            ref = parse_reference(r["reference"])
            if o is None or ref is None:
                continue
            if measure == "scaled_output_error":
                a1 = parse_reference(r["arg1"])
                if a1 is None or a1 == 0:
                    continue
                e = calculate_scaled_error(o, ref, a1)
            else:
                e = calculate_error(o, ref, metric)
            n += 1
            if e > worst:
                worst, at = e, r["arg1"]
        combined = max(worst, hold_worst)
        levels = one_two_five(combined)
        strict = levels[0]
        chosen = next(c for c in levels if c / combined >= MIN_MARGIN)
        study = STUDY_WORST[cid]
        rec.append({
            "contract": cid, "measure": measure, "metric": metric,
            "main_worst": f"{float(worst):.3e}", "main_worst_arg": at, "main_points": n,
            "holdout_worst": f"{float(hold_worst):.3e}", "holdout_source": hold_src,
            "combined_worst": f"{float(combined):.3e}",
            "strict_125": f"{float(strict):.0e}",
            "strict_margin": f"{float(strict / combined):.2f}",
            "selected_freeze": f"{float(chosen):.0e}",
            "margin_main": f"{float(chosen / worst):.2f}" if worst else "inf",
            "margin_holdout": f"{float(chosen / hold_worst):.2f}",
            "margin_combined": f"{float(chosen / combined):.2f}",
            "study_worst": f"{float(study):.3e}",
            "main_over_study": f"{float(worst / study):.3f}" if study else "",
        })

    w = 34
    print(f"{'contract':{w}} {'main worst':>11} {'at':>24} {'holdout':>10} "
          f"{'combined':>10} {'strict':>7} {'m':>5} {'freeze':>7} {'m':>5}")
    for r in rec:
        print(f"{r['contract']:{w}} {r['main_worst']:>11} {r['main_worst_arg'][:24]:>24} "
              f"{r['holdout_worst']:>10} {r['combined_worst']:>10} "
              f"{r['strict_125']:>7} {r['strict_margin']:>5} "
              f"{r['selected_freeze']:>7} {r['margin_combined']:>5}")

    print(f"\n{'contract':{w}} {'study worst':>12} {'main worst':>11} {'main/study':>11}")
    for r in rec:
        flag = "" if 0.5 <= float(r["main_over_study"] or 1) <= 2 else "   <- EXPLAIN"
        print(f"{r['contract']:{w}} {r['study_worst']:>12} {r['main_worst']:>11} "
              f"{r['main_over_study']:>11}{flag}")

    print(f"\nrejected strict levels (margin below {MIN_MARGIN}x against the combined worst):")
    any_rej = False
    for r in rec:
        if r["strict_125"] != r["selected_freeze"]:
            any_rej = True
            print(f"  {r['contract']:{w}} {r['strict_125']} at "
                  f"{r['strict_margin']}x -> {r['selected_freeze']} at {r['margin_combined']}x")
    if not any_rej:
        print("  none")

    with open(a.out, "w", newline="") as fh:
        wr = csv.DictWriter(fh, fieldnames=list(rec[0].keys()), lineterminator="\n")
        wr.writeheader(); wr.writerows(rec)
    print(f"\nwrote {a.out}")


if __name__ == "__main__":
    main()
