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
#              expected main-grid points,
#              holdout file, holdout quantity, holdout regimes,
#              expected holdout points)
#
# Holdout worsts are RECOMPUTED from the committed holdout CSVs, never
# hard-coded. A transcribed constant makes the freeze record reproducible only
# for as long as somebody remembers to retype it; reading the file makes the
# record regenerate itself. Expected point counts are asserted so a partially
# missing evidence set fails rather than silently freezing a threshold from
# whatever happens to be present.
REGIMES_HOLDOUT = "loggamma_regimes_study/loggamma_regimes_holdout.csv"
LG1P_HOLDOUT = "loggamma1p_study/loggamma1p_holdout.csv"

SPEC = [
    ("LogGamma.small_positive.log_abs", "LogGamma", "small_positive",
     "output_error", "absolute", 39,
     REGIMES_HOLDOUT, "LogGamma", ("small_positive",), 28),
    ("LogGamma.near_zero.log_abs", "LogGamma", "near_zero",
     "output_error", "absolute", 11,
     # exact_zero is folded into near_zero: same absolute contract, and a
     # relative one is undefined at a zero.
     REGIMES_HOLDOUT, "LogGamma", ("near_zero", "exact_zero"), 16),
    ("LogGamma.general.output_rel", "LogGamma", "general",
     "output_error", "relative", 16,
     REGIMES_HOLDOUT, "LogGamma", ("general",), 10),
    ("LogGamma1p.small.scaled_abs", "LogGamma1p", "small",
     "scaled_output_error", "absolute", 32,
     LG1P_HOLDOUT, "LogGamma1p", ("small",), 32),
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


def worst_in(path, quantity, regimes, measure, metric):
    """Worst error over a study holdout file, computed the same way the gate
    computes it: same parsing, same metric arithmetic, same scaled rule."""
    full = path if os.path.isabs(path) else os.path.join(HERE, path)
    if not os.path.exists(full):
        raise SystemExit(f"holdout file not found: {full}")
    worst = Decimal(0); at = ""; n = 0
    for r in csv.DictReader(open(full)):
        if r.get("quantity") != quantity or r.get("regime") not in regimes:
            continue
        o = parse_observed(r["observed_vba"])
        ref = parse_reference(r["reference"])
        if o is None or ref is None:
            raise SystemExit(
                f"{os.path.basename(full)}: unusable row at arg1={r['arg1']!r}; "
                f"a freeze must not be derived from incomplete evidence")
        if measure == "scaled_output_error":
            a1 = parse_reference(r["arg1"])
            if a1 is None or a1 == 0:
                raise SystemExit(
                    f"{os.path.basename(full)}: scaled row with unusable arg1")
            e = calculate_scaled_error(o, ref, a1)
        else:
            e = calculate_error(o, ref, metric)
        n += 1
        if e > worst:
            worst, at = e, r["arg1"]
    return worst, at, n


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grid", default=os.path.join(HERE, "probability_accuracy_grid.csv"))
    ap.add_argument("--out", default=os.path.join(HERE, "freeze_record.csv"))
    a = ap.parse_args()
    rows = list(csv.DictReader(open(a.grid)))

    rec = []
    for (cid, fn, regime, measure, metric, want_main,
         hold_path, hold_qty, hold_regimes, want_hold) in SPEC:
        matched = [r for r in rows if r["function"] == fn and r["regime"] == regime]
        if len(matched) != want_main:
            raise SystemExit(
                f"{cid}: expected {want_main} main-grid rows, found {len(matched)}. "
                f"A freeze must not be derived from an incomplete evidence set.")
        worst = Decimal(0); at = ""; n = 0
        for r in matched:
            o = parse_observed(r["observed_vba"])
            ref = parse_reference(r["reference"])
            if o is None or ref is None:
                raise SystemExit(
                    f"{cid}: unusable row at arg1={r['arg1']!r}; a freeze must "
                    f"not be derived from incomplete evidence")
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
        if n != want_main:
            raise SystemExit(f"{cid}: scored {n} of {want_main} main-grid rows")
        hold_worst, hold_at, hold_n = worst_in(
            hold_path, hold_qty, hold_regimes, measure, metric)
        if hold_n != want_hold:
            raise SystemExit(
                f"{cid}: expected {want_hold} holdout points in "
                f"{os.path.basename(hold_path)}, found {hold_n}")
        hold_src = os.path.basename(hold_path)
        combined = max(worst, hold_worst)
        levels = one_two_five(combined)
        strict = levels[0]
        chosen = next(c for c in levels if c / combined >= MIN_MARGIN)
        study = STUDY_WORST[cid]
        rec.append({
            "contract": cid, "measure": measure, "metric": metric,
            "main_worst": f"{float(worst):.3e}", "main_worst_arg": at, "main_points": n,
            "holdout_worst": f"{float(hold_worst):.3e}",
            "holdout_worst_arg": hold_at, "holdout_points": hold_n,
            "holdout_source": hold_src,
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
