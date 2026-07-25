"""
Compute accuracy verdicts from the regime-aware contract and act as the
numerical release gate.

Contract schema (accuracy_contracts.csv):
  contract_id,function,regime,measure,metric,threshold,domain,provenance,status,evidence,notes

Grid schema (probability_accuracy_grid.csv), positional up to observed_vba:
  function,vba_kernel,claim,metric,arg1,arg2,arg3,reference,observed_vba,regime,evidence_set

Join: grid (function, regime) -> contract (function, regime). One observation may
feed several contracts (e.g. an inverse quantile supports both a quantile_error
and a tail_probability_residual contract). Verdicts are grouped by contract_id.

Measures:
  output_error / quantile_error  -> error of observed vs reference (rel or abs)
  log_absolute_error             -> absolute error
  tail_probability_residual      -> requires the forward CDF at the observed
                                    quantile; verified in its study directory,
                                    surfaced here as CHARACTERIZATION ONLY unless
                                    integrated into the grid.
"""
import argparse, csv, os, sys
from collections import defaultdict
from decimal import Decimal, getcontext, InvalidOperation
getcontext().prec = 50

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)  # single-sourced benchmark/ helpers
from _contract_eval import (parse_observed, parse_reference, calculate_error,
                            normalize_tail_residual, dispositions, expected_error_drift)

_IBETA_IMPORT_ERROR = None
try:
    from _ibeta import ibeta as _ibeta_cdf, f_cdf as _f_cdf
    import mpmath as _mp
    _mp.mp.dps = 50
    _HAVE_IBETA = True
except (ImportError, ModuleNotFoundError, SyntaxError) as _e:
    # Missing, corrupt, or incompatible reference helper. Retain the exact reason;
    # an ACTIVE contract that needs this evaluator must then BLOCK, never silently
    # downgrade to a non-blocking CHARACTERIZATION ONLY verdict.
    _HAVE_IBETA = False
    _IBETA_IMPORT_ERROR = f"{type(_e).__name__}: {_e}"


def load_contracts(path=None):
    if path is None:
        path = os.path.join(HERE, "accuracy_contracts.csv")
    with open(path, newline="") as f:
        return list(csv.DictReader(f))


def measure_error(rows, metric):
    """Worst error (Decimal) over rows for output_error/quantile_error/log_absolute_error."""
    worst = Decimal(-1); worst_at = ""; n = 0
    for r in rows:
        o = parse_observed(r["observed_vba"])
        if o is None:
            continue
        ref = parse_reference(r["reference"])
        if ref is None:
            continue
        e = calculate_error(o, ref, metric)
        n += 1
        if e > worst:
            worst = e; worst_at = ", ".join(x for x in (r["arg1"], r["arg2"], r["arg3"]) if x)
    return (worst, worst_at, n) if n else (None, "", 0)


def tail_residual(rows, fn):
    if not _HAVE_IBETA:
        return None, "", 0
    worst = Decimal(-1); worst_at = ""; cnt = 0
    for r in rows:
        xo = parse_observed(r["observed_vba"])
        if xo is None:
            continue
        p = _mp.mpf(r["arg1"]); a2v = _mp.mpf(r["arg2"]); a3v = _mp.mpf(r["arg3"])
        xv = _mp.mpf(str(xo))
        recovered = _ibeta_cdf(xv, a2v, a3v) if fn == "Beta_InverseCumulative" else _f_cdf(xv, a2v, a3v)
        e = normalize_tail_residual(recovered, p)
        cnt += 1
        if e > worst:
            worst = e; worst_at = ", ".join(z for z in (r["arg1"], r["arg2"], r["arg3"]) if z)
    return (worst, worst_at, cnt) if cnt else (None, "", 0)


def _force_utf8_stdout():
    # Windows consoles default to cp1252 and cannot print the verdict icons.
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def main():
    _force_utf8_stdout()
    ap = argparse.ArgumentParser(description="Regime-aware accuracy verdicts + release gate.")
    ap.add_argument("--grid", default=os.path.join(HERE, "probability_accuracy_grid.csv"))
    ap.add_argument("--out", default=os.path.join(HERE, "accuracy_summary.md"))
    ap.add_argument("--allow-known-limitations", action="store_true",
                    help="Development mode: KNOWN LIMITATION does not fail the gate.")
    args = ap.parse_args()

    contracts = load_contracts()
    rows = list(csv.DictReader(open(args.grid)))

    # The stored expected_error column must match the envelope predicate exactly,
    # so a hand-edited or stale marker cannot silently reclassify a row. A
    # mismatch is a hard gate failure, not a warning.
    drift = expected_error_drift(rows)
    if drift:
        print(f"  gate FAILED (exit 1): {len(drift)} grid row(s) have an expected_error "
              f"marker inconsistent with the envelope predicate, e.g. "
              f"{drift[0].get('function','?')} df=({drift[0].get('arg2','')},{drift[0].get('arg3','')}).")
        sys.exit(1)

    # index grid rows by (function, regime)
    grid_by = defaultdict(list)
    for r in rows:
        grid_by[(r["function"], r.get("regime", "all") or "all")].append(r)

    lines = ["# Accuracy summary", "",
             "Generated by `compute_errors.py` from the regime-aware contract.", "",
             "| Contract | Measure | Metric | Threshold | Worst error | Points | Verdict |",
             "|---|---|---|---|---:|---:|---|"]

    n_fail = n_known = n_char = n_pending = 0
    unevaluated = []   # (contract_id, reason) for active contracts that could not be evaluated

    for c in sorted(contracts, key=lambda c: c["contract_id"]):
        cid = c["contract_id"]; fn = c["function"]; regime = c["regime"]
        measure = c["measure"]; metric = c["metric"]; status = c["status"]
        try:
            threshold = Decimal(c["threshold"]) if c["threshold"].strip() else None
        except InvalidOperation:
            threshold = None

        matched = grid_by.get((fn, regime), [])

        if measure == "tail_probability_residual":
            mt = grid_by.get((fn, regime), [])
            # CHARACTERIZATION ONLY is reserved for contracts EXPLICITLY marked so.
            if status == "characterization_only":
                n_char += 1
                lines.append(f"| {cid} | {measure} | {metric} | {c['threshold']} | — | "
                             f"study | \U0001f9ea CHARACTERIZATION ONLY |")
                continue
            # An ACTIVE tail contract that cannot be evaluated must BLOCK.
            if not _HAVE_IBETA:
                n_pending += 1
                unevaluated.append((cid, "reference helper unavailable: "
                                    + (_IBETA_IMPORT_ERROR or "unknown import failure")))
                lines.append(f"| {cid} | {measure} | {metric} | {c['threshold']} | — | "
                             f"evaluator | \u23f3 PENDING - evaluator unavailable |")
                continue
            if not mt:
                n_pending += 1
                unevaluated.append((cid, "no matching observations in the grid"))
                lines.append(f"| {cid} | {measure} | {metric} | {c['threshold']} | — | "
                             f"0 | \u23f3 PENDING - no observations |")
                continue
            to_measure, n_expected, n_missing, n_error, n_violation = dispositions(mt)
            in_env = len(mt) - n_expected
            if status == "active" and (n_missing or n_error or n_violation):
                n_pending += 1
                bits = []
                if n_missing:
                    bits.append(f"{n_missing} unobserved")
                if n_error:
                    bits.append(f"{n_error} unexpected ERROR in-envelope")
                if n_violation:
                    bits.append(f"{n_violation} envelope-reject row(s) that did not return #NUM!")
                unevaluated.append((cid, "; ".join(bits) + "; strict mode requires full "
                                    "in-envelope observation"))
                lines.append(f"| {cid} | {measure} | {metric} | {c['threshold']} | — | "
                             f"{len(to_measure)}/{in_env} | \u23f3 PENDING - incomplete evidence |")
                continue
            try:
                worst, at, nn = tail_residual(to_measure, fn)
            except Exception as _te:
                n_pending += 1
                unevaluated.append((cid, f"evaluator error: {type(_te).__name__}: {_te}"))
                lines.append(f"| {cid} | {measure} | {metric} | {c['threshold']} | — | "
                             f"evaluator | \u23f3 PENDING - evaluator error |")
                continue
            if worst is None:
                n_pending += 1
                unevaluated.append((cid, "no usable in-envelope observations"))
                lines.append(f"| {cid} | {measure} | {metric} | {c['threshold']} | — | "
                             f"0/{in_env} | \u23f3 PENDING |")
                continue
            ok = threshold is not None and worst <= threshold
            if ok:
                verdict = "\u2705 PASS"
            else:
                n_fail += 1; verdict = "\u274c FAIL"
            lines.append(f"| {cid} | {measure} | {metric} | {c['threshold']} | "
                         f"{float(worst):.2e} | {nn}/{in_env} | {verdict} |")
            continue

        if not matched:
            n_pending += 1
            lines.append(f"| {cid} | {measure} | {metric} | {c['threshold']} | — | "
                         f"0 | \u23f3 PENDING |")
            continue

        # An ACTIVE contract must have every IN-ENVELOPE matched row observed with
        # a real value. A blank row (unobserved) or an unexpected ERROR (a kernel
        # failure inside the accuracy envelope) BLOCKS as PENDING rather than being
        # silently dropped. Rows in the F envelope-reject region (expected_error)
        # carry no accuracy claim and are excluded from scoring entirely.
        to_measure, n_expected, n_missing, n_error, n_violation = dispositions(matched)
        in_env = len(matched) - n_expected

        if status == "active" and (n_missing or n_error or n_violation):
            n_pending += 1
            bits = []
            if n_missing:
                bits.append(f"{n_missing} unobserved")
            if n_error:
                bits.append(f"{n_error} unexpected ERROR in-envelope")
            if n_violation:
                bits.append(f"{n_violation} envelope-reject row(s) that did not return #NUM!")
            unevaluated.append((cid, "; ".join(bits) + "; strict mode requires full "
                                "in-envelope observation"))
            lines.append(f"| {cid} | {measure} | {metric} | {c['threshold']} | — | "
                         f"{len(to_measure)}/{in_env} | "
                         f"\u23f3 PENDING - incomplete evidence |")
            continue

        worst, at, n = measure_error(to_measure, metric)
        if worst is None:
            n_pending += 1
            lines.append(f"| {cid} | {measure} | {metric} | {c['threshold']} | — | "
                         f"0/{in_env} | \u23f3 PENDING |")
            continue

        ok = threshold is not None and worst <= threshold
        if status == "known_limitation":
            n_known += 1; verdict = "\U0001f537 KNOWN LIMITATION"
        elif status == "characterization_only":
            n_char += 1; verdict = "\U0001f9ea CHARACTERIZATION ONLY"
        elif ok:
            verdict = "\u2705 PASS"
        else:
            n_fail += 1; verdict = "\u274c FAIL"

        lines.append(f"| {cid} | {measure} | {metric} | {c['threshold']} | "
                     f"{float(worst):.2e} | {n}/{in_env} | {verdict} |")

    # The verdict tally counts CONTRACT states. A reader can otherwise take
    # "KNOWN LIMITATION: 0" to mean the library has no numerical boundaries,
    # which the separate register contradicts. Derive the cross-reference from
    # numerical_limitations.csv so the two can never drift apart.
    reg_path = os.path.join(HERE, "numerical_limitations.csv")
    reg = []
    if os.path.exists(reg_path):
        with open(reg_path, newline="", encoding="utf-8") as f:
            reg = [(r.get("limitation_id", ""), r.get("status", "")) for r in csv.DictReader(f)]

    if reg:
        listed = "; ".join(f"`{lid}` ({st})" for lid, st in reg)
        register_note = (
            f"> **Register cross-reference** — the tally above counts *contract states*. "
            f"`KNOWN LIMITATION: {n_known}` means no contract is currently held open as a "
            f"documented defect; it is **not** a claim that the library has no numerical "
            f"boundaries. Those live in `numerical_limitations.csv`, which currently holds "
            f"{len(reg)} entr{'y' if len(reg) == 1 else 'ies'}: {listed}. Characterized "
            f"boundaries are carried by their own regime contracts and study directories, "
            f"so they appear above as PASS against the honest threshold rather than as a defect.")
    else:
        register_note = ("> **Register cross-reference** — `numerical_limitations.csv` is absent "
                         "or empty; the tally above counts contract states only.")

    lines += ["",
              f"> **Verdict tally** — FAIL: {n_fail}, KNOWN LIMITATION: {n_known}, "
              f"CHARACTERIZATION ONLY: {n_char}, PENDING: {n_pending}.",
              "",
              "> States: **PASS** meets the contract; **FAIL** exceeds it; **KNOWN LIMITATION** "
              "is a documented defect; **CHARACTERIZATION ONLY** is measured but not held to a "
              "pass/fail claim (or verified in a study directory); **PENDING** is not yet "
              "measured in the main grid. Errors are Decimal from the two-part hi;lo export.",
              "",
              register_note]

    with open(args.out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    mode = "development (known limitations allowed)" if args.allow_known_limitations else "strict"
    print(f"{os.path.basename(args.out)}: FAIL={n_fail} KNOWN_LIMITATION={n_known} "
          f"CHARACTERIZATION_ONLY={n_char} PENDING={n_pending} [gate: {mode}]")
    if _IBETA_IMPORT_ERROR:
        print(f"  WARNING: reference helper failed to load ({_IBETA_IMPORT_ERROR}); "
              f"tail-residual contracts cannot be evaluated.")
    for cid, reason in unevaluated:
        print(f"  UNEVALUATED active contract: {cid} - {reason}")

    fail_block = n_fail + (0 if args.allow_known_limitations else n_known)
    if fail_block:
        print(f"  gate FAILED (exit 1): {fail_block} blocking item(s).")
        sys.exit(1)
    if n_pending:
        print(f"  gate INCOMPLETE (exit 2): {n_pending} contract(s) unevaluated "
              f"(not measured, or evaluator unavailable). The strict gate never "
              f"passes with an active contract unevaluated.")
        sys.exit(2)
    print("  gate passed (exit 0).")
    sys.exit(0)


if __name__ == "__main__":
    main()
