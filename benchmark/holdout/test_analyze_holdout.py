"""
Regression test for analyze_holdout.worst_for evaluation semantics.

Guards the contract that `metric` (not `measure`) selects absolute-versus-relative
evaluation for ordinary contracts, while `measure` only selects the special-case
tail_probability_residual path. This test exists because an earlier analyzer keyed
abs-vs-rel off `measure`, which silently evaluated metric=absolute contracts
(e.g. PROB_LogBeta.tiny_unbalanced) as relative.

Proves:
  1. measure=output_error,       metric=absolute -> Abs(observed - reference)
  2. measure=log_absolute_error, metric=absolute -> Abs(observed - reference)
  3. any measure,                metric=relative -> Abs(observed - reference)/Abs(reference)
Run: python3 test_analyze_holdout.py   (exit 0 = pass, nonzero = fail)
"""
from decimal import Decimal
from analyze_holdout import worst_for
from _contract_eval import predicted_expected_error, validate_measure


def _row(observed, reference):
    # Only the fields worst_for touches on the ordinary (non-residual) path.
    return {"observed_vba": observed, "reference": reference,
            "arg1": "", "arg2": "", "arg3": ""}


def main():
    # reference = 10, observed = 10 + 1e-13  ->  abs error 1e-13, rel error 1e-14
    rows = [_row("10.0000000000001", "10")]
    abs_expected = Decimal("1E-13")
    rel_expected = Decimal("1E-14")
    fails = []

    # Guard: the envelope fixtures below are only meaningful if their df really
    # is beyond the current cap. Pin that to the shared predicate rather than to
    # a literal, so raising PROB_F_MAX_DF again fails here loudly instead of
    # quietly turning these into in-envelope rows.
    if not predicted_expected_error("F_InverseCumulative", "2E10", "4"):
        fails.append("envelope fixture df 2E10 is no longer beyond the F cap; "
                     "update the fixtures to a df that is")

    # 1. output_error + absolute -> absolute
    w, _, n, _, _, _, _, _ = worst_for("output_error", "absolute", rows, "Fn")
    if not (n == 1 and w == abs_expected):
        fails.append(f"(output_error, absolute): got {w} across {n}, expected {abs_expected}")

    # 2. log_absolute_error + absolute -> absolute (same result as case 1)
    w, _, n, _, _, _, _, _ = worst_for("log_absolute_error", "absolute", rows, "Fn")
    if not (n == 1 and w == abs_expected):
        fails.append(f"(log_absolute_error, absolute): got {w} across {n}, expected {abs_expected}")

    # 3. output_error + relative -> relative
    w, _, n, _, _, _, _, _ = worst_for("output_error", "relative", rows, "Fn")
    if not (n == 1 and w == rel_expected):
        fails.append(f"(output_error, relative): got {w} across {n}, expected {rel_expected}")

    # 4. cases 1 and 2 must agree: metric, not measure, decides the kind
    w1, _, _, _, _, _, _, _ = worst_for("output_error", "absolute", rows, "Fn")
    w2, _, _, _, _, _, _, _ = worst_for("log_absolute_error", "absolute", rows, "Fn")
    if w1 != w2:
        fails.append(f"metric=absolute disagreed across measures: {w1} vs {w2}")

    # 5. an unknown metric must fail loudly, not silently default
    try:
        worst_for("output_error", "", rows, "Fn")
        fails.append("empty metric did not raise")
    except ValueError:
        pass

    # 6. an envelope-reject row (F CDF/Survival/Inverse, df beyond PROB_F_MAX_DF)
    #    that returns ERROR is EXCLUDED, not scored, and not flagged as unexpected.
    #    The df here must stay beyond the CURRENT cap: it was 5E5 while the cap was
    #    1E5, and silently became an in-envelope row when the cap rose to 1E10,
    #    which is what broke this test.
    env_error = [{"function": "F_InverseCumulative", "observed_vba": "ERROR", "reference": "1",
                  "arg1": "0.5", "arg2": "2E10", "arg3": "4"}]
    w, _, n, n_missing, n_error, n_violation, n_invalid, _ = worst_for("output_error", "relative", env_error, "F_InverseCumulative")
    if not (n == 0 and n_missing == 0 and n_error == 0 and n_violation == 0):
        fails.append(f"envelope ERROR row not excluded cleanly: n={n} missing={n_missing} error={n_error} viol={n_violation}")

    # 7. an in-envelope ERROR row is UNEXPECTED and reported so the caller can block.
    in_env_error = [{"function": "F_InverseCumulative", "observed_vba": "ERROR", "reference": "1",
                     "arg1": "0.5", "arg2": "50", "arg3": "4"}]
    w, _, n, n_missing, n_error, n_violation, n_invalid, _ = worst_for("output_error", "relative", in_env_error, "F_InverseCumulative")
    if not (n == 0 and n_error == 1):
        fails.append(f"in-envelope ERROR not flagged: n={n} error={n_error}")

    # 8. Direction 2: an envelope-reject row that returns a VALUE (should have been
    #    #NUM!) is a violation, reported so the caller can block.
    env_value = [{"function": "F_InverseCumulative", "observed_vba": "1.23", "reference": "1",
                  "arg1": "0.5", "arg2": "2E10", "arg3": "4"}]
    w, _, n, n_missing, n_error, n_violation, n_invalid, _ = worst_for("output_error", "relative", env_value, "F_InverseCumulative")
    if not (n == 0 and n_violation == 1):
        fails.append(f"envelope-reject VALUE not flagged as violation: n={n} viol={n_violation}")

    # --- scaled_output_error: an invalid arg1 must be caught by preflight,
    # not silently skipped by the measurer. Without this a ten-row contract
    # could PASS having scored nine.
    scaled_rows = [
        {"function": "LogGamma1p", "arg1": "1e-13", "arg2": "", "arg3": "", "arg4": "",
         "reference": "-5.7721566490145e-14", "observed_vba": "-5.7721566490145E-014;0E+000",
         "expected_error": "", "regime": "small"},
        {"function": "LogGamma1p", "arg1": "0", "arg2": "", "arg3": "", "arg4": "",
         "reference": "0", "observed_vba": "0E+000;0E+000",
         "expected_error": "", "regime": "small"},
    ]
    _, _, n, _, _, _, n_invalid, n_skipped = worst_for(
        "scaled_output_error", "absolute", scaled_rows, "LogGamma1p")
    if n_invalid != 1:
        fails.append(f"scaled arg1=0 must be invalid, got n_invalid={n_invalid}")
    if n != 1:
        fails.append(f"only the valid scaled row should be scored, got n={n}")
    if n_skipped != 0:
        fails.append(f"no row may be silently skipped, got n_skipped={n_skipped}")

    blank = [dict(scaled_rows[1], arg1="")]
    _, _, _, _, _, _, n_invalid, _ = worst_for(
        "scaled_output_error", "absolute", blank, "LogGamma1p")
    if n_invalid != 1:
        fails.append(f"scaled blank arg1 must be invalid, got {n_invalid}")

    # --- an unknown measure must fail loudly here too, not be scored as
    # ordinary error. The gate validates it; without the same call the
    # whitelist would protect only one of the two evaluation arms.
    try:
        validate_measure("scaled_abs")
        fails.append("holdout must reject an unknown measure")
    except ValueError:
        pass

    # --- the tail dispatch must be fail-closed on this arm too --------------
    # The analyzer previously used the same binary branch as the gate:
    #   ibeta(...) if fn == "Beta_InverseCumulative" else f_cdf(...)
    # so an unsupported inverse would have been scored with the F CDF here as
    # well. Both arms now share _contract_eval.TAIL_SUPPORTED, and these
    # fixtures exist so the two cannot drift to different supported sets.
    from _contract_eval import (TAIL_SUPPORTED, tail_cdf_name,
                                UnsupportedTailFunction, dispositions,
                                row_validity, normalize_tail_residual)
    import analyze_holdout as AH
    from analyze_holdout import mp

    if set(AH._TAIL_CDFS) != {spec["cdf"] for spec in TAIL_SUPPORTED.values()}:
        fails.append("holdout CDF table does not match the shared registry")
    if tail_cdf_name("Beta_InverseCumulative") != "ibeta":
        fails.append("holdout: Beta must route to ibeta")
    if tail_cdf_name("F_InverseCumulative") != "f_cdf":
        fails.append("holdout: F must route to f_cdf")
    if AH._TAIL_CDFS.get("ibeta") is not AH.ibeta:
        fails.append("holdout: ibeta name is not bound to the ibeta callable")
    if AH._TAIL_CDFS.get("f_cdf") is not AH.f_cdf:
        fails.append("holdout: f_cdf name is not bound to the f_cdf callable")

    for fn_bad in ("ChiSquare_InverseCumulative", "StudentT_InverseCumulative",
                   "Gamma_InverseCumulative", ""):
        try:
            tail_cdf_name(fn_bad)
            fails.append(f"holdout: unsupported tail function must raise: {fn_bad!r}")
        except UnsupportedTailFunction:
            pass

    # Plausible three-argument data for an unsupported function must be blocked
    # before scoring, so it can never reach the F evaluator.
    tail_row = {"function": "F_InverseCumulative", "arg1": "0.9", "arg2": "5",
                "arg3": "7", "arg4": "", "reference": "",
                "observed_vba": "3.9715343910;0E+000", "expected_error": ""}
    d_ok = dispositions([tail_row], "tail_probability_residual")
    if len(d_ok.to_measure) != 1:
        fails.append("holdout: a supported tail row must reach the measurer")
    for fn_bad in ("ChiSquare_InverseCumulative", "StudentT_InverseCumulative"):
        d_bad = dispositions([dict(tail_row, function=fn_bad)],
                             "tail_probability_residual")
        if d_bad.to_measure or d_bad.n_invalid != 1:
            fails.append(f"holdout: unsupported tail row must be blocked: {fn_bad}")

    # Missing or malformed required arguments still fail on this arm.
    for k in ("arg1", "arg2", "arg3"):
        if row_validity(dict(tail_row, **{k: ""}), "tail_probability_residual") is None:
            fails.append(f"holdout: blank {k} must be invalid")
        if row_validity(dict(tail_row, **{k: "abc"}), "tail_probability_residual") is None:
            fails.append(f"holdout: unparseable {k} must be invalid")

    # --- the HOLDOUT EVALUATOR must actually use the registry too -----------
    # Same gap as the main arm: the registry assertions above do not prove that
    # worst_for() consults it. Reintroducing the binary branch in
    # analyze_holdout.py alone left every committed fixture green. Assert the
    # value worst_for RETURNS, which is only reproducible if it dispatched to
    # the correct CDF; ibeta and f_cdf differ widely at these arguments.
    def tail_row(fn, p_, a2, a3, x):
        return {"function": fn, "arg1": p_, "arg2": a2, "arg3": a3, "arg4": "",
                "reference": "", "observed_vba": x, "expected_error": ""}

    bx, ba, bb, bp = "0.3", "2", "5", "0.4"
    got_b = AH.worst_for("tail_probability_residual", "relative",
                         [tail_row("Beta_InverseCumulative", bp, ba, bb, bx)],
                         "Beta_InverseCumulative")[0]
    want_b = normalize_tail_residual(AH.ibeta(mp.mpf(bx), mp.mpf(ba), mp.mpf(bb)),
                                     mp.mpf(bp))
    wrong_b = normalize_tail_residual(AH.f_cdf(mp.mpf(bx), mp.mpf(ba), mp.mpf(bb)),
                                      mp.mpf(bp))
    if got_b != want_b:
        fails.append("holdout evaluator did not dispatch Beta to ibeta")
    if got_b == wrong_b:
        fails.append("holdout evaluator dispatched Beta to f_cdf")

    fx, f1, f2, fp = "0.5", "5", "7", "0.4"
    got_f = AH.worst_for("tail_probability_residual", "relative",
                         [tail_row("F_InverseCumulative", fp, f1, f2, fx)],
                         "F_InverseCumulative")[0]
    want_f = normalize_tail_residual(AH.f_cdf(mp.mpf(fx), mp.mpf(f1), mp.mpf(f2)),
                                     mp.mpf(fp))
    wrong_f = normalize_tail_residual(AH.ibeta(mp.mpf(fx), mp.mpf(f1), mp.mpf(f2)),
                                      mp.mpf(fp))
    if got_f != want_f:
        fails.append("holdout evaluator did not dispatch F to f_cdf")
    if got_f == wrong_f:
        fails.append("holdout evaluator dispatched F to ibeta")

    # An unsupported function must never be scored here either. Note the
    # architectures differ: worst_for() calls dispositions() itself, so a
    # function-aware row_validity blocks the row BEFORE the dispatch and it
    # returns None rather than raising. compute_errors.tail_residual receives
    # already-filtered rows, so there the dispatch itself must raise. Both are
    # fail-closed; only the layer that stops them differs. Assert the property
    # that matters - the row is not scored - rather than the mechanism.
    for fn_bad in ("ChiSquare_InverseCumulative", "StudentT_InverseCumulative"):
        res = AH.worst_for("tail_probability_residual", "relative",
                           [tail_row(fn_bad, "0.9", "1000000", "1", "1039569.3")],
                           fn_bad)
        worst_bad, _, n_bad = res[0], res[1], res[2]
        if worst_bad is not None or n_bad != 0:
            fails.append(f"holdout evaluator scored an unsupported function: {fn_bad}")
    # And the dispatch SEAM itself must refuse, independently of preflight.
    # Without this, a fall-through reintroduced in analyze_holdout.py alone is
    # caught by no fixture: worst_for() blocks unsupported rows before the
    # dispatch, and for Beta and F the binary branch returns identical values.
    # Testing forward_cdf directly makes that regression visible.
    if AH.forward_cdf("Beta_InverseCumulative", mp.mpf(bx), mp.mpf(ba),
                      mp.mpf(bb)) != AH.ibeta(mp.mpf(bx), mp.mpf(ba), mp.mpf(bb)):
        fails.append("forward_cdf did not dispatch Beta to ibeta")
    if AH.forward_cdf("F_InverseCumulative", mp.mpf(fx), mp.mpf(f1),
                      mp.mpf(f2)) != AH.f_cdf(mp.mpf(fx), mp.mpf(f1), mp.mpf(f2)):
        fails.append("forward_cdf did not dispatch F to f_cdf")
    for fn_bad in ("ChiSquare_InverseCumulative", "StudentT_InverseCumulative",
                   "Gamma_InverseCumulative", ""):
        try:
            AH.forward_cdf(fn_bad, mp.mpf("0.5"), mp.mpf(5), mp.mpf(7))
            fails.append(f"forward_cdf must refuse {fn_bad!r}")
        except UnsupportedTailFunction:
            pass

    if fails:
        print("FAIL: analyze_holdout metric semantics")
        for f in fails:
            print("  - " + f)
        raise SystemExit(1)
    print("PASS: metric controls abs-vs-rel; measure selects the residual and\n"
          "      scaled paths; invalid scaled arg1 is caught by preflight; an\n"
          "      unknown measure is rejected; tail dispatch is fail-closed,\n"
          "      shares the gate's registry, and is asserted through the\n"
          "      evaluator itself")


if __name__ == "__main__":
    main()
