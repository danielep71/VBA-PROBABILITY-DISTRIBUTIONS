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

    if fails:
        print("FAIL: analyze_holdout metric semantics")
        for f in fails:
            print("  - " + f)
        raise SystemExit(1)
    print("PASS: metric controls abs-vs-rel; measure selects the residual and\n      scaled paths; invalid scaled arg1 is caught by preflight; an\n      unknown measure is rejected")


if __name__ == "__main__":
    main()
