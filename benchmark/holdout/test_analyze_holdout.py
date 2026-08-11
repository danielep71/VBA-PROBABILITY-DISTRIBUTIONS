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
from _contract_eval import predicted_expected_error


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
    w, _, n, _, _, _, _ = worst_for("output_error", "absolute", rows, "Fn")
    if not (n == 1 and w == abs_expected):
        fails.append(f"(output_error, absolute): got {w} across {n}, expected {abs_expected}")

    # 2. log_absolute_error + absolute -> absolute (same result as case 1)
    w, _, n, _, _, _, _ = worst_for("log_absolute_error", "absolute", rows, "Fn")
    if not (n == 1 and w == abs_expected):
        fails.append(f"(log_absolute_error, absolute): got {w} across {n}, expected {abs_expected}")

    # 3. output_error + relative -> relative
    w, _, n, _, _, _, _ = worst_for("output_error", "relative", rows, "Fn")
    if not (n == 1 and w == rel_expected):
        fails.append(f"(output_error, relative): got {w} across {n}, expected {rel_expected}")

    # 4. cases 1 and 2 must agree: metric, not measure, decides the kind
    w1, _, _, _, _, _, _ = worst_for("output_error", "absolute", rows, "Fn")
    w2, _, _, _, _, _, _ = worst_for("log_absolute_error", "absolute", rows, "Fn")
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
    w, _, n, n_missing, n_error, n_violation, n_invalid = worst_for("output_error", "relative", env_error, "F_InverseCumulative")
    if not (n == 0 and n_missing == 0 and n_error == 0 and n_violation == 0):
        fails.append(f"envelope ERROR row not excluded cleanly: n={n} missing={n_missing} error={n_error} viol={n_violation}")

    # 7. an in-envelope ERROR row is UNEXPECTED and reported so the caller can block.
    in_env_error = [{"function": "F_InverseCumulative", "observed_vba": "ERROR", "reference": "1",
                     "arg1": "0.5", "arg2": "50", "arg3": "4"}]
    w, _, n, n_missing, n_error, n_violation, n_invalid = worst_for("output_error", "relative", in_env_error, "F_InverseCumulative")
    if not (n == 0 and n_error == 1):
        fails.append(f"in-envelope ERROR not flagged: n={n} error={n_error}")

    # 8. Direction 2: an envelope-reject row that returns a VALUE (should have been
    #    #NUM!) is a violation, reported so the caller can block.
    env_value = [{"function": "F_InverseCumulative", "observed_vba": "1.23", "reference": "1",
                  "arg1": "0.5", "arg2": "2E10", "arg3": "4"}]
    w, _, n, n_missing, n_error, n_violation, n_invalid = worst_for("output_error", "relative", env_value, "F_InverseCumulative")
    if not (n == 0 and n_violation == 1):
        fails.append(f"envelope-reject VALUE not flagged as violation: n={n} viol={n_violation}")

    if fails:
        print("FAIL: analyze_holdout metric semantics")
        for f in fails:
            print("  - " + f)
        raise SystemExit(1)
    print("PASS: metric controls abs-vs-rel; measure only selects the residual path (5/5 checks)")


if __name__ == "__main__":
    main()
