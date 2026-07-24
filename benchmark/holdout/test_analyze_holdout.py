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

    # 1. output_error + absolute -> absolute
    w, _, n = worst_for("output_error", "absolute", rows, "Fn")
    if not (n == 1 and w == abs_expected):
        fails.append(f"(output_error, absolute): got {w} across {n}, expected {abs_expected}")

    # 2. log_absolute_error + absolute -> absolute (same result as case 1)
    w, _, n = worst_for("log_absolute_error", "absolute", rows, "Fn")
    if not (n == 1 and w == abs_expected):
        fails.append(f"(log_absolute_error, absolute): got {w} across {n}, expected {abs_expected}")

    # 3. output_error + relative -> relative
    w, _, n = worst_for("output_error", "relative", rows, "Fn")
    if not (n == 1 and w == rel_expected):
        fails.append(f"(output_error, relative): got {w} across {n}, expected {rel_expected}")

    # 4. cases 1 and 2 must agree: metric, not measure, decides the kind
    w1, _, _ = worst_for("output_error", "absolute", rows, "Fn")
    w2, _, _ = worst_for("log_absolute_error", "absolute", rows, "Fn")
    if w1 != w2:
        fails.append(f"metric=absolute disagreed across measures: {w1} vs {w2}")

    # 5. an unknown metric must fail loudly, not silently default
    try:
        worst_for("output_error", "", rows, "Fn")
        fails.append("empty metric did not raise")
    except ValueError:
        pass

    if fails:
        print("FAIL: analyze_holdout metric semantics")
        for f in fails:
            print("  - " + f)
        raise SystemExit(1)
    print("PASS: metric controls abs-vs-rel; measure only selects the residual path (5/5 checks)")


if __name__ == "__main__":
    main()
