"""
Regression test for the single-sourced contract-evaluation primitives.

Locks the arithmetic that compute_errors.py (gate) and analyze_holdout.py
(holdout) both import, so the two can never diverge again. Proves:
  * metric (not measure) chooses absolute vs relative;
  * a zero reference is handled as 0 (when observed is 0) or Infinity;
  * observation/reference/threshold parsing and evidence classification.
Run: python3 test_contract_eval.py   (exit 0 = pass, nonzero = fail)
"""
from decimal import Decimal
from _contract_eval import (calculate_scaled_error, validate_scaled_metric,
                            validate_measure, KNOWN_MEASURES,
                            
    parse_observed, parse_reference, parse_threshold, validate_metric,
    calculate_error, normalize_tail_residual, observation_state,
    evidence_gaps, OK, MISSING, ERROR,
    predicted_expected_error, classify_row, dispositions, expected_error_drift,
    row_expected_error, MEASURE, EXCLUDE_EXPECTED, BLOCK_MISSING, BLOCK_ERROR,
    BLOCK_EXPECTED, BLOCK_INVALID, row_validity, F_MAX_DF,
)

fails = []
def check(cond, msg):
    if not cond:
        fails.append(msg)

# --- calculate_error: metric selects the arithmetic -------------------------
o = Decimal("10.0000000000001"); ref = Decimal("10")   # abs 1e-13, rel 1e-14
check(calculate_error(o, ref, "absolute") == Decimal("1E-13"), "abs metric")
check(calculate_error(o, ref, "relative") == Decimal("1E-14"), "rel metric")
# same observed/reference, metric alone flips the result kind
check(calculate_error(o, ref, "absolute") != calculate_error(o, ref, "relative"),
      "metric must change the result")

# --- calculate_error: zero reference ---------------------------------------
check(calculate_error(Decimal(0), Decimal(0), "relative") == Decimal(0),
      "0/0 relative -> 0")
check(calculate_error(Decimal("1E-20"), Decimal(0), "relative") == Decimal("Infinity"),
      "nonzero-over-zero relative -> Infinity")
check(calculate_error(Decimal("1E-20"), Decimal(0), "absolute") == Decimal("1E-20"),
      "absolute is defined at zero reference")

# --- calculate_error: unknown metric must raise ----------------------------
for bad in ("", "rel", "ABS", None):
    try:
        calculate_error(o, ref, bad)
        fails.append(f"metric {bad!r} did not raise")
    except ValueError:
        pass

# --- validate_metric normalises case/whitespace ----------------------------
check(validate_metric(" Absolute ") == "absolute", "validate_metric normalise")

# --- parse_observed: two-part hi;lo sum, blanks, ERROR ---------------------
check(parse_observed("1.5;0.25") == Decimal("1.75"), "hi;lo sum")
check(parse_observed(" 3 ") == Decimal("3"), "single value strip")
check(parse_observed("") is None and parse_observed("   ") is None, "blank -> None")
check(parse_observed("ERROR") is None and parse_observed("error") is None, "ERROR -> None")

# --- parse_reference / parse_threshold -------------------------------------
check(parse_reference("2.5") == Decimal("2.5"), "reference parse")
check(parse_reference("not-a-number") is None, "malformed reference -> None")
check(parse_threshold("5E-14") == Decimal("5E-14"), "threshold parse")
check(parse_threshold("") is None and parse_threshold("  ") is None, "blank threshold -> None")

# --- observation_state / evidence_gaps -------------------------------------
check(observation_state("1;2") == OK, "state OK")
check(observation_state("") == MISSING, "state MISSING")
check(observation_state("ERROR") == ERROR, "state ERROR")
rows = [{"observed_vba": "1;0"}, {"observed_vba": ""}, {"observed_vba": "ERROR"},
        {"observed_vba": "ERROR"}]
check(evidence_gaps(rows) == (1, 2), "evidence_gaps counts (missing, error)")

# --- normalize_tail_residual: |recovered - p| / min(p, 1-p) -----------------
# recovered 0.2000000000000001, p 0.2 -> ~1e-16 / 0.2 = ~5e-16
r = normalize_tail_residual(Decimal("0.2000000000000001"), Decimal("0.2"))
check(Decimal("4E-16") < r < Decimal("6E-16"), "tail residual normalisation")

# --- predicted_expected_error: F envelope-reject region ---------------------
check(predicted_expected_error("F_Cumulative", "1.5", "2E10") is True, "F_Cumulative df2>max -> expected")
check(predicted_expected_error("F_InverseCumulative", "2E10", "4") is True, "F_Inverse df1>max -> expected")
check(predicted_expected_error("F_Survival", "50", "50") is False, "F in-envelope -> not expected")
check(predicted_expected_error("F_Density", "1000000", "3") is False, "F_Density is NOT enveloped")

# The inverses share the forward caps: benchmark/inverse_probe measured the
# inverse kernels clean to df 2E10 with no refusals, including the unbalanced
# ratios that had been assumed to fail. A df inside the forward cap must
# therefore be in-envelope for the inverse too.
check(predicted_expected_error("F_Cumulative", "1E6", "1E6") is False, "F_Cumulative 1E6 inside cap")
check(predicted_expected_error("F_InverseCumulative", "1E6", "3") is False, "F_Inverse 1E6 inside cap (shares the forward envelope)")
check(predicted_expected_error("StudentT_Cumulative", "1E7", "") is False, "t cdf 1E7 inside cap")
check(predicted_expected_error("StudentT_InverseCumulative", "1E7", "") is False, "t inverse 1E7 inside cap")
check(predicted_expected_error("ChiSquare_Survival", "1E7", "") is False, "chi survival 1E7 inside cap")
check(predicted_expected_error("ChiSquare_InverseCumulative", "1E7", "") is False, "chi inverse 1E7 inside cap")
check(predicted_expected_error("ChiSquare_InverseCumulative", "2E8", "") is True, "chi inverse beyond the shared cap -> expected")
check(predicted_expected_error("Beta_Cumulative", "1e9", "1e9") is False, "non-F -> not expected")
check(predicted_expected_error("F_Cumulative", str(F_MAX_DF), "3") is False, "exactly at F_MAX_DF -> not expected")

# --- classify_row (Direction 2: expected must be ERROR) ---------------------
check(classify_row("ERROR", True) == EXCLUDE_EXPECTED, "expected + ERROR -> EXCLUDE_EXPECTED")
check(classify_row("1;2", True) == BLOCK_EXPECTED, "expected + value -> BLOCK_EXPECTED")
check(classify_row("", True) == BLOCK_EXPECTED, "expected + blank -> BLOCK_EXPECTED")
check(classify_row("", False) == BLOCK_MISSING, "blank in-envelope -> BLOCK_MISSING")
check(classify_row("ERROR", False) == BLOCK_ERROR, "ERROR in-envelope -> BLOCK_ERROR")
check(classify_row("1;2", False) == MEASURE, "value in-envelope -> MEASURE")

# --- dispositions (namedtuple result, measure-aware validity) ---------------
mixed = [
    {"function": "F_Cumulative", "observed_vba": "0.5", "reference": "0.5", "arg2": "3", "arg3": "4"},  # measure
    {"function": "F_Cumulative", "observed_vba": "ERROR", "reference": "0.5", "arg2": "3", "arg3": "2E10"},  # expected
    {"function": "F_Cumulative", "observed_vba": "", "reference": "0.5", "arg2": "3", "arg3": "4"},     # missing
    {"function": "F_Cumulative", "observed_vba": "ERROR", "reference": "0.5", "arg2": "3", "arg3": "4"},  # unexpected error
    {"function": "F_Cumulative", "observed_vba": "0.9", "reference": "0.5", "arg2": "3", "arg3": "2E10"},  # violation
]
d = dispositions(mixed, "output_error")
check((len(d.to_measure), d.n_expected, d.n_missing, d.n_error, d.n_violation, d.n_invalid)
      == (1, 1, 1, 1, 1, 0), "dispositions partition")
check(len(d.reasons) == 3, "dispositions reasons name each blocking row")

# --- row_validity + the 8 required release-gate completeness cases -----------
NT = "output_error"; TL = "tail_probability_residual"
def one(measure, **row):
    row.setdefault("function", "F_Cumulative")
    return dispositions([row], measure)

# 1. blank ordinary observation -> missing
check(one(NT, observed_vba="", reference="1", arg2="3", arg3="4").n_missing == 1, "1 blank ordinary")
# 2. ERROR ordinary observation (in-envelope) -> error
check(one(NT, observed_vba="ERROR", reference="1", arg2="3", arg3="4").n_error == 1, "2 ERROR ordinary")
# 3. malformed hi;lo observation -> invalid
check(one(NT, observed_vba="1.0;abc", reference="1", arg2="3", arg3="4").n_invalid == 1, "3 malformed hi;lo")
# 4. malformed reference -> invalid
check(one(NT, observed_vba="1.0", reference="not_a_number", arg2="3", arg3="4").n_invalid == 1, "4 malformed reference")
# 5. blank tail-residual observation -> missing
check(one(TL, function="Beta_InverseCumulative", observed_vba="", arg1="0.5", arg2="2", arg3="3").n_missing == 1, "5 blank tail")
# 6. ERROR tail-residual observation -> error
check(one(TL, function="Beta_InverseCumulative", observed_vba="ERROR", arg1="0.5", arg2="2", arg3="3").n_error == 1, "6 ERROR tail")
# 7. missing tail argument -> invalid
check(one(TL, function="Beta_InverseCumulative", observed_vba="0.6", arg1="0.5", arg2="", arg3="3").n_invalid == 1, "7 missing tail arg")
# 8. mixed valid/invalid contract must NOT be fully measurable (a block remains)
mix8 = [
    {"function": "F_Cumulative", "observed_vba": "0.5", "reference": "0.5", "arg2": "3", "arg3": "4"},
    {"function": "F_Cumulative", "observed_vba": "0.6", "reference": "bad", "arg2": "3", "arg3": "4"},
]
d8 = dispositions(mix8, NT)
check(len(d8.to_measure) == 1 and (d8.n_missing + d8.n_error + d8.n_violation + d8.n_invalid) == 1,
      "8 mixed valid/invalid must block")

# row_validity direct: a clean row passes, malformed ones return a reason
check(row_validity({"observed_vba": "1.0", "reference": "2.0"}, NT) is None, "row_validity clean ordinary")
check(row_validity({"observed_vba": "1.0", "reference": "x"}, NT) is not None, "row_validity bad reference")
check(row_validity({"observed_vba": "1;2;x", "reference": "2"}, NT) is not None, "row_validity bad hi;lo")

# --- expected_error_drift ---------------------------------------------------
clean = [{"function": "F_Cumulative", "arg2": "3", "arg3": "2E10", "expected_error": "1"},
         {"function": "F_Cumulative", "arg2": "3", "arg3": "4", "expected_error": ""}]
check(expected_error_drift(clean) == [], "no drift when stored matches predicate")
drifted = [{"function": "F_Cumulative", "arg2": "3", "arg3": "4", "expected_error": "1"}]  # in-envelope mismarked
check(len(expected_error_drift(drifted)) == 1, "drift detected when stored disagrees")

# --- scaled_output_error: the scaling IS the measure ------------------------
# PROB_TryLogGamma1p at X = 1E-13. The scaled error is thirteen orders larger
# than the absolute one; scoring a scaled contract as ordinary absolute error
# would make it pass by that margin no matter how badly the kernel regressed.
so = Decimal("-5.7721566490145e-14")
sr = Decimal("-5.7721566490153e-14")
sx = Decimal("1e-13")
check(calculate_scaled_error(so, sr, sx) == abs(so - sr) / sx, "scaled divides by arg1")
check(calculate_scaled_error(so, sr, sx) > calculate_error(so, sr, "absolute") * Decimal(10) ** 12,
      "scaled dwarfs absolute at small arg1")

# Log(Gamma(1 + 0)) is exactly zero by contract: an equality for the VBA
# regression suite, never a scaled-error row.
for _bad in (Decimal(0), None):
    try:
        calculate_scaled_error(Decimal(1), Decimal(1), _bad)
        check(False, f"scaled must reject arg1={_bad!r}")
    except ValueError:
        pass

check(validate_scaled_metric("absolute") == "absolute", "scaled accepts metric=absolute")
for _bad in ("relative", "rel", ""):
    try:
        validate_scaled_metric(_bad)
        check(False, f"scaled must reject metric={_bad!r}")
    except ValueError:
        pass

# --- an unknown measure must fail loudly ------------------------------------
# The gate branches on tail_probability_residual and otherwise falls through to
# ordinary error, so a misspelled or unimplemented measure would be scored as
# something it is not.
for _good in sorted(KNOWN_MEASURES):
    check(validate_measure(_good) == _good, f"measure {_good} accepted")
for _bad in ("scaled_abs", "typo_error", "", None):
    try:
        validate_measure(_bad)
        check(False, f"measure must reject {_bad!r}")
    except ValueError:
        pass


# --- scaled_output_error preflight: arg1 must be present, parseable, non-zero
# Without this the measurer would skip such a row while dispositions() still
# called it valid, and a contract could PASS having scored nine of ten rows.
_base = {"function": "LogGamma1p", "arg2": "", "arg3": "", "arg4": "",
         "reference": "-5.7721566490145e-14",
         "observed_vba": "-5.7721566490145E-014;0E+000", "expected_error": ""}
check(row_validity(dict(_base, arg1="1e-13"), "scaled_output_error") is None,
      "valid scaled row passes preflight")
check(row_validity(dict(_base, arg1=""), "scaled_output_error") is not None,
      "scaled row with blank arg1 is invalid")
check(row_validity(dict(_base, arg1="0"), "scaled_output_error") is not None,
      "scaled row with arg1=0 is invalid")
check(row_validity(dict(_base, arg1="abc"), "scaled_output_error") is not None,
      "scaled row with unparseable arg1 is invalid")
check(row_validity(dict(_base, arg1="0"), "output_error") is None,
      "arg1=0 is fine for a non-scaled measure")

# --- tail-residual dispatch must be fail-closed, never a fall-through -------
# Both evaluators previously chose the forward CDF with
#   ibeta(...) if fn == "Beta_InverseCumulative" else f_cdf(...)
# so ANY other function was scored with the F CDF: silently, and returning a
# plausible number rather than raising. No registry verdict was affected,
# because the old preflight demanded arg1/arg2/arg3 and the unsupported
# inverses take only (p, df) - but that was accidental, and a populated arg3
# would have defeated it. These fixtures pin the guard, not the accident.
from _contract_eval import (TAIL_SUPPORTED, tail_cdf_name, tail_required_args,
                            UnsupportedTailFunction)

TR = "tail_probability_residual"

check(tail_cdf_name("Beta_InverseCumulative") == "ibeta",
      "Beta routes to the incomplete-beta CDF")
check(tail_cdf_name("F_InverseCumulative") == "f_cdf",
      "F routes to the F CDF")
check(set(TAIL_SUPPORTED) == {"Beta_InverseCumulative", "F_InverseCumulative"},
      "only Beta and F are registered at this commit")

# Every unsupported inverse must RAISE, not default to F.
for _fn in ("ChiSquare_InverseCumulative", "StudentT_InverseCumulative",
            "Gamma_InverseCumulative", "Normal_InverseCumulative", "", "beta_inversecumulative"):
    try:
        tail_cdf_name(_fn)
        check(False, f"unsupported tail function must raise: {_fn!r}")
    except UnsupportedTailFunction:
        pass
    try:
        tail_required_args(_fn)
        check(False, f"unsupported tail function must raise for args: {_fn!r}")
    except UnsupportedTailFunction:
        pass

# THE CASE THAT MATTERED: plausible three-argument data for an unsupported
# function. Under the old code this passed preflight and was scored with the F
# CDF; it must now be rejected before scoring.
_tail_ok = {"function": "F_InverseCumulative", "arg1": "0.9", "arg2": "5",
            "arg3": "7", "arg4": "", "reference": "",
            "observed_vba": "3.9715343910;0E+000",
            "expected_error": ""}
check(row_validity(_tail_ok, TR) is None, "valid F tail row passes preflight")
check(row_validity(dict(_tail_ok, function="Beta_InverseCumulative"), TR) is None,
      "valid Beta tail row passes preflight")

for _fn in ("ChiSquare_InverseCumulative", "StudentT_InverseCumulative"):
    _bad = dict(_tail_ok, function=_fn)
    _why = row_validity(_bad, TR)
    check(_why is not None,
          f"unsupported function with plausible 3-arg data is rejected: {_fn}")
    check("no registered" in (_why or ""),
          f"rejection names the missing registration, not a parse error: {_fn}")

# Missing or malformed required arguments still fail, per function.
for _k in ("arg1", "arg2", "arg3"):
    check(row_validity(dict(_tail_ok, **{_k: ""}), TR) is not None,
          f"F tail row with blank {_k} is invalid")
    check(row_validity(dict(_tail_ok, **{_k: "abc"}), TR) is not None,
          f"F tail row with unparseable {_k} is invalid")

# No unsupported function can reach the F evaluator: dispositions must not
# place such a row in to_measure.
_disp = dispositions([dict(_tail_ok, function="ChiSquare_InverseCumulative")], TR)
check(len(_disp.to_measure) == 0 and _disp.n_invalid == 1,
      "an unsupported tail row is blocked, not measured")
_disp_ok = dispositions([_tail_ok], TR)
check(len(_disp_ok.to_measure) == 1,
      "a supported tail row still reaches the measurer")


if fails:
    for f in fails:
        print("FAIL:", f)
    raise SystemExit(1)
print("PASS: shared contract-evaluation primitives (metric arithmetic, "
      "zero-reference, parsing, evidence classification, tail residual, "
      "scaled_output_error, measure whitelist, fail-closed tail dispatch)")
