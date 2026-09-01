"""
Regression test for the single-sourced contract-evaluation primitives.

Locks the arithmetic that compute_errors.py (gate) and analyze_holdout.py
(holdout) both import, so the two can never diverge again. Proves:
  * metric (not measure) chooses absolute vs relative;
  * a zero reference is handled as 0 (when observed is 0) or Infinity;
  * observation/reference/threshold parsing and evidence classification.
Run: python3 test_contract_eval.py   (exit 0 = pass, nonzero = fail)
"""
import os
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
                            tail_shape_args, UnsupportedTailFunction)

TR = "tail_probability_residual"

check(tail_cdf_name("Beta_InverseCumulative") == "ibeta",
      "Beta routes to the incomplete-beta CDF")
check(tail_cdf_name("F_InverseCumulative") == "f_cdf",
      "F routes to the F CDF")
check(tail_cdf_name("StudentT_InverseCumulative") == "t_cdf",
      "Student-t routes to the Student-t CDF")
check(tail_cdf_name("ChiSquare_InverseCumulative") == "chi2_cdf",
      "Chi-square routes to the incomplete-gamma CDF")
check(set(TAIL_SUPPORTED) == {"Beta_InverseCumulative", "F_InverseCumulative",
                              "StudentT_InverseCumulative",
                              "ChiSquare_InverseCumulative"},
      "Beta, F, Student-t and Chi-square are registered")
check(tail_shape_args("ChiSquare_InverseCumulative") == ("arg2",),
      "Chi-square takes one shape argument")
# Arity comes from the registry, not a fixed three-column assumption.
check(tail_shape_args("Beta_InverseCumulative") == ("arg2", "arg3"),
      "Beta takes two shape arguments")
check(tail_shape_args("F_InverseCumulative") == ("arg2", "arg3"),
      "F takes two shape arguments")
check(tail_shape_args("StudentT_InverseCumulative") == ("arg2",),
      "Student-t takes one shape argument")

# Every unsupported inverse must RAISE, not default to F.
for _fn in ("Gamma_InverseCumulative", "Normal_InverseCumulative", "",
            "beta_inversecumulative"):
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

for _fn in ("Gamma_InverseCumulative", "Normal_InverseCumulative"):
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
_disp = dispositions([dict(_tail_ok, function="Gamma_InverseCumulative")], TR)
check(len(_disp.to_measure) == 0 and _disp.n_invalid == 1,
      "an unsupported tail row is blocked, not measured")
_disp_ok = dispositions([_tail_ok], TR)
check(len(_disp_ok.to_measure) == 1,
      "a supported tail row still reaches the measurer")

# --- the MAIN EVALUATOR must actually use the registry ----------------------
# The fixtures above pin _contract_eval's registry. They do NOT prove that
# compute_errors.tail_residual consults it: reintroducing the binary branch in
# compute_errors.py alone left every committed fixture green, because those
# fixtures never call the evaluator. This block closes that gap by asserting
# the value the evaluator RETURNS, which is only reproducible if it dispatched
# to the correct CDF. ibeta and f_cdf differ by a wide margin at these
# arguments (0.5798 vs 0.2467), so a mis-dispatch cannot coincide.
import compute_errors as _CE
from _ibeta import ibeta as _ib, f_cdf as _fc
import mpmath as _mpx
_mpx.mp.dps = 50

def _tail_row(fn, p_, a2, a3, x):
    return {"function": fn, "arg1": p_, "arg2": a2, "arg3": a3, "arg4": "",
            "reference": "", "observed_vba": x, "expected_error": ""}

# Beta must route to ibeta, NOT f_cdf.
_bx, _ba, _bb, _bp = "0.3", "2", "5", "0.4"
_row_b = _tail_row("Beta_InverseCumulative", _bp, _ba, _bb, _bx)
_got_b, _, _n_b = _CE.tail_residual([_row_b], "Beta_InverseCumulative")
_want_b = normalize_tail_residual(_ib(_mpx.mpf(_bx), _mpx.mpf(_ba), _mpx.mpf(_bb)), _mpx.mpf(_bp))
_wrong_b = normalize_tail_residual(_fc(_mpx.mpf(_bx), _mpx.mpf(_ba), _mpx.mpf(_bb)), _mpx.mpf(_bp))
check(_n_b == 1, "Beta tail row is scored")
check(_got_b == _want_b, "main evaluator dispatched Beta to ibeta")
check(_got_b != _wrong_b, "main evaluator did NOT dispatch Beta to f_cdf")

# F must route to f_cdf, NOT ibeta.
_fx, _f1, _f2, _fp = "0.5", "5", "7", "0.4"
_row_f = _tail_row("F_InverseCumulative", _fp, _f1, _f2, _fx)
_got_f, _, _n_f = _CE.tail_residual([_row_f], "F_InverseCumulative")
_want_f = normalize_tail_residual(_fc(_mpx.mpf(_fx), _mpx.mpf(_f1), _mpx.mpf(_f2)), _mpx.mpf(_fp))
_wrong_f = normalize_tail_residual(_ib(_mpx.mpf(_fx), _mpx.mpf(_f1), _mpx.mpf(_f2)), _mpx.mpf(_fp))
check(_n_f == 1, "F tail row is scored")
check(_got_f == _want_f, "main evaluator dispatched F to f_cdf")
check(_got_f != _wrong_f, "main evaluator did NOT dispatch F to ibeta")

# An unsupported function must RAISE out of the evaluator itself, so it can
# never be scored through F. compute_errors turns this into PENDING, never a
# verdict.
for _fn in ("Gamma_InverseCumulative", "Normal_InverseCumulative"):
    try:
        _CE.tail_residual([_tail_row(_fn, "0.9", "1000000", "1", "1039569.3")], _fn)
        check(False, f"main evaluator must refuse to score {_fn!r}")
    except UnsupportedTailFunction:
        pass

# --- Student-t: sign-aware CDF, registry-driven arity ----------------------
from _ibeta import t_cdf as _tc

# Branch correctness. The positive-t expression alone returns 1 - h for
# negative t, which is the reflection about 1/2 - a plausible probability and
# therefore a silently wrong residual. Every Student-t row at p <= 0.5 has a
# negative or zero quantile, so the lower half of the surface depends on this.
check(_tc(0, 30) == _mpx.mpf(1) / 2, "t_cdf(0, df) is exactly 1/2")
check(_tc("-1.5", 30) + _tc("1.5", 30) == 1, "t_cdf is symmetric about zero")
check(_tc("-1.28155241212994", 1e6) < _mpx.mpf("0.11"),
      "negative t maps to the lower tail, not its reflection")
check(_tc("1.28155241212994", 1e6) > _mpx.mpf("0.89"),
      "positive t maps to the upper tail")

# The evaluator itself, on committed-shaped rows, at all three envelope df.
# Values are the observed quantiles from the grid; the residual must be tiny.
# Quantiles are the committed observations for these rows, taken from the
# grid rather than invented, so a fixture failure means the evaluator moved
# and not that the fixture was wrong.
# Per-row residual bounds, NOT a blanket tolerance. The p=0.99 rows carry a
# genuinely larger residual that grows with df, and the committed `reference`
# column independently implies the same magnitude (1.85E-10 relative at
# df=1E8), so the tail residual is reproducing evidence already in the grid
# rather than reporting an oracle artifact. A single loose tolerance would
# hide that; a single tight one would fail on real data.
_T_ROWS = (
    ("0.9",  "1000000.0",   "1.28155241212994E+000;-3.77475828372553E-015", "1e-14"),
    ("0.9",  "10000000.0",  "1.28155165020308E+000;1.11022302462516E-015",  "1e-14"),
    ("0.9",  "100000000.0", "1.28155157401045E+000;-1.11022302462516E-015", "1e-14"),
    ("0.99", "100000000.0", "2.32634791090167E+000;0E+000",                 "1e-8"),
    ("0.5",  "1000000.0",   "0E+000;0E+000",                                "0"),
    ("0.05", "1.0",         "-6.31375151467504E+000;-3.55271367880050E-015", "1e-14"),
)
for _p, _df, _x, _bound in _T_ROWS:
    _row = _tail_row("StudentT_InverseCumulative", _p, _df, "", _x)
    _got, _, _n = _CE.tail_residual([_row], "StudentT_InverseCumulative")
    _want = normalize_tail_residual(
        _tc(_mpx.mpf(str(parse_observed(_x))), _mpx.mpf(_df)), _mpx.mpf(_p))
    check(_n == 1, f"Student-t row scored: p={_p} df={_df}")
    check(_got == _want, f"Student-t dispatched to t_cdf: p={_p} df={_df}")
    check(_got <= Decimal(_bound),
          f"Student-t residual within its recorded bound: p={_p} df={_df}")

# The median row read BOTH ways. As tail-residual evidence the residual is
# exactly zero; as output_error evidence the same row is the zero-reference
# case. One row, two readings, and neither may be assumed.
_med = _tail_row("StudentT_InverseCumulative", "0.5", "1000000", "", "0E+000")
_got_med, _, _ = _CE.tail_residual([_med], "StudentT_InverseCumulative")
check(_got_med == Decimal(0), "Student-t median residual is exactly zero")
check(calculate_error(Decimal(0), Decimal(0), "relative") == Decimal(0),
      "the same median row is the zero-reference case for output_error")
check(calculate_error(Decimal("1e-30"), Decimal(0), "relative") == Decimal("Infinity"),
      "a non-zero quantile against the zero median reference is infinite error")

# A deliberately degraded observation must produce a large residual, so the
# measure has discrimination rather than merely not erroring.
_bad = _tail_row("StudentT_InverseCumulative", "0.9", "1000000", "",
                 "1.29155241212994E+000")
_got_bad, _, _ = _CE.tail_residual([_bad], "StudentT_InverseCumulative")
check(_got_bad > Decimal("1e-4"),
      "a perturbed Student-t quantile gives a large residual")

# Compensated hi;lo observations must be summed before scoring, not truncated
# to hi. The lo limb here is chosen to matter at the residual scale.
_hi = "1.28155241212994E+000"
_comp = _tail_row("StudentT_InverseCumulative", "0.9", "1000000", "",
                  _hi + ";-3.7E-017")
_got_c, _, _ = _CE.tail_residual([_comp], "StudentT_InverseCumulative")
_got_h, _, _ = _CE.tail_residual(
    [_tail_row("StudentT_InverseCumulative", "0.9", "1000000", "", _hi)],
    "StudentT_InverseCumulative")
check(_got_c != _got_h,
      "the lo limb of a compensated observation changes the Student-t residual")

# Registry-driven arity: a Student-t row has no arg3, and must still validate
# and score. Under the previous fixed (arg2, arg3) dispatch this was impossible.
check(row_validity(_tail_row("StudentT_InverseCumulative", "0.9", "1000000", "",
                             "1.28155241212994E+000"), TR) is None,
      "a two-argument Student-t row validates without arg3")
check(row_validity(_tail_row("StudentT_InverseCumulative", "0.9", "", "",
                             "1.28"), TR) is not None,
      "a Student-t row missing df is still invalid")

# Beta and F must be untouched by the arity change.
check(_CE.tail_residual([_row_b], "Beta_InverseCumulative")[0] == _want_b,
      "Beta residual unchanged after registry-driven arity")
check(_CE.tail_residual([_row_f], "F_InverseCumulative")[0] == _want_f,
      "F residual unchanged after registry-driven arity")

# If the helper table loses a registered callable, the evaluator must fail
# rather than score - the caller turns that into PENDING.
_saved = dict(_CE._TAIL_CDFS)
try:
    del _CE._TAIL_CDFS["t_cdf"]
    try:
        _CE.tail_residual([_med], "StudentT_InverseCumulative")
        check(False, "a missing callable must not be scored")
    except KeyError:
        pass
finally:
    _CE._TAIL_CDFS.clear(); _CE._TAIL_CDFS.update(_saved)
check(_CE.tail_residual([_med], "StudentT_InverseCumulative")[2] == 1,
      "the callable table is restored after the missing-helper case")

# --- Chi-square: non-convergence must never become a verdict ---------------
import _igamma as _IG

_chi_row = _tail_row("ChiSquare_InverseCumulative", "0.5", "1000000.0", "",
                     "9.99999333333411E+005")
_got_chi, _, _n_chi = _CE.tail_residual([_chi_row], "ChiSquare_InverseCumulative")
check(_n_chi == 1, "Chi-square tail row is scored")
check(_got_chi < Decimal("1e-10"),
      "a committed Chi-square envelope observation gives a small residual")
check(row_validity(_chi_row, TR) is None,
      "a two-argument Chi-square row validates without arg3")

# The kernel must RAISE when a route exhausts its cap, not return a truncated
# sum. Both caps are exercised: the series below the seam, the CF above it.
_saved_series, _saved_cf = _IG.MAX_SERIES_TERMS, _IG.MAX_CF_ITERATIONS
try:
    _IG.MAX_SERIES_TERMS = 3
    try:
        _IG.lower_series(_mpx.mpf(500000), _mpx.mpf(499999))
        check(False, "an exhausted series cap must raise")
    except _IG.IGammaNonConvergence:
        pass
    _IG.MAX_SERIES_TERMS = _saved_series
    _IG.MAX_CF_ITERATIONS = 3
    try:
        _IG.upper_cf(_mpx.mpf(500000), _mpx.mpf(500002))
        check(False, "an exhausted CF cap must raise")
    except _IG.IGammaNonConvergence:
        pass
finally:
    _IG.MAX_SERIES_TERMS, _IG.MAX_CF_ITERATIONS = _saved_series, _saved_cf

# There must be no mpmath.gammainc fallback anywhere in the kernel: a fallback
# would silently reintroduce the failure the module exists to avoid, and it
# fails only at some shapes, so its absence cannot be inferred from a passing
# evaluation at df = 1E6.
with open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "_igamma.py"), encoding="utf-8") as _f:
    _igamma_src = _f.read()
check("gammainc" not in _igamma_src.split('"""', 2)[2],
      "the kernel never calls mpmath.gammainc outside its docstring")

# Non-convergence must propagate out of the evaluator so the caller can record
# PENDING. It must NOT be swallowed into a number.
_saved_series = _IG.MAX_SERIES_TERMS
try:
    _IG.MAX_SERIES_TERMS = 3
    try:
        _CE.tail_residual([_chi_row], "ChiSquare_InverseCumulative")
        check(False, "non-convergence must not yield a Chi-square residual")
    except _IG.IGammaNonConvergence:
        pass
finally:
    _IG.MAX_SERIES_TERMS = _saved_series

# A registered function whose callable is missing must fail, not score.
_saved = dict(_CE._TAIL_CDFS)
try:
    del _CE._TAIL_CDFS["chi2_cdf"]
    try:
        _CE.tail_residual([_chi_row], "ChiSquare_InverseCumulative")
        check(False, "a missing chi2_cdf callable must not be scored")
    except KeyError:
        pass
finally:
    _CE._TAIL_CDFS.clear(); _CE._TAIL_CDFS.update(_saved)

# Wrong dispatch: Chi-square scored through any other CDF must differ, so a
# mis-registration cannot pass unnoticed.
_chi_x = _mpx.mpf("9.99999333333411E+005")
check(_IG.chi2_cdf(_chi_x, _mpx.mpf("1000000.0")) != _tc(_chi_x, _mpx.mpf("1000000.0")),
      "Chi-square and Student-t CDFs are distinguishable at the same arguments")


if fails:
    for f in fails:
        print("FAIL:", f)
    raise SystemExit(1)
print("PASS: shared contract-evaluation primitives (metric arithmetic, "
      "zero-reference, parsing, evidence classification, tail residual, "
      "scaled_output_error, measure whitelist, fail-closed tail dispatch)")
