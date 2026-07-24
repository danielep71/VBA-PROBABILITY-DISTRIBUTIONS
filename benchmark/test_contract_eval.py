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
from _contract_eval import (
    parse_observed, parse_reference, parse_threshold, validate_metric,
    calculate_error, normalize_tail_residual, observation_state,
    evidence_gaps, OK, MISSING, ERROR,
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

if fails:
    print("FAIL: _contract_eval primitives")
    for f in fails:
        print("  - " + f)
    raise SystemExit(1)
print("PASS: shared contract-evaluation primitives (metric arithmetic, zero-reference, "
      "parsing, evidence classification, tail residual)")
