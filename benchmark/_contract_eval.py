"""
Single-sourced contract-evaluation primitives.

Both the release gate (`compute_errors.py`) and the independent holdout
analyzer (`holdout/analyze_holdout.py`) import from here so they can never
diverge on HOW a contract is measured. This module exists because they once
did diverge: `analyze_holdout.py` keyed absolute-vs-relative off `measure`
while the gate keyed it off `metric`, so `PROB_LogBeta.tiny_unbalanced` was
evaluated relative in the holdout and absolute in the gate for the same row.

Division of responsibility, kept identical everywhere:
  * `measure` selects the evaluation PATH (ordinary error vs tail residual).
  * `metric`  selects the ARITHMETIC (absolute vs relative).

This module deliberately does NOT import the reference helper (`_ibeta`) or
mpmath, so it stays importable even when those are unavailable; the tail path
passes already-evaluated `recovered`/`target` values in.
"""
from decimal import Decimal, InvalidOperation

# Observation-cell states.
OK = "ok"            # a usable numeric observation (possibly two-part hi;lo)
MISSING = "missing"  # blank cell: measurement not yet run
ERROR = "error"      # explicit ERROR sentinel: the kernel returned an error


def observation_state(raw):
    """Classify a raw observed_vba cell as OK / MISSING / ERROR."""
    t = (raw or "").strip()
    if t == "":
        return MISSING
    if t.upper() == "ERROR":
        return ERROR
    return OK


def parse_observed(raw):
    """Sum the two-part hi;lo export to a Decimal, or None for MISSING/ERROR."""
    t = (raw or "").strip()
    if t == "" or t.upper() == "ERROR":
        return None
    total = Decimal(0)
    for part in t.split(";"):
        total += Decimal(part.strip())
    return total


def parse_reference(raw):
    """Parse a reference cell to Decimal, or None if malformed."""
    try:
        return Decimal(str(raw).strip())
    except (InvalidOperation, AttributeError, ValueError):
        return None


def parse_threshold(raw):
    """Parse a contract threshold to Decimal, or None if blank/malformed."""
    if raw is None or str(raw).strip() == "":
        return None
    try:
        return Decimal(str(raw).strip())
    except (InvalidOperation, ValueError):
        return None


def validate_metric(metric):
    """Normalise and validate a contract metric; raise on anything else."""
    m = (metric or "").strip().lower()
    if m not in ("absolute", "relative"):
        raise ValueError(f"Unsupported metric: {metric!r}; expected 'absolute' or 'relative'")
    return m


def calculate_error(observed, reference, metric):
    """
    Ordinary (non-residual) error. `metric` selects the arithmetic:
      absolute -> Abs(observed - reference)
      relative -> Abs(observed - reference) / Abs(reference)
    A non-zero observed against a zero reference is infinite relative error;
    a zero observed against a zero reference is zero error.
    """
    m = validate_metric(metric)
    absolute_error = abs(observed - reference)
    if m == "absolute":
        return absolute_error
    if reference != 0:
        return absolute_error / abs(reference)
    return Decimal(0) if observed == 0 else Decimal("Infinity")


def normalize_tail_residual(recovered, target):
    """
    Tail-residual normalisation for an inverse round-trip: the forward-CDF value
    `recovered` at the returned quantile vs the target probability `target`,
    normalised by min(p, 1 - p). Both inputs are numeric (mpmath mpf); returns
    a Decimal. Kept here so the gate and holdout normalise residuals identically.
    """
    return Decimal(str(abs(recovered - target) / min(target, 1 - target)))


def evidence_gaps(rows, observed_key="observed_vba"):
    """
    Count evidence gaps among a contract's matched grid rows.
    Returns (n_missing, n_error): cells that are blank or the ERROR sentinel.
    Callers decide the policy (both currently block on MISSING; ERROR policy is
    contract/regime-specific because some ERROR results are the correct #NUM!
    envelope response).
    """
    n_missing = n_error = 0
    for r in rows:
        st = observation_state(r.get(observed_key, ""))
        if st == MISSING:
            n_missing += 1
        elif st == ERROR:
            n_error += 1
    return n_missing, n_error


# --- Envelope / expected-error marker ---------------------------------------
# Mirror of the VBA constant PROB_F_MAX_DF (M_STATS_PROBDIST_TFAMILY.bas): the
# F CDF/Survival/Inverse public UDFs return #NUM! when either degree of freedom
# exceeds this. F_Density is closed-form and deliberately NOT enveloped. Rows in
# this region carry no accuracy claim, so a #NUM! there is the CORRECT response,
# not a defect. Keep this in sync with the VBA constant.
F_MAX_DF = 100000.0
_F_ENVELOPED = ("F_Cumulative", "F_Survival", "F_InverseCumulative")


def predicted_expected_error(function, arg2, arg3):
    """
    Design-intent predicate: True when a row lies in the F envelope-reject region
    (an enveloped F function with either df > F_MAX_DF), where #NUM! is correct.
    Derived from the row's own args, independent of what was observed, so it can
    never launder an unexpected error into an expected one.
    """
    if function not in _F_ENVELOPED:
        return False
    for a in (arg2, arg3):
        try:
            if float(a) > F_MAX_DF:
                return True
        except (TypeError, ValueError):
            continue
    return False


def row_expected_error(row):
    """Truth of the stored `expected_error` grid column (blank/0 -> False)."""
    v = (row.get("expected_error", "") or "").strip().lower()
    return v in ("1", "true", "yes")


# Row dispositions for an active contract's matched grid rows.
MEASURE = "measure"                    # a usable in-envelope observation -> score it
EXCLUDE_EXPECTED = "exclude_expected"  # envelope-reject region -> no accuracy claim
BLOCK_MISSING = "block_missing"        # active row not yet observed
BLOCK_ERROR = "block_error"            # unexpected kernel error inside the envelope


def classify_row(observed_raw, expected_error):
    """
    Disposition of one matched grid row. `expected_error` is truthy when the row
    is outside the validated envelope (a #NUM! is the correct response).

      expected_error            -> EXCLUDE_EXPECTED  (no accuracy claim; not scored)
      blank observation         -> BLOCK_MISSING     (active contract, unobserved)
      ERROR observation         -> BLOCK_ERROR       (unexpected error in-envelope)
      otherwise                 -> MEASURE
    """
    if expected_error:
        return EXCLUDE_EXPECTED
    st = observation_state(observed_raw)
    if st == MISSING:
        return BLOCK_MISSING
    if st == ERROR:
        return BLOCK_ERROR
    return MEASURE


def dispositions(rows):
    """
    Partition a contract's matched grid rows for scoring. The envelope predicate
    (derived from each row's own args) identifies expected #NUM! rows, which are
    excluded from scoring, and distinguishes them from unexpected in-envelope
    errors, which must block.
    Returns (to_measure, n_expected, n_missing, n_error).
    """
    to_measure = []
    n_expected = n_missing = n_error = 0
    for r in rows:
        exp = predicted_expected_error(r.get("function", ""), r.get("arg2", ""), r.get("arg3", ""))
        d = classify_row(r.get("observed_vba", ""), exp)
        if d == MEASURE:
            to_measure.append(r)
        elif d == EXCLUDE_EXPECTED:
            n_expected += 1
        elif d == BLOCK_MISSING:
            n_missing += 1
        else:
            n_error += 1
    return to_measure, n_expected, n_missing, n_error


def expected_error_drift(rows):
    """
    Rows whose stored `expected_error` column disagrees with the predicate, so a
    hand-edited or stale marker cannot silently diverge from design intent.
    Rows without the column are skipped (predicate stays authoritative).
    """
    bad = []
    for r in rows:
        if "expected_error" not in r:
            continue
        pred = predicted_expected_error(r.get("function", ""), r.get("arg2", ""), r.get("arg3", ""))
        if row_expected_error(r) != pred:
            bad.append(r)
    return bad
