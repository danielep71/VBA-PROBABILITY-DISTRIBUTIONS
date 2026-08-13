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
from collections import namedtuple

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


SCALED_MEASURE = "scaled_output_error"

# Every measure the evaluators know how to score. An unrecognised measure must
# fail loudly: the gate branches on tail_probability_residual and otherwise
# falls through to ordinary absolute/relative error, so a misspelled or
# not-yet-implemented measure would be scored as something it is not. That is
# how scaled_output_error would have behaved before it existed - passing by
# thirteen orders of magnitude and unable ever to fail.
KNOWN_MEASURES = frozenset({
    "output_error",
    "log_absolute_error",
    "quantile_error",
    "tail_probability_residual",
    SCALED_MEASURE,
})


def validate_measure(measure):
    """Reject any measure the evaluators cannot score."""
    m = (measure or "").strip()
    if m not in KNOWN_MEASURES:
        raise ValueError(
            f"Unsupported measure: {measure!r}; expected one of "
            f"{sorted(KNOWN_MEASURES)}. An unknown measure would be scored as "
            f"ordinary error, which silently changes what the contract claims.")
    return m


def calculate_scaled_error(observed, reference, arg1):
    """
    Scaled output error: Abs(observed - reference) / Abs(arg1).

    Distinct from `absolute`, which is the unscaled difference. A contract
    declaring measure=scaled_output_error MUST route here; scoring it as
    ordinary absolute error would make it unfailable. PROB_TryLogGamma1p is
    the motivating case: at X = 1E-13 its scaled error is 1.4E-16 while the
    absolute error is 1.4E-29, so a 5E-16 absolute threshold would pass by
    thirteen orders of magnitude no matter how badly the kernel regressed.

    The metric is what the caller actually propagates. The scaled Gamma
    inverse computes [LogProbability + LogGamma1p(Shape)] / Shape, so this
    quantity IS the relative error of the quantile it produces.

    arg1 = 0 is rejected rather than divided. Log(Gamma(1 + 0)) is exactly
    zero by contract, which belongs in the VBA regression suite as an exact
    equality, not in a scaled-error row.
    """
    if arg1 is None:
        raise ValueError("scaled_output_error requires arg1")
    if arg1 == 0:
        raise ValueError(
            "scaled_output_error is undefined at arg1 = 0; the exact-zero case "
            "belongs in the VBA regression suite, not in a scaled contract row")
    return abs(observed - reference) / abs(arg1)


def validate_scaled_metric(metric):
    """A scaled contract must declare metric=absolute: the scaling is the
    measure, and calling it relative would imply a second division."""
    m = (metric or "").strip().lower()
    if m != "absolute":
        raise ValueError(
            f"measure={SCALED_MEASURE!r} requires metric='absolute', got {metric!r}")
    return m


def calculate_error(observed, reference, metric):
    """
    Ordinary (non-residual) error. `metric` selects the arithmetic:
      absolute -> Abs(observed - reference)
      relative -> Abs(observed - reference) / Abs(reference)
    A non-zero observed against a zero reference is infinite relative error;
    a zero observed against a zero reference is zero error.

    Scaled contracts do NOT come here; see calculate_scaled_error.
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
# Mirrors of the VBA envelope constants in M_STATS_PROBDIST_TFAMILY.bas. The
# enveloped CDF/Survival/Inverse public UDFs return #NUM! when a degree of
# freedom exceeds the cap; the densities are closed-form and deliberately NOT
# enveloped here (they carry their own PROB_DENSITY_SHAPE_MAX). Rows in a
# reject region carry no accuracy claim, so a #NUM! there is the CORRECT
# response, not a defect.
#
# KEEP THESE IN SYNC WITH THE VBA CONSTANTS. They were raised on 2026-07-29
# from 1E5/1E6/1E6 after benchmark/envelope_probe showed the previous
# boundaries were measuring the CR-P1-02 prefactor cancellation rather than an
# intrinsic limit; check_source_thresholds.py verifies the pairing.
# The inverses share the forward caps. They were briefly capped lower, on the
# assumption that a measured F_InverseCumulative refusal reflected the inverse
# iteration giving up; benchmark/inverse_probe showed the refusal was the newly
# installed inverse cap itself, and that the inverse kernel converges cleanly to
# df 2E10 with no refusals at any probed shape, including the unbalanced ratios.
F_MAX_DF = 10000000000.0        # PROB_F_MAX_DF
T_MAX_DF = 100000000.0          # PROB_T_MAX_DF
CHI_MAX_DF = 100000000.0        # PROB_CHI_MAX_DF

_F_ENVELOPED = ("F_Cumulative", "F_Survival", "F_InverseCumulative")
_T_ENVELOPED = ("StudentT_Cumulative", "StudentT_Survival",
                "StudentT_InverseCumulative")
_CHI_ENVELOPED = ("ChiSquare_Cumulative", "ChiSquare_Survival",
                  "ChiSquare_InverseCumulative")


def predicted_expected_error(function, arg2, arg3):
    """
    Design-intent predicate: True when a row lies in an envelope-reject region,
    where #NUM! is the correct response. Derived from the row's own args,
    independent of what was observed, so it can never launder an unexpected
    error into an expected one.

    F is enveloped on BOTH degrees of freedom; StudentT and ChiSquare take
    their single df in arg2.
    """
    if function in _F_ENVELOPED:
        cap, args = F_MAX_DF, (arg2, arg3)
    elif function in _T_ENVELOPED:
        cap, args = T_MAX_DF, (arg2,)
    elif function in _CHI_ENVELOPED:
        cap, args = CHI_MAX_DF, (arg2,)
    else:
        return False

    for a in args:
        try:
            if float(a) > cap:
                return True
        except (TypeError, ValueError):
            continue
    return False


def row_expected_error(row):
    """Truth of the stored `expected_error` grid column (blank/0 -> False)."""
    v = (row.get("expected_error", "") or "").strip().lower()
    return v in ("1", "true", "yes")


# Row dispositions for an active contract's matched grid rows.
MEASURE = "measure"                    # a usable, fully-valid in-envelope observation
EXCLUDE_EXPECTED = "exclude_expected"  # envelope-reject region, correctly #NUM! -> not scored
BLOCK_MISSING = "block_missing"        # active row not yet observed
BLOCK_ERROR = "block_error"            # unexpected kernel error inside the envelope
BLOCK_EXPECTED = "block_expected"      # envelope-reject row that did NOT return #NUM!
BLOCK_INVALID = "block_invalid"        # observation/reference/args present but unparseable


def classify_row(observed_raw, expected_error):
    """
    Disposition of one matched grid row by OBSERVATION STATE and envelope marker.
    A MEASURE result here is provisional: the row still has to pass row_validity
    (parseable observation, reference, and required args) before it is scored.

      expected_error & observed ERROR -> EXCLUDE_EXPECTED (correct refusal; not scored)
      expected_error & not ERROR      -> BLOCK_EXPECTED    (envelope failed to fire)
      blank observation               -> BLOCK_MISSING     (active contract, unobserved)
      ERROR observation               -> BLOCK_ERROR       (unexpected error in-envelope)
      otherwise                       -> MEASURE (pending validity)
    """
    st = observation_state(observed_raw)
    if expected_error:
        # Direction 2: an envelope-reject row MUST observe ERROR - the #NUM! the
        # predicate says is correct. A value or a blank there means the envelope
        # did not fire, so stale or envelope-inconsistent data cannot pass silently.
        return EXCLUDE_EXPECTED if st == ERROR else BLOCK_EXPECTED
    if st == MISSING:
        return BLOCK_MISSING
    if st == ERROR:
        return BLOCK_ERROR
    return MEASURE


def _parses_float(v):
    try:
        float(v)
        return True
    except (TypeError, ValueError):
        return False


def row_validity(row, measure):
    """
    Preflight for a row already classified MEASURE. Returns None if the row is
    fully evaluable, else a human-readable reason. This closes the gap where a
    non-blank, non-ERROR row could still be silently skipped by the measurer
    because its reference, observation, or a required argument would not parse.
      * the observation must parse in full (every hi;lo part), not merely be non-blank;
      * ordinary contracts need a parseable reference;
      * tail_probability_residual contracts need parseable arg1/arg2/arg3.
    """
    obs = (row.get("observed_vba", "") or "").strip()
    try:
        total = Decimal(0)
        for part in obs.split(";"):
            total += Decimal(part.strip())
    except (InvalidOperation, ValueError):
        return f"unparseable observation {obs!r}"
    if measure == "tail_probability_residual":
        for k in ("arg1", "arg2", "arg3"):
            v = (row.get(k, "") or "").strip()
            if v == "":
                return f"missing required {k}"
            if not _parses_float(v):
                return f"unparseable {k} {v!r}"
    else:
        if parse_reference(row.get("reference", "")) is None:
            return f"unparseable reference {(row.get('reference','') or '').strip()!r}"
    return None


def _row_id(row):
    args = ", ".join(f"{k}={row.get(k,'')}" for k in ("arg1", "arg2", "arg3", "arg4")
                     if (row.get(k, "") or "").strip())
    return args or "(no args)"


# Self-describing partition of a contract's matched rows. Attribute access keeps
# callers stable as fields are added; `reasons` lists every blocking/invalid row
# so the gate can name exactly what failed rather than silently omitting it.
Dispositions = namedtuple(
    "Dispositions",
    "to_measure n_expected n_missing n_error n_violation n_invalid reasons")


def dispositions(rows, measure=None):
    """
    Partition a contract's matched grid rows for scoring, measure-aware.
      to_measure   fully-valid, in-envelope rows to score
      n_expected   envelope-reject rows correctly returning #NUM! (excluded)
      n_missing    blank observations (block)
      n_error      unexpected in-envelope ERROR (block)
      n_violation  envelope-reject rows that did NOT return #NUM! (block)
      n_invalid    rows whose observation/reference/args will not parse (block)
      reasons      (row-id, reason) for every blocking/invalid row
    len(to_measure) == matched - n_expected only when there are no blocks; any
    nonzero block count means the evidence is incomplete and must not PASS.
    """
    to_measure = []
    n_expected = n_missing = n_error = n_violation = n_invalid = 0
    reasons = []
    for r in rows:
        exp = predicted_expected_error(r.get("function", ""), r.get("arg2", ""), r.get("arg3", ""))
        d = classify_row(r.get("observed_vba", ""), exp)
        if d == MEASURE:
            reason = row_validity(r, measure)
            if reason is None:
                to_measure.append(r)
            else:
                n_invalid += 1
                reasons.append((_row_id(r), reason))
        elif d == EXCLUDE_EXPECTED:
            n_expected += 1
        elif d == BLOCK_MISSING:
            n_missing += 1
            reasons.append((_row_id(r), "unobserved (blank)"))
        elif d == BLOCK_ERROR:
            n_error += 1
            reasons.append((_row_id(r), "unexpected ERROR in-envelope"))
        else:                       # BLOCK_EXPECTED
            n_violation += 1
            reasons.append((_row_id(r), "envelope-reject row did not return #NUM!"))
    return Dispositions(to_measure, n_expected, n_missing, n_error, n_violation, n_invalid, reasons)


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
