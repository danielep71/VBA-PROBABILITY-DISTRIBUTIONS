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
