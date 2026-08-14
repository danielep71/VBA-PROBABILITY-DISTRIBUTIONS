#!/usr/bin/env python3
"""
test_analyze_positive_ratio_subnormal.py
========================================

Unit tests for the Phase A analyzer.

WHY THIS EXISTS
---------------
The analyzer defines the evidence contract for #13: what a row means, how the
exact mathematical ratio is reconstructed from the binary64 operands VBA
actually received, how bucket membership is validated, and how hard underflow
is separated from positive-subnormal precision loss.

The exact-rational reconstruction is subtle enough to deserve coverage that
does not depend on the VBA exporter existing or being correct. These fixtures
are synthetic by design: every value is constructed from first principles here,
so a failure indicts the analyzer rather than the evidence.

The deliberately-mislabelled cases matter as much as the correct ones. An
analyzer that reports a crossover from data it should have rejected is worse
than one that reports nothing.

Run with: python -m pytest test_analyze_positive_ratio_subnormal.py
      or: python test_analyze_positive_ratio_subnormal.py
"""

from __future__ import annotations

import csv
import math
import subprocess
import sys
import tempfile
from fractions import Fraction
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from analyze_positive_ratio_subnormal import (      # noqa: E402
    COLUMNS, MIN_SUBNORMAL, Point, bits_available, load, parse_hilo,
    REGISTERED_CLAIMS, classify_output, crossover_of, envelope,
    reference, twin_check,
    verify_constructions, verify_reference_invariants,
)

HERE = Path(__file__).parent
ANALYZER = HERE / "analyze_positive_ratio_subnormal.py"


# ---------------------------------------------------------------------------
# construction of synthetic evidence
# ---------------------------------------------------------------------------

def hilo(x: float) -> str:
    """Mirror the exporter's two-part token so the parser is tested against the
    format it will actually receive, not a convenient approximation."""
    if x == 0.0:
        return "0E+000;0E+000"
    hi = float(f"{x:.14E}")
    return f"{x:.14E};{x - hi:.14E}"


def gamma_rows(k: int) -> list[dict]:
    """The four Gamma constructions for bucket k.

    N * m is the landmark carrying exactly k bits. The three stress points have
    exact quotients at 1/3, 1/2 and 2/3 of a subnormal ULP above it, reached
    with integer numerators over scales of 3, 2 and 3. Every X below is exactly
    representable -- that is the property the recipe was chosen for.
    """
    m = Fraction(2) ** -1074
    N = 2 ** (k - 1)
    cases = [
        ("landmark", "landmark", N * m, 1),
        ("s13", "transform_stress", (3 * N + 1) * m, 3),
        ("s12", "transform_stress", (2 * N + 1) * m, 2),
        ("s23", "transform_stress", (3 * N + 2) * m, 3),
    ]
    rows = []
    for tag, construction, x_exact, scale in cases:
        x = float(x_exact)
        assert Fraction(x) == x_exact, f"k={k} {tag}: X is not representable"
        for surface in ("density", "cumulative", "survival"):
            rows.append({
                "family": "gamma", "surface": surface,
                "point_id": f"g_{tag}_k{k}", "construction": construction,
                "shape_id": "s0.5", "bucket_bits": k,
                "echo_x": hilo(x), "echo_shape": hilo(0.5),
                "echo_scale": hilo(float(scale)),
                "stored_standardx": hilo(x / scale),
                "log_standardx": hilo(math.log(x) - math.log(scale)),
                "current_status": "OK", "current_value": hilo(0.5),
                "candidate_status": "OK", "candidate_value": hilo(0.5),
            })
    return rows


def chisquare_rows(k: int) -> list[dict]:
    """Chi-square reaches only two offsets: exact, or a 1/2 ULP tie. There is no
    continuum, because halving a binary64 either preserves the low bit or lands
    on a midpoint."""
    m = Fraction(2) ** -1074
    N = 2 ** (k - 1)
    rows = []
    for tag, construction, x_exact in (
        ("exact", "landmark", 2 * N * m),
        ("tie", "transform_stress", (2 * N + 1) * m),
    ):
        x = float(x_exact)
        assert Fraction(x) == x_exact, f"k={k} {tag}: X is not representable"
        rows.append({
            "family": "chisquare", "surface": "cumulative",
            "point_id": f"c_{tag}_k{k}", "construction": construction,
            "shape_id": "df1", "bucket_bits": k,
            "echo_x": hilo(x), "echo_df": hilo(1.0),
            "stored_standardx": hilo(0.5 * x),
            "log_standardx": hilo(math.log(x) - math.log(2.0)),
            "current_status": "OK", "current_value": hilo(0.5),
            "candidate_status": "OK", "candidate_value": hilo(0.5),
        })
    return rows


def write_csv(rows: list[dict], path: Path) -> Path:
    """Column set is family-specific: Gamma echoes Shape and ScaleParam,
    Chi-square echoes DegreesFreedom and standardises by a fixed 2. The two are
    different public inputs and are not forced into shared column names."""
    with path.open("w", newline="") as fh:
        writer = csv.DictWriter(fh, list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
    return path


BUCKETS = (52, 48, 44, 40, 36, 32, 28, 24, 20, 16, 12, 8, 4, 2, 1)


def valid_rows() -> list[dict]:
    """Gamma only. The loader rejects mixed families in one file, deliberately:
    they have different raw schemas and different transform mechanics."""
    rows: list[dict] = []
    for k in BUCKETS:
        rows += gamma_rows(k)
    return rows


def valid_chisquare_rows() -> list[dict]:
    rows: list[dict] = []
    for k in BUCKETS:
        rows += chisquare_rows(k)
    return rows


# ---------------------------------------------------------------------------
# tests
# ---------------------------------------------------------------------------

def test_hilo_round_trips_exactly():
    """Every value in the study is subnormal or near the normal floor, where a
    15-digit decimal alone cannot round-trip. The two-part token must."""
    for x in (0.25, 2.0 ** -1074, 2.0 ** -1023, 3.0e-308, 1.0, 0.0,
              2.0 ** -1024 + 2.0 ** -1074):
        assert parse_hilo(hilo(x)) == x, f"{x!r} did not round-trip"


def test_bits_available_matches_bucket_definition():
    """A stored subnormal with its leading bit at 2^-(1075-k) carries exactly k
    significand bits. This is the axis the whole study is reported against."""
    for k in BUCKETS:
        assert bits_available(2.0 ** -(1075 - k)) == k
    assert bits_available(1.0) == 53                    # normal
    assert bits_available(2.2250738585072014e-308) == 53  # min normal
    assert bits_available(0.0) == 0                     # hard underflow


def test_exact_standardx_is_reconstructed_not_read():
    """The exporter never serialises the exact quotient. The analyzer derives it
    as a rational from the echoed operands, which is what keeps the reference
    tied to the Double the function actually received."""
    m = Fraction(2) ** -1074
    N = 2 ** 31
    x = float((3 * N + 1) * m)
    p = Point("gamma", "cumulative", "t", "transform_stress", "s0.5", 32,
              x, 0.5, 3.0, x / 3.0, None, "OK", None, "OK", None)
    assert p.exact_standardx == Fraction(x) / 3
    # exactly one third of a ULP above the landmark, by construction
    assert p.exact_standardx - N * m == m / 3


def test_stress_offsets_land_where_declared():
    """1/3, 1/2 and 2/3 ULP for Gamma; 0 and 1/2 only for Chi-square."""
    m = Fraction(2) ** -1074
    for k in BUCKETS:
        N = 2 ** (k - 1)
        landmark = N * m
        seen = set()
        for row in gamma_rows(k)[::3]:                  # one surface per point
            p = load_one(row)
            seen.add(p.exact_standardx - landmark)
        assert seen == {Fraction(0), m / 3, m / 2, 2 * m / 3}, f"k={k}: {seen}"

        cs = {load_one(r).exact_standardx - landmark for r in chisquare_rows(k)}
        assert cs == {Fraction(0), m / 2}, f"k={k} chisquare: {cs}"


def test_transform_relative_error_tracks_two_to_the_minus_k():
    """Half a subnormal ULP over 2^(k-1) ULPs is about 2^-k. If this drifts, the
    reconstruction is wrong somewhere -- nothing in the exporter computes it."""
    for k in (52, 32, 8):
        row = gamma_rows(k)[3]                          # the 1/3 ULP stress point
        err = load_one(row).transform_relative_error
        assert err is not None
        assert 2.0 ** -(k + 2) < err < 2.0 ** -(k - 2), f"k={k}: {err}"


def test_reference_scale_metamorphic():
    """Hold the exact standardized argument and shape fixed; vary the scale.

    CDF and survival are functions of the standardized argument alone, so their
    references must be invariant. The density is a density of X, so its
    reference must scale exactly as 1/ScaleParam. Stating invariance and
    covariance together is what pins the Jacobian: a reference that omits it
    passes the CDF half and fails the density half.

    Spanned across two shapes as well as several scales, because shape is now
    part of evidence identity and a reference bug could be shape-dependent.
    """
    z = 0.125
    scales = (1.0, 2.0, 3.0, 64.0)

    for shape in (0.5, 1.0):
        for surface in ("cumulative", "survival"):
            refs = [reference(Point("gamma", surface, "m", "landmark", "sX",
                                    53, z * s, shape, s, z, None,
                                    "OK", None, "OK", None))
                    for s in scales]
            for r in refs[1:]:
                assert abs(r - refs[0]) < 1e-40, f"{surface} moved with scale"

        unit = reference(Point("gamma", "density", "m", "landmark", "sX", 53,
                               z, shape, 1.0, z, None, "OK", None, "OK", None))
        for s in scales:
            r = reference(Point("gamma", "density", "m", "landmark", "sX", 53,
                                z * s, shape, s, z, None,
                                "OK", None, "OK", None))
            assert abs(r * s - unit) < 1e-40, f"density did not scale as 1/{s}"


def test_runtime_oracle_invariant_passes():
    """The same property, checked by the analyzer on every run so that a future
    reference refactor cannot reintroduce the defect silently."""
    assert verify_reference_invariants() == []


def test_chisquare_kernel_shape_is_half_the_df():
    """Shape and DegreesFreedom are different public inputs. The analyzer
    normalises df/2 rather than pretending they are the same."""
    g = Point("gamma", "cumulative", "g", "landmark", "s0.5", 53,
              0.25, 0.5, 1.0, 0.25, None, "OK", None, "OK", None)
    c = Point("chisquare", "cumulative", "c", "landmark", "df1", 53,
              0.5, 1.0, 2.0, 0.25, None, "OK", None, "OK", None)
    assert g.kernel_shape == Fraction(1, 2)
    assert c.kernel_shape == Fraction(1, 2)
    assert g.exact_standardx == c.exact_standardx
    assert abs(reference(g) - reference(c)) < 1e-40


def test_output_classification_is_exhaustive():
    """Four classes, from the independent reference. exact_zero is kept apart
    from underflowed_output because a probability that is genuinely zero is a
    different claim from a positive one no Double can hold."""
    from mpmath import mpf
    assert classify_output(mpf("0.5")) == "normal_output"
    assert classify_output(mpf(2) ** -1022) == "normal_output"   # min normal
    assert classify_output(mpf(2) ** -1023) == "subnormal_output"
    assert classify_output(mpf(2) ** -1074) == "subnormal_output"
    assert classify_output(mpf(2) ** -1080) == "underflowed_output"
    assert classify_output(mpf(0)) == "exact_zero"
    assert classify_output(None) == "no_reference"


def test_output_limited_rows_do_not_choose_the_dispatch_boundary():
    """The CDF confound, reproduced synthetically.

    At the same bucket, Shape = 1 gives a subnormal true output while
    Shape = 0.75 gives a normal one. The subnormal row cannot be recovered by
    any branch, so including it in the envelope can invert the crossover. The
    public envelope must see it; the algorithmic envelope must not.
    """
    m = Fraction(2) ** -1074
    N = 2 ** 31
    x = float((3 * N + 1) * m)

    def pt(shape, shape_id, cur, cand):
        return Point("gamma", "cumulative", "p", "transform_stress", shape_id,
                     32, x, shape, 3.0, x / 3.0, None,
                     "OK", cur, "OK", cand)

    # Shape = 1: P(1,z) ~ z, so the true output is itself subnormal
    limited = pt(1.0, "s1.0", 1e-314, 1e-314)
    assert classify_output(reference(limited)) == "subnormal_output"
    # Shape = 0.75: the true output is comfortably normal
    algo = pt(0.75, "s0.75", 1e-20, 1e-20)
    assert classify_output(reference(algo)) == "normal_output"

    both = [limited, algo]
    assert 32 in envelope(both, normal_only=False)
    alg = envelope(both, normal_only=True)
    assert alg[32][2] != "s1.0" and alg[32][3] != "s1.0", \
        "an output-limited row bound the algorithmic envelope"


def test_missing_candidate_cannot_win_a_bucket():
    """Failing to answer is not the same as answering accurately.

    If the candidate errored at every point in a bucket, its worst error must
    be reported as unavailable rather than defaulting to zero and appearing to
    beat a current path that did produce values.
    """
    from analyze_positive_ratio_subnormal import crossover_of
    m = Fraction(2) ** -1074
    N = 2 ** 31
    x = float((3 * N + 1) * m)
    p = Point("gamma", "density", "p", "transform_stress", "s0.001", 32,
              x, 0.001, 3.0, x / 3.0, None, "OK", 1.0, "ERROR", None)
    w = envelope([p], normal_only=False)
    assert w[32][1] is None, "a missing candidate was recorded as zero error"
    cross, _, _ = crossover_of(w)
    assert cross is None, "a bucket with no candidate value reported a crossover"


def test_overflowed_output_is_its_own_class():
    """A reference above DoubleMax must not be called normal_output.

    float(ref) collapses everything above the Double range to infinity, and inf
    satisfies `>= MIN_NORMAL`, so the naive test classified 139 overflowing
    density rows across the two committed arms as normal. The bound is compared
    exactly, before conversion.
    """
    from analyze_positive_ratio_subnormal import DOUBLE_MAX
    from mpmath import mpf
    dmax = mpf(DOUBLE_MAX.numerator) / mpf(DOUBLE_MAX.denominator)
    assert classify_output(dmax) == "normal_output"            # exactly DoubleMax
    assert classify_output(dmax * (1 + mpf(2) ** -52)) == "overflowed_output"
    assert classify_output(mpf(10) ** 400) == "overflowed_output"


def test_tiny_shape_density_overflows_in_both_families():
    """The real case, one per family: a tiny-shape density whose true value is
    above the Double range. Neither branch can represent it, so it must be
    excluded from the algorithmic envelope rather than counted as normal."""
    z = 2.0 ** -1051
    g = Point("gamma", "density", "g", "transform_stress", "s0.001", 24,
              z, 0.001, 1.0, z, None, "ERROR", None, "ERROR", None)
    c = Point("chisquare", "density", "c", "transform_stress", "s0.002", 24,
              z * 2, 0.002, 2.0, z, None, "ERROR", None, "ERROR", None)
    for p in (g, c):
        assert classify_output(reference(p)) == "overflowed_output", p.family
        assert envelope([p], normal_only=True) == {}, \
            f"{p.family}: an overflowing row entered the algorithmic envelope"


def test_crossover_reports_both_binding_shapes():
    """Worst current and worst candidate routinely bind at different shapes, so
    one generic 'binding shape' is ambiguous and produced contradictory numbers
    for the same crossover in two reports."""
    w = {32: [1e-10, 1e-13, "s1.0", "s0.1"],
         28: [1e-9, 1e-13, "s1.0", "s0.1"]}
    cross, binds_current, binds_candidate = crossover_of(w)
    assert cross == 32
    assert binds_current == "s1.0"
    assert binds_candidate == "s0.1"


# Imported under an alias: the runner collects globals starting with "test_",
# and the analyzer's own test_claims() takes an argument, so importing it under
# its real name makes the runner try to call it as a test case.
from analyze_positive_ratio_subnormal import (      # noqa: E402
    test_claims as run_claim_test,
)


def test_registered_claims_match_the_design_document():
    """The claims are pre-registered in HOLDOUT_DESIGN.md. If this table and
    that document disagree, the holdout is testing something nobody agreed
    to."""
    assert REGISTERED_CLAIMS[("gamma", "density")] == 48
    assert REGISTERED_CLAIMS[("gamma", "cumulative")] == 40
    assert REGISTERED_CLAIMS[("chisquare", "density")] == 48
    assert REGISTERED_CLAIMS[("chisquare", "cumulative")] == 48
    assert ("gamma", "survival") not in REGISTERED_CLAIMS
    assert ("chisquare", "survival") not in REGISTERED_CLAIMS


def _claim_point(surface, bits, shape, cur, cand):
    m = Fraction(2) ** -1074
    N = 2 ** (bits - 1)
    x = float((3 * N + 1) * m)
    return Point("gamma", surface, f"p{bits}", "transform_stress", "sX",
                 bits, x, shape, 3.0, x / 3.0, None, "OK", cur, "OK", cand)


def test_claim_passes_when_candidate_dominates_below_the_cutoff(capsys=None):
    """A claim holds when the candidate wins at the claimed bucket and every
    bucket below. Buckets above the claim are not tested — the claim says
    nothing about them."""
    pts = [_claim_point("cumulative", b, 0.5, 1e-10, 1e-14)
           for b in (40, 33, 25, 17, 9, 3)]
    pts += [_claim_point("cumulative", 50, 0.5, 1e-14, 1e-10)]   # above; ignored
    assert run_claim_test(pts) is True


def test_claim_fails_when_a_tested_bucket_violates_it():
    """A single violating bucket at or below the cutoff rejects the claim. The
    verdict must not be softened by the buckets that do pass."""
    pts = [_claim_point("cumulative", b, 0.5, 1e-10, 1e-14)
           for b in (40, 33, 25, 17, 9)]
    pts += [_claim_point("cumulative", 3, 0.5, 1e-14, 1e-10)]    # violates
    assert run_claim_test(pts) is False


def test_claim_mode_proposes_no_crossover():
    """The holdout must not offer a replacement cutoff. Adopting one would make
    it a second fitting set and leave no independent evidence."""
    import io as _io, contextlib
    pts = [_claim_point("cumulative", b, 0.5, 1e-14, 1e-10)
           for b in (40, 33, 25)]
    buf = _io.StringIO()
    with contextlib.redirect_stdout(buf):
        run_claim_test(pts)
    out = buf.getvalue().lower()
    assert "crossover" not in out or "no crossover is proposed" in out
    assert "bits and below" not in out


def test_unavailable_claimed_bucket_does_not_pass():
    """Fail closed: a claimed bucket with no comparable evidence must not
    silently vanish from the denominator.

    Skipping it would let the remaining buckets carry a PASS on a claim that
    was never established there — the opposite of what a holdout is for.
    """
    pts = [_claim_point("cumulative", b, 0.5, 1e-10, 1e-14)
           for b in (40, 33, 25, 17, 9)]
    # a claimed bucket where the candidate produced nothing at all
    pts.append(_claim_point("cumulative", 3, 0.5, 1e-10, None))
    assert run_claim_test(pts) is False

    # and the mirror case: current unavailable
    pts2 = [_claim_point("density", b, 0.5, 1e-10, 1e-14)
            for b in (48, 41, 33, 25, 17, 9, 3)]
    pts2.append(_claim_point("density", 46, 0.5, None, 1e-14))
    assert run_claim_test(pts2) is False


def test_cli_exits_nonzero_when_a_claim_is_rejected():
    """The verdict must reach the process. A hosted gate running this step
    would otherwise go green on evidence that rejects the claim."""
    import csv as _csv
    with tempfile.TemporaryDirectory() as tmp:
        rows = valid_rows()
        # force a violation inside the claimed range by making the candidate
        # worse than current everywhere
        for r in rows:
            r["current_value"] = hilo(1e-16)
            r["candidate_value"] = hilo(1.0)
        path = write_csv(rows, Path(tmp) / "reject.csv")
        result = subprocess.run(
            [sys.executable, str(ANALYZER), str(path), "--claims"],
            capture_output=True, text=True)
        assert result.returncode != 0, \
            "a rejected claim exited 0; the gate would go green"
        assert "REJECTED" in result.stdout or "PENDING" in result.stdout


def test_hard_underflow_is_not_bucket_zero():
    """A stored zero is not merely imprecise: it has become indistinguishable
    from the support boundary. It must be classified apart, never pooled into
    the bucket axis."""
    p = Point("gamma", "cumulative", "u", "underflow", "s0.5", 0,
              1e-320, 0.5, 1e300, 0.0, None, "OK", None, "OK", None)
    assert p.stored_is_zero
    assert p.exact_standardx > 0          # the mathematical ratio is positive


def test_valid_evidence_passes_both_checks():
    for rows in (valid_rows(), valid_chisquare_rows()):
        points = [load_one(r) for r in rows]
        assert verify_constructions(points) == []
        assert twin_check(points) == []


def test_mixed_families_in_one_file_are_rejected():
    """Gamma and Chi-square have different raw schemas. A file containing both
    could only be read by guessing, so it is refused."""
    with tempfile.TemporaryDirectory() as tmp:
        rows = gamma_rows(32)
        rows.append(dict(rows[0], family="chisquare"))
        path = write_csv(rows, Path(tmp) / "mixed.csv")
        try:
            load(path)
        except SystemExit as exc:
            assert "mixed families" in str(exc)
        else:
            raise AssertionError("a mixed-family file was accepted")


def test_k1_stress_may_round_into_the_two_bit_bucket():
    """A 1-bit stress point sits half a ULP above 2^-1074, so round-to-even
    carries it up to the 2-bit grid point. That is the rounding behaviour under
    test, not a mislabelled point -- landmarks must stay put, stress points may
    move one bucket up."""
    rows = [r for r in gamma_rows(1) if r["point_id"] == "g_s12_k1"]
    points = [load_one(r) for r in rows]
    assert bits_available(points[0].stored) == 2
    assert verify_constructions(points) == []


def test_mislabelled_bucket_is_rejected():
    rows = gamma_rows(52)
    for r in rows:
        r["bucket_bits"] = 44
    problems = verify_constructions([load_one(r) for r in rows])
    assert problems, "a landmark 8 buckets off its boundary was accepted"


def test_chisquare_cannot_reach_a_third_ulp_offset():
    """Stronger than a guard: the offset is physically unconstructible.

    For a binary64 X, X/2 is either exact or a midpoint tie. Aiming at a 1/3
    ULP quotient requires an X that is not representable, and rounding it lands
    on exactly 1/2. Gamma reaches 1/3 only because it can divide by 3.
    """
    m = Fraction(2) ** -1074
    N = 2 ** 31
    target = N * m + m / 3
    x_needed = target * 2
    assert Fraction(float(x_needed)) != x_needed, "unexpectedly representable"
    landed = Fraction(float(x_needed)) / 2 - N * m
    assert landed / m == Fraction(1, 2), f"landed at {landed / m} ULP"


def test_impossible_stress_offset_is_rejected():
    """The guard itself, exercised with a row a malformed exporter could emit:
    a Chi-square point claiming a 1/3 ULP offset. Unreachable through valid
    data, which is exactly why the guard has to be checked directly."""
    m = Fraction(2) ** -1074
    N = 2 ** 31
    x = float((3 * N + 1) * m)
    bad = Point("chisquare", "cumulative", "c_bad_k32", "transform_stress",
                "df1", 32, x, 1.0, 3.0, x / 3.0, None,
                "OK", None, "OK", None)
    problems = verify_constructions([bad])
    assert problems, "a 1/3 ULP Chi-square offset was accepted"
    assert "not one of" in problems[0]


def test_analyzer_exits_nonzero_on_bad_evidence():
    """Fail closed. A crossover derived from a point that is not where it claims
    to be would corrupt every bucket it appears in."""
    with tempfile.TemporaryDirectory() as tmp:
        rows = valid_rows()
        rows[0]["bucket_bits"] = 44
        path = write_csv(rows, Path(tmp) / "bad.csv")
        result = subprocess.run([sys.executable, str(ANALYZER), str(path)],
                                capture_output=True, text=True)
        assert result.returncode != 0
        assert "CONSTRUCTION CHECK FAILED" in result.stdout


def test_analyzer_accepts_valid_evidence():
    with tempfile.TemporaryDirectory() as tmp:
        path = write_csv(valid_rows(), Path(tmp) / "good.csv")
        result = subprocess.run([sys.executable, str(ANALYZER), str(path)],
                                capture_output=True, text=True)
        assert result.returncode == 0, result.stdout + result.stderr
        assert "construction check passed" in result.stdout


def test_missing_column_is_rejected():
    with tempfile.TemporaryDirectory() as tmp:
        rows = valid_rows()
        for r in rows:
            del r["log_standardx"]
        path = Path(tmp) / "short.csv"
        with path.open("w", newline="") as fh:
            cols = [c for c in COLUMNS if c != "log_standardx"]
            w = csv.DictWriter(fh, cols)
            w.writeheader()
            w.writerows(rows)
        try:
            load(path)
        except SystemExit as exc:
            assert "log_standardx" in str(exc)
        else:
            raise AssertionError("a missing column was accepted")


# ---------------------------------------------------------------------------

def load_one(row: dict) -> Point:
    with tempfile.TemporaryDirectory() as tmp:
        return load(write_csv([row], Path(tmp) / "one.csv"))[0]


def main() -> int:
    tests = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    failed = 0
    for t in tests:
        try:
            t()
            print(f"  PASS  {t.__name__}")
        except AssertionError as exc:
            failed += 1
            print(f"  FAIL  {t.__name__}: {exc}")
    print(f"\n{len(tests) - failed}/{len(tests)} passed")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
