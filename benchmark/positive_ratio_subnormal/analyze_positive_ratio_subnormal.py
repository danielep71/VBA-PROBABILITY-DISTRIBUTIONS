#!/usr/bin/env python3
"""
analyze_positive_ratio_subnormal.py
===================================

Phase A analyzer for ICR-P1-01A (#13).

WHAT THIS MEASURES
------------------
A mathematically positive standardized argument can become a low-precision
subnormal, or zero, during the scale transformation -- after which downstream
code treats the rounded transformed value as an exact support coordinate.

The transformation itself is NOT defective. IEEE-754 guarantees `X / Scale` and
`0.5 * X` are correctly rounded, and measurement confirms it. The defect is one
of composition and representation: information that the mathematical ratio was
positive but unrepresentable is silently lost.

The two families reach that state by different routes and are measured
separately:

    Gamma       PROB_TryDivide(X, ScaleParam, StandardX)   -- division
    Chi-square  0.5 * X, inline, with no Try wrapper       -- halving

Their rounding structures differ in kind, not degree. Gamma's transform error
is graded: the exact quotient can fall anywhere between two representable
subnormals. Chi-square's is binary: halving a binary64 X is either exact or
lands exactly on a midpoint. Pooling them would hide that, so every report is
produced per family first and combined only at the end.

REFERENCE DISCIPLINE
--------------------
VBA never serialises the exact quotient. It echoes the binary64 operands it
actually received, and this analyzer reconstructs the exact mathematical
standardized value as a RATIONAL from those echoed operands. That keeps the
reference-at-actual-input rule established under #17: the reference is
evaluated at the Double the function received, not at the decimal it was
written from.

INPUT SCHEMA (long form, one row per public surface)
----------------------------------------------------
    family              gamma | chisquare
    surface             density | cumulative | survival
    point_id            stable identifier, unique within family
    construction        landmark | transform_stress | decimal_twin | underflow
    bucket_bits         significand bits available in the target quotient
    echo_x              public X as received, hi;lo
    echo_param          ScaleParam (gamma) or DegreesFreedom (chisquare), hi;lo
    stored_standardx    what VBA stored for the standardized argument, hi;lo
    log_standardx       Log(X) - Log(param) as VBA formed it, hi;lo
    current_status      OK | ERROR | CVERR
    current_value       current public result, hi;lo
    candidate_status    OK | ERROR | CVERR
    candidate_value     log-path mirror result, hi;lo

DERIVED HERE
------------
    exact_standardx             exact rational from the echoed operands
    stored_error_ulps           stored vs exact, in subnormal ULPs
    transform_relative_error    |stored - exact| / exact
    stored_is_zero              hard-underflow class, qualitatively different
    log_path_error              exact log(z) vs the log VBA formed
    current_error               vs high-precision reference
    candidate_error             vs high-precision reference
    winner                      current | candidate | tie

Phase A is measurement only. It proposes no cutoff and changes no source. The
crossover constants quoted in #13 (2^-1030 CDF, 2^-1038 survival, 2^-1022
density) are treated as HYPOTHESES until this analyzer reports otherwise, and
are not assumed shared between the two families.
"""

from __future__ import annotations

import argparse
import csv
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from fractions import Fraction
from pathlib import Path

try:
    from mpmath import mp, mpf, gammainc, loggamma, log as mp_log, exp as mp_exp
except ImportError:                                        # pragma: no cover
    sys.exit("mpmath is required: pip install mpmath")

# The repository's grid generator uses _ORACLE_DPS = 110. This study uses 60,
# which is a deliberate choice rather than an oversight: the two agree to
# 4.3E-62 on the study's own points, against a smallest measured error of about
# 1E-16 -- a margin of 2E+45. Raising it changes no reported figure and no
# crossover, while roughly doubling a run over 1,862 observations.
#
# ORACLE_STABILITY_DPS below is checked at startup, so the claim is verified on
# every run rather than asserted here.
mp.dps = 60
ORACLE_STABILITY_DPS = 110
ORACLE_STABILITY_TOLERANCE = 1e-40

MIN_SUBNORMAL = Fraction(2) ** -1074       # spacing of every subnormal
MIN_NORMAL = Fraction(2) ** -1022
SURFACES = ("density", "cumulative", "survival")
FAMILIES = ("gamma", "chisquare")

COMMON = ["family", "surface", "point_id", "construction", "shape_id",
          "bucket_bits", "echo_x", "stored_standardx", "log_standardx",
          "current_status", "current_value",
          "candidate_status", "candidate_value"]
GAMMA_ONLY = ["echo_shape", "echo_scale"]
CHISQUARE_ONLY = ["echo_df"]
COLUMNS = COMMON + GAMMA_ONLY          # gamma is the default raw shape


# ---------------------------------------------------------------------------
# binary64 helpers
# ---------------------------------------------------------------------------

def parse_hilo(token: str) -> float | None:
    """Recover the exact Double from a "hi;lo" export token.

    The exporter writes 15 significant digits plus the residual, so summing the
    two parts in Double precision reproduces the original bit pattern. An empty
    field means the value was not produced -- an error row, not a zero.
    """
    token = (token or "").strip()
    if not token or ";" not in token:
        return None
    hi, lo = token.split(";", 1)
    return float(hi) + float(lo)


def bits_available(x: float) -> int:
    """Significand bits a stored Double actually carries.

    A normal carries the full 53. A subnormal carries only the bits between its
    leading one and the 2^-1074 floor, which is the whole point of the bucket
    axis: a subnormal at 2^-1043 has 32 bits, not 53, and no downstream
    algorithm can recover what was never stored.
    """
    if x == 0.0:
        return 0
    f = Fraction(x)
    if abs(f) >= MIN_NORMAL:
        return 53
    n = abs(f) / MIN_SUBNORMAL          # exact integer count of subnormal ULPs
    return int(n).bit_length()


def ulps_between(a: Fraction, b: Fraction) -> Fraction:
    """Separation in subnormal ULPs. Meaningful because subnormal spacing is
    uniform, so a single unit applies across the whole range being measured."""
    return abs(a - b) / MIN_SUBNORMAL


# ---------------------------------------------------------------------------
# records
# ---------------------------------------------------------------------------

@dataclass
class Point:
    family: str
    surface: str
    point_id: str
    construction: str
    shape_id: str
    bucket_bits: int
    x: float
    shape_input: float          # echo_shape for gamma, echo_df for chi-square
    scale: float                # ScaleParam for gamma, always 2 for chi-square
    stored: float | None
    log_stored: float | None
    current_status: str
    current: float | None
    candidate_status: str
    candidate: float | None

    @property
    def kernel_shape(self) -> Fraction:
        """The shape the downstream kernel receives. Gamma passes Shape
        through; Chi-square halves the degrees of freedom. Normalising here
        keeps the public-input provenance explicit: Shape and DegreesFreedom
        are different public inputs and are not pretended to be the same."""
        if self.family == "gamma":
            return Fraction(self.shape_input)
        return Fraction(self.shape_input) / 2

    # -- the exact mathematical standardized argument -----------------------
    @property
    def exact_standardx(self) -> Fraction:
        """Reconstructed as a rational from the echoed binary64 operands, never
        read back from the exporter -- that is what keeps the reference tied to
        the Double the function actually received."""
        return Fraction(self.x) / Fraction(self.scale)

    @property
    def stored_is_zero(self) -> bool:
        """Hard underflow. Qualitatively different from a positive subnormal:
        the value is not merely imprecise, it has become indistinguishable from
        the support boundary."""
        return self.stored == 0.0

    @property
    def stored_error_ulps(self) -> Fraction | None:
        if self.stored is None:
            return None
        return ulps_between(Fraction(self.stored), self.exact_standardx)

    @property
    def transform_relative_error(self) -> float | None:
        """At the lower edge of a k-bit bucket this is bounded by about 2^-k:
        half a subnormal ULP over 2^(k-1) ULPs."""
        if self.stored is None or self.exact_standardx == 0:
            return None
        return float(abs(Fraction(self.stored) - self.exact_standardx)
                     / self.exact_standardx)

    @property
    def log_path_error(self) -> float | None:
        """Exact log of the exact ratio, against the log VBA actually formed.

        Kept separate from the transform error on purpose. #13 anticipates the
        log branch being limited near 1E-12 by subtracting logs of magnitude
        ~700; this is the field that measures whether that is true rather than
        assumed.
        """
        if self.log_stored is None or self.exact_standardx <= 0:
            return None
        exact = mp_log(mpf(self.exact_standardx.numerator)
                       / mpf(self.exact_standardx.denominator))
        if exact == 0:
            return None
        return float(abs(mpf(self.log_stored) - exact) / abs(exact))


# ---------------------------------------------------------------------------
# references
# ---------------------------------------------------------------------------

def reference(point: Point) -> mpf | None:
    """High-precision value of the surface at the EXACT standardized argument
    and the EXACT kernel shape -- not at the stored ones. The gap between them
    is the defect under measurement."""
    z = point.exact_standardx
    sh = point.kernel_shape
    if z <= 0 or sh <= 0:
        return None
    zz = mpf(z.numerator) / mpf(z.denominator)
    a = mpf(sh.numerator) / mpf(sh.denominator)

    if point.surface == "density":
        # The Jacobian of the standardisation. The public function returns the
        # density of X, not of the standardized variable, so the standardized
        # density must be divided by the scale. Omitting it produces a constant
        # relative error of 1 - 1/scale at every point sharing that scale --
        # which is a defect in the reference, not in the kernel.
        scale = Fraction(point.scale)
        jacobian = mpf(scale.numerator) / mpf(scale.denominator)
        return mp_exp((a - 1) * mp_log(zz) - zz - loggamma(a)) / jacobian
    if point.surface == "cumulative":
        return gammainc(a, 0, zz, regularized=True)
    return gammainc(a, zz, mp.inf, regularized=True)


def classify_output(ref: mpf | None) -> str:
    """Where the TRUE answer sits on the binary64 lattice.

    Classified from the independent high-precision reference, never from an
    observed result -- an observation cannot be evidence about its own
    representability.

    The distinction matters because a row whose true output is itself subnormal
    must round back onto the same lattice the input lost, so no branch can
    recover information there. Such rows legitimately show a relative error of
    the same order as the transform error, and pooling them with algorithmically
    limited rows inverts the measured crossover. They are labelled, counted and
    reported -- never silently dropped.
    """
    if ref is None:
        return "no_reference"
    if ref == 0:
        return "exact_zero"                 # not a positive value that underflowed
    rounded = float(ref)
    if rounded == 0.0:
        return "underflowed_output"         # positive, but no Double can hold it
    if abs(rounded) >= float(MIN_NORMAL):
        return "normal_output"
    return "subnormal_output"


def relative_error(value: float | None, ref: mpf | None) -> float | None:
    if value is None or ref is None or ref == 0:
        return None
    return float(abs(mpf(value) - ref) / abs(ref))


# ---------------------------------------------------------------------------
# loading
# ---------------------------------------------------------------------------

def load(path: Path) -> list[Point]:
    rows = list(csv.DictReader(path.open(newline="")))
    if not rows:
        raise SystemExit(f"{path}: no data rows")

    family = rows[0].get("family")
    if family not in FAMILIES:
        raise SystemExit(f"{path}: unknown family {family!r}")

    required = COMMON + (GAMMA_ONLY if family == "gamma" else CHISQUARE_ONLY)
    missing = [c for c in required if c not in rows[0]]
    if missing:
        raise SystemExit(f"{path}: missing columns {missing}")

    points: list[Point] = []
    for i, r in enumerate(rows, 2):
        if r["family"] != family:
            raise SystemExit(f"{path} line {i}: mixed families in one file")
        if r["surface"] not in SURFACES:
            raise SystemExit(f"{path} line {i}: unknown surface {r['surface']!r}")

        x = parse_hilo(r["echo_x"])
        if family == "gamma":
            shape_input = parse_hilo(r["echo_shape"])
            scale = parse_hilo(r["echo_scale"])
        else:
            shape_input = parse_hilo(r["echo_df"])
            scale = 2.0
        if x is None or shape_input is None or scale is None:
            raise SystemExit(f"{path} line {i}: an echoed input is missing -- "
                             f"every row must record what it was given")

        points.append(Point(
            family=family, surface=r["surface"], point_id=r["point_id"],
            construction=r["construction"], shape_id=r["shape_id"],
            bucket_bits=int(r["bucket_bits"]),
            x=x, shape_input=shape_input, scale=scale,
            stored=parse_hilo(r["stored_standardx"]),
            log_stored=parse_hilo(r["log_standardx"]),
            current_status=r["current_status"],
            current=parse_hilo(r["current_value"]),
            candidate_status=r["candidate_status"],
            candidate=parse_hilo(r["candidate_value"]),
        ))
    return points


# ---------------------------------------------------------------------------
# checks that must pass before any number is believed
# ---------------------------------------------------------------------------

def verify_reference_invariants() -> list[str]:
    """A runtime invariant on the oracle itself, independent of any observation.

    Holding the exact standardized argument and shape fixed while varying the
    scale, the CDF and survival references must not move, and the density
    reference must scale exactly as 1/scale. Checked on every run because the
    omitted-Jacobian defect was invisible in the reported numbers: it looked
    like a catastrophic kernel error, uniformly, at every bucket.
    """
    def pt(surface, s_):
        return Point("gamma", surface, "inv", "landmark", "inv", 53,
                     0.125 * s_, 1.0, s_, 0.125, None, "OK", None, "OK", None)

    problems = []
    for surface in ("cumulative", "survival"):
        base = reference(pt(surface, 1.0))
        for s_ in (2.0, 3.0, 64.0):
            if abs(reference(pt(surface, s_)) - base) > mpf("1e-40"):
                problems.append(f"{surface} reference moved with the scale "
                                f"at s={s_}")
    # Oracle stability: the reference must not move materially when computed at
    # the repository's own 110-dps standard. A single-oracle study is only as
    # good as that oracle, and self-consistency at one precision proves
    # convergence, not correctness -- but a value that shifts under added
    # precision is disqualifying regardless.
    for surface, shape in (("cumulative", 0.0001), ("cumulative", 1.0),
                           ("survival", 0.001), ("density", 0.5)):
        p60 = Point("gamma", surface, "dps", "landmark", "dps", 32,
                    2.0 ** -1043, shape, 1.0, 2.0 ** -1043, None,
                    "OK", None, "OK", None)
        lo = reference(p60)
        saved = mp.dps
        mp.dps = ORACLE_STABILITY_DPS
        hi = reference(p60)
        mp.dps = saved
        if lo is None or hi is None or hi == 0:
            continue
        if abs(lo - hi) / abs(hi) > ORACLE_STABILITY_TOLERANCE:
            problems.append(f"{surface} reference at shape {shape} moved "
                            f"between {saved} and {ORACLE_STABILITY_DPS} dps")

    unit = reference(pt("density", 1.0))
    for s_ in (2.0, 3.0, 64.0):
        if abs(reference(pt("density", s_)) * s_ - unit) > mpf("1e-40"):
            problems.append(f"density reference did not scale as 1/{s_}")
    return problems


def verify_constructions(points: list[Point]) -> list[str]:
    """The constructions are exact by design, so a mismatch means the exporter
    or the VBE parser altered something. Better to fail loudly than to report a
    crossover derived from a point that is not where it claims to be."""
    problems: list[str] = []
    for p in points:
        if p.construction in ("decimal_twin", "underflow"):
            continue    # no bucket identity: diagnostic and recovery classes

        declared = p.bucket_bits
        landmark = Fraction(2) ** (declared - 1) * MIN_SUBNORMAL
        offset = p.exact_standardx - landmark

        if p.construction == "landmark" and offset != 0:
            problems.append(f"{p.point_id}: landmark is {offset / MIN_SUBNORMAL} "
                            f"ULP off its bucket boundary")

        if p.construction == "transform_stress":
            frac = offset / MIN_SUBNORMAL
            allowed = ({Fraction(1, 3), Fraction(1, 2), Fraction(2, 3)}
                       if p.family == "gamma" else {Fraction(1, 2)})
            if frac not in allowed:
                problems.append(f"{p.point_id}: stress offset {frac} ULP is not "
                                f"one of {sorted(allowed)} for {p.family}")

        if p.stored is not None and not p.stored_is_zero:
            got = bits_available(p.stored)
            # A landmark is exactly representable, so it must land in its own
            # bucket. A stress point sits a fraction of a ULP above the
            # landmark and rounds to one of the two neighbouring grid points --
            # at the smallest buckets that can carry it up into the next
            # bucket, which is the rounding behaviour under test, not an error.
            allowed = {declared} if p.construction == "landmark" \
                else {declared, declared + 1}
            if got not in allowed:
                problems.append(f"{p.point_id}: stored value carries {got} bits, "
                                f"bucket declares {declared}")
    return problems


def twin_check(points: list[Point]) -> list[str]:
    """A decimal twin exists solely to prove the VBE parser did not alter a
    value. If a twin disagrees with its constructed partner, the fault is in
    parsing, not in the kernel -- and every other number here is suspect."""
    by_key = {(p.shape_id, p.point_id, p.surface): p for p in points}
    problems = []
    for p in points:
        if p.construction != "decimal_twin":
            continue
        partner = by_key.get((p.shape_id,
                              p.point_id.replace("twin_", ""), p.surface))
        if partner is None:
            problems.append(f"{p.point_id}: no constructed partner to compare")
        elif p.x != partner.x:
            problems.append(f"{p.shape_id}/{p.point_id}: parsed to {p.x!r}, "
                            f"constructed partner is {partner.x!r}")
    return problems


# ---------------------------------------------------------------------------
# reporting
# ---------------------------------------------------------------------------

def errors_for(p: Point) -> tuple[float | None, float | None]:
    ref = reference(p)
    return relative_error(p.current, ref), relative_error(p.candidate, ref)


def is_saturated(slice_points: list[Point]) -> bool:
    """A slice where every observation returns the same value on both branches
    carries no information about dispatch. That is a real characterisation
    result -- the public value is already correctly saturated -- and must be
    classified rather than treated as fitting evidence for a crossover."""
    cur = {p.current for p in slice_points if p.current is not None}
    can = {p.candidate for p in slice_points if p.candidate is not None}
    return len(cur) <= 1 and len(can) <= 1


def envelope(surf: list[Point], normal_only: bool) -> dict:
    """Worst current and candidate error per bucket, over stress points.

    `normal_only` selects the ALGORITHMIC envelope: rows whose true output is
    normally representable, which is where a branch change can actually recover
    information. Without it the PUBLIC envelope is returned, which is what a
    Double-returning UDF really delivers.
    """
    # Worsts start as None, not 0.0. A bucket where the candidate never
    # produced a value must not appear to win with a worst error of zero --
    # failing to answer is not the same as answering accurately.
    w: dict[int, list] = defaultdict(lambda: [None, None, None, None])
    for p in surf:
        if p.construction != "transform_stress":
            continue
        ref = reference(p)
        if normal_only and classify_output(ref) != "normal_output":
            continue
        ec, ea = errors_for(p)
        x = w[p.bucket_bits]
        if ec is not None and (x[0] is None or ec > x[0]):
            x[0], x[2] = ec, p.shape_id
        if ea is not None and (x[1] is None or ea > x[1]):
            x[1], x[3] = ea, p.shape_id
    return w


def crossover_of(w: dict) -> tuple[int | None, str | None]:
    """Lowest bucket from which the candidate wins every bucket below it."""
    cross = binds = None
    for bits in sorted(w, reverse=True):
        wc, wa, _sc, sa = w[bits]
        if wc is None or wa is None:
            cross = binds = None          # cannot compare; not a crossover
            continue
        if wa < wc:
            if cross is None:
                cross, binds = bits, sa
        else:
            cross = binds = None
    return cross, binds


def print_envelope(title: str, w: dict) -> None:
    print(f"\n   {title}")
    print(f"   {'bits':>5} {'worst current':>14} {'binds':>9} "
          f"{'worst candidate':>16} {'binds':>9}  better")
    for bits in sorted(w, reverse=True):
        wc, wa, sc, sa = w[bits]
        if wc is None or wa is None:
            verdict = ("candidate unavailable" if wa is None
                       else "current unavailable")
        else:
            verdict = "candidate" if wa < wc else "current"
        fc = "n/a" if wc is None else f"{wc:.3g}"
        fa = "n/a" if wa is None else f"{wa:.3g}"
        print(f"   {bits:>5} {fc:>14} {str(sc):>9} "
              f"{fa:>16} {str(sa):>9}  {verdict}")


def decomposition(surf: list[Point]) -> None:
    """Landmark versus stress, per shape and bucket.

    A landmark carries zero transform error, so any error there is downstream
    kernel structure. The ratio therefore separates the two mechanisms: near 1
    means the transform is not what is hurting, and a large ratio means the
    quotient representation is binding. The numbers are reported; no causal
    label is assigned from an arbitrary threshold.
    """
    shapes = sorted({p.shape_id for p in surf})
    print(f"\n   LANDMARK vs STRESS (current path; landmark has no transform "
          f"error)")
    print(f"   {'shape':>9} {'bits':>5} {'landmark':>12} {'stress':>12} "
          f"{'ratio':>14}")
    for sid in shapes:
        for bits in (52, 32, 8, 1):
            lm = [p for p in surf if p.shape_id == sid and p.bucket_bits == bits
                  and p.construction == "landmark"]
            st = [p for p in surf if p.shape_id == sid and p.bucket_bits == bits
                  and p.construction == "transform_stress"]
            if not lm or not st:
                continue
            el = errors_for(lm[0])[0]
            stress = [errors_for(p)[0] for p in st]
            es = max((e for e in stress if e is not None), default=None)
            if el is None or es is None:
                continue
            ratio = (es / el) if el else float("inf")
            print(f"   {sid:>9} {bits:>5} {el:>12.3g} {es:>12.3g} "
                  f"{ratio:>14.4g}")


def report(points: list[Point]) -> None:
    family = points[0].family
    print(f"\n{'=' * 78}\n{family.upper()}  ({len(points)} observations)"
          f"\n{'=' * 78}")

    for surface in SURFACES:
        surf = [p for p in points if p.surface == surface]
        if not surf:
            continue
        print(f"\n{'#' * 78}\n# {surface.upper()}\n{'#' * 78}")

        counts = Counter(classify_output(reference(p)) for p in surf)
        print(f"\n   rows {len(surf)}")
        for cls in ("normal_output", "subnormal_output",
                    "underflowed_output", "exact_zero", "no_reference"):
            if counts[cls]:
                print(f"   {cls:<22} {counts[cls]}")

        for shape_id in sorted({p.shape_id for p in surf}):
            sl = [p for p in surf if p.shape_id == shape_id
                  and p.construction in ("landmark", "transform_stress")]
            if not sl:
                continue
            if is_saturated(sl):
                print(f"\n-- shape {shape_id}: NON_DISCRIMINATING_SATURATION")
                print("   Every observation returns the same value on both "
                      "branches. The public\n   result is already correctly "
                      "saturated here, so this slice cannot\n   constrain "
                      "dispatch. Retained as characterisation, not fitting.")
                continue
            print(f"\n-- shape {shape_id} " + "-" * max(4, 58 - len(shape_id)))
            print(f"{'bits':>5} {'construction':>17} {'out class':>18} "
                  f"{'transform rel':>14} {'current':>11} {'candidate':>11}")
            for p in sorted(sl, key=lambda q: (-q.bucket_bits, q.point_id)):
                ec, ea = errors_for(p)
                tre = p.transform_relative_error
                print(f"{p.bucket_bits:>5} {p.construction:>17} "
                      f"{classify_output(reference(p)):>18} "
                      f"{('-' if tre is None else f'{tre:.3g}'):>14} "
                      f"{('-' if ec is None else f'{ec:.3g}'):>11} "
                      f"{('-' if ea is None else f'{ea:.3g}'):>11}")

        pub = envelope(surf, normal_only=False)
        alg = envelope(surf, normal_only=True)
        if pub:
            print(f"\n{'-' * 78}")
            print(f"{surface.upper()} ENVELOPES  (stress points only)")
            print(f"{'-' * 78}")
            print("   PUBLIC     all rows: what the Double-returning UDF "
                  "actually delivers.")
            print("   ALGORITHMIC normal_output rows only: where a branch "
                  "change can\n               recover information rather than "
                  "being limited by the\n               final binary64 "
                  "representation.")
            print_envelope("PUBLIC ENVELOPE", pub)
            pc, pb = crossover_of(pub)
            print(f"\n   public crossover        "
                  f"{f'{pc} bits and below' if pc else 'none'}")
            print(f"   public binding shape    {pb or '-'}")
            if alg:
                print_envelope("ALGORITHMIC ENVELOPE", alg)
                ac, ab = crossover_of(alg)
                print(f"\n   algorithmic crossover      "
                      f"{f'{ac} bits and below' if ac else 'none'}")
                print(f"   algorithmic binding shape  {ab or '-'}")
                if pc != ac:
                    print("\n   The two disagree. The public envelope is "
                          "confounded by rows whose\n   true output is not "
                          "normally representable; those cannot be recovered\n"
                          "   by any branch and must not choose the dispatch "
                          "boundary.")
            decomposition(surf)

    under = [p for p in points if p.construction == "underflow"]
    if under:
        print(f"\n{'#' * 78}\n# HARD UNDERFLOW  ({len(under)} observations)"
              f"\n{'#' * 78}")
        print("A positive mathematical ratio stored as zero. A requirement "
              "separate from\nthe bucket crossover: the log path is mandatory "
              "for information recovery,\nwhatever the crossover turns out to "
              "be.\n")
        print(f"{'surface':>11} {'shape':>9} {'point':>21} "
              f"{'current':>12} {'candidate':>14} {'out class':>19}")
        for p in sorted(under, key=lambda q: (q.surface, q.shape_id, q.point_id)):
            cur = (p.current_status if p.current_status != "OK"
                   else f"{p.current:.4g}")
            can = (p.candidate_status if p.candidate_status != "OK"
                   else f"{p.candidate:.4g}")
            print(f"{p.surface:>11} {p.shape_id:>9} {p.point_id:>21} "
                  f"{cur:>12} {can:>14} "
                  f"{classify_output(reference(p)):>19}")


def main() -> None:
    ap = argparse.ArgumentParser(description="Phase A analyzer for #13")
    ap.add_argument("csv", nargs="+", type=Path)
    args = ap.parse_args()

    oracle = verify_reference_invariants()
    if oracle:
        print("REFERENCE INVARIANT FAILED -- the oracle itself is wrong, so no "
              "observation\ncan be interpreted:\n")
        for o in oracle:
            print(f"   {o}")
        raise SystemExit(2)

    for path in args.csv:
        points = load(path)
        problems = verify_constructions(points) + twin_check(points)
        if problems:
            print(f"CONSTRUCTION CHECK FAILED for {path.name} -- no crossover "
                  f"is reported,\nbecause a point that is not where it claims "
                  f"to be would corrupt every\nbucket it appears in:\n")
            for p in problems:
                print(f"   {p}")
            raise SystemExit(1)
        print(f"{path.name}: construction check passed, {len(points)} "
              f"observations")
        report(points)

    print("\nPhase A is measurement only. No cutoff is proposed here and no "
          "source is changed.")


if __name__ == "__main__":
    main()
