#!/usr/bin/env python3
"""
Claim test for Amendment 1 of the positive-ratio study (#13).

AMENDMENT_1.md registers, before the second holdout is exported, what the
cumulative needs on the integrated incomplete-gamma kernel:

  C1  rounded transform, 40 retained bits or fewer: the log-domain candidate's
      worst error is below the current path's, at every tested bucket.
  C2  exact transform, every bucket: the current path's worst error is not
      above the candidate's. This is why an exact transform is not dispatched.
  C3  rounded transform, more than 40 bits: the current path's worst error is
      at most SEAM_BOUND. This is what stays unrecovered above the cutoff.
  C4  Gamma only, exact transform with a scale other than one, 40 bits or
      fewer: the candidate's worst error is at most PROXY_BOUND. Production
      tests ScaleParam <> 1, so these points are recovered although their
      transform is exact; C4 bounds what that costs.

The file is tested AGAINST these claims. No cutoff is proposed from it: a
rejection sends the amendment back to design.

The exporters write only the cumulative. A point's class is decided here from
the echoed operands, as exact rationals, never from its label: the label is
checked against it.

Usage: python3 analyze_amendment1.py gamma_holdout2.csv chisquare_holdout2.csv
Exit 0 only when every claim holds on every file.
"""
from __future__ import annotations

import argparse
import sys
from fractions import Fraction
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import analyze_positive_ratio_subnormal as A  # noqa: E402

CLAIM_BITS = 40
SEAM_BOUND = 5e-13
PROXY_BOUND = 3e-13
BUCKETS = (51, 47, 43, 42, 39, 35, 31, 27, 23, 19, 15, 11, 7)
OFFSETS = {
    "gamma": {Fraction(0), Fraction(2, 5), Fraction(3, 7), Fraction(1, 6), Fraction(1, 3)},
    "chisquare": {Fraction(0), Fraction(1, 2)},
}


def is_exact(p) -> bool:
    return p.stored is not None and Fraction(p.stored) == p.exact_standardx


def group(p) -> str:
    """exact_unit, exact_scaled or rounded, from the operands."""
    if not is_exact(p):
        return "rounded"
    if p.family == "gamma" and p.scale == 1.0:
        return "exact_unit"
    return "exact_scaled"


def verify(points) -> list[str]:
    problems = []
    for p in points:
        where = f"{p.point_id} {p.shape_id}"
        if p.surface != "cumulative":
            problems.append(f"{where}: surface {p.surface!r}; only the cumulative is registered")
        if p.bucket_bits not in BUCKETS:
            problems.append(f"{where}: bucket {p.bucket_bits} is not a registered bucket")
            continue
        if p.stored is None or p.stored == 0.0:
            problems.append(f"{where}: no positive stored argument")
            continue
        landmark = Fraction(2) ** (p.bucket_bits - 1) * A.MIN_SUBNORMAL
        offset = (p.exact_standardx - landmark) / A.MIN_SUBNORMAL
        if offset not in OFFSETS[p.family]:
            problems.append(f"{where}: exact argument is {offset} ULP above its bucket "
                            "boundary, not a registered offset")
        if group(p) != p.construction:
            problems.append(f"{where}: labelled {p.construction}, but the operands make it "
                            f"{group(p)}")
        if A.bits_available(p.stored) not in (p.bucket_bits, p.bucket_bits + 1):
            problems.append(f"{where}: stored value carries {A.bits_available(p.stored)} "
                            f"bits, bucket declares {p.bucket_bits}")
    return problems


def envelope(points) -> tuple[dict, list[str]]:
    """Worst current and candidate error per (group, bucket), over rows whose
    true result is a normal Double. A row without a comparable value is
    returned apart: it must not vanish from a claim's denominator."""
    worst: dict = {}
    unavailable = []
    for p in points:
        ref = A.reference(p)
        if A.classify_output(ref) != "normal_output":
            continue
        cur, cand = A.relative_error(p.current, ref), A.relative_error(p.candidate, ref)
        if cur is None or cand is None or p.current_status != "OK" or p.candidate_status != "OK":
            unavailable.append(f"{p.point_id} {p.shape_id}")
            continue
        key = (group(p), p.bucket_bits)
        w = worst.setdefault(key, [0.0, 0.0, 0])
        w[0], w[1], w[2] = max(w[0], cur), max(w[1], cand), w[2] + 1
    return worst, unavailable


def test_claims(points, out=print) -> bool:
    family = points[0].family
    worst, unavailable = envelope(points)
    ok = True

    def verdict(name, text, rows, failed):
        nonlocal ok
        if not rows:
            ok = False
            out(f"-- {name}: FAIL - no testable bucket ({text})")
            return
        if failed:
            ok = False
            out(f"-- {name}: FAIL - {text}; {len(failed)} of {len(rows)} buckets violate it")
        else:
            out(f"-- {name}: PASS - {text}; {len(rows)} buckets")
        out(f"   {'bits':>4} {'group':>12} {'rows':>4} {'worst current':>14} {'worst candidate':>16}")
        for (g, bits), (wc, wa, n) in rows:
            mark = "  <-- violates" if ((g, bits), (wc, wa, n)) in failed else ""
            out(f"   {bits:>4} {g:>12} {n:>4} {wc:>14.3g} {wa:>16.3g}{mark}")

    items = sorted(worst.items(), key=lambda kv: (-kv[0][1], kv[0][0]))
    out(f"\nAMENDMENT 1 - {family.upper()} cumulative, second holdout")
    if unavailable:
        ok = False
        out(f"-- PENDING: {len(unavailable)} row(s) have no comparable value: "
            + ", ".join(unavailable[:5]))

    rows = [kv for kv in items if kv[0][0] == "rounded" and kv[0][1] <= CLAIM_BITS]
    verdict("C1", f"rounded, <= {CLAIM_BITS} bits: candidate below current",
            rows, [kv for kv in rows if not kv[1][1] < kv[1][0]])

    rows = [kv for kv in items if kv[0][0] != "rounded"]
    verdict("C2", "exact, every bucket: current not above candidate",
            rows, [kv for kv in rows if kv[1][0] > kv[1][1]])

    rows = [kv for kv in items if kv[0][0] == "rounded" and kv[0][1] > CLAIM_BITS]
    verdict("C3", f"rounded, > {CLAIM_BITS} bits: current at most {SEAM_BOUND:g}",
            rows, [kv for kv in rows if kv[1][0] > SEAM_BOUND])

    if family == "gamma":
        rows = [kv for kv in items if kv[0][0] == "exact_scaled" and kv[0][1] <= CLAIM_BITS]
        verdict("C4", f"exact with scale <> 1, <= {CLAIM_BITS} bits: candidate at most "
                f"{PROXY_BOUND:g}", rows, [kv for kv in rows if kv[1][1] > PROXY_BOUND])

    out("CLAIMS HOLD" if ok else "AT LEAST ONE CLAIM REJECTED OR PENDING")
    return ok


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    ap.add_argument("csv", nargs="+", type=Path)
    a = ap.parse_args(argv)
    oracle = A.verify_reference_invariants()
    if oracle:
        print("REFERENCE INVARIANT FAILED:\n   " + "\n   ".join(oracle))
        return 2
    holds = True
    for path in a.csv:
        points = A.load(path)
        problems = verify(points)
        if problems:
            print(f"CONSTRUCTION CHECK FAILED for {path.name}:")
            for p in problems[:20]:
                print("   " + p)
            return 1
        print(f"{path.name}: construction check passed, {len(points)} observations")
        holds = test_claims(points) and holds
    return 0 if holds else 1


if __name__ == "__main__":
    sys.exit(main())
