#!/usr/bin/env python3
"""
cross_validate_oracle.py
========================

Cross-validates the mpmath references this study depends on against MPFR,
via Rmpfr, and writes a committable record.

WHY
---
The positive-ratio study (#13) measures VBA against mpmath. A single-oracle
study is only as good as that oracle. mpmath's self-consistency under
increasing precision proves convergence, not correctness -- a systematically
wrong algorithm converges to a wrong answer just as smoothly as a right one
converges to a right one.

MPFR is an independent C implementation sharing no lineage with mpmath's pure
Python. Agreement between them is meaningful in a way that mpmath agreeing with
itself is not.

R IS AN OPTIONAL DEPENDENCY
---------------------------
The hosted gate and the Windows workstation do not have R, and this must not
become a hard requirement. The script therefore emits a committed record that
downstream readers cite, rather than recomputing at gate time. When R is
absent it says so and exits 0, leaving the existing record in place.

Usage:
    python cross_validate_oracle.py                  # regenerate the record
    python cross_validate_oracle.py --check          # verify, do not rewrite
"""

from __future__ import annotations

import argparse
import csv
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from mpmath import mp, mpf, gammainc, exp as mp_exp, log as mp_log, loggamma

HERE = Path(__file__).parent
RECORD = HERE / "oracle_cross_validation.md"
R_SCRIPT = HERE / "cross_validate_oracle.R"

# Agreement required between the two oracles. Far tighter than anything the
# study measures (smallest observed error is about 1E-16), so a failure here
# means a real disagreement rather than a precision artefact.
AGREEMENT_TOLERANCE = mpf(10) ** -28

mp.dps = 110


def study_points() -> list[tuple[str, str, str]]:
    """The points the study actually relies on.

    Every shape slice used by either arm, crossed with buckets spanning the
    ladder, plus the hard-underflow quotients. Chosen to cover the regimes the
    conclusions rest on rather than to be numerous.
    """
    shapes = ["0.0001", "0.0005", "0.001", "0.002", "0.01",
              "0.1", "0.25", "0.5", "0.75", "1"]
    exps = [-1023, -1030, -1038, -1043, -1051, -1067, -1074]
    pts = []
    for a in shapes:
        for e in exps:
            z = mp.nstr(mpf(2) ** e, 25)
            for surface in ("cumulative", "survival", "density"):
                pts.append((a, z, surface))
    # hard-underflow quotients, where the current implementation fails outright
    for a in ("0.5", "1"):
        for z in ("1e-300", "1e-200"):
            for surface in ("cumulative", "density"):
                pts.append((a, z, surface))
    return pts


def mpmath_value(a_dec: str, z_dec: str, surface: str) -> mpf:
    a, z = mpf(a_dec), mpf(z_dec)
    if surface == "cumulative":
        return gammainc(a, 0, z, regularized=True)
    if surface == "survival":
        return gammainc(a, z, mp.inf, regularized=True)
    return mp_exp((a - 1) * mp_log(z) - z - loggamma(a))


def run_r(points: list[tuple[str, str, str]]) -> list[dict] | None:
    if shutil.which("Rscript") is None:
        return None
    with tempfile.TemporaryDirectory() as tmp:
        pin, pout = Path(tmp) / "points.csv", Path(tmp) / "r_oracle.csv"
        with pin.open("w", newline="") as fh:
            w = csv.writer(fh)
            w.writerow(["a_dec", "z_dec", "surface"])
            w.writerows(points)
        proc = subprocess.run(["Rscript", str(R_SCRIPT), str(pin), str(pout)],
                              capture_output=True, text=True)
        if proc.returncode != 0:
            print("Rscript failed:\n" + proc.stderr, file=sys.stderr)
            return None
        print(proc.stdout.strip())
        rows = list(csv.DictReader(pout.open(newline="")))
        # last field of the R banner: "... Rmpfr 0.9.5, MPFR 4.2.1"
        banner = proc.stdout.strip().split(":")[-1].strip()
        for r in rows:
            r["_banner"] = banner
        return rows


def compare(points, r_rows) -> tuple[list, mpf, list]:
    by_key = {(r["a_dec"], r["z_dec"], r["surface"]): r for r in r_rows}
    results, worst, problems = [], mpf(0), []
    for a, z, surface in points:
        r = by_key.get((a, z, surface))
        if r is None:
            problems.append(f"MPFR produced no value for {a}/{z}/{surface}")
            continue
        if r["stable"].strip().lower() != "true":
            problems.append(f"MPFR value unstable under added precision at "
                            f"{a}/{z}/{surface} -- cancellation not resolved")
            continue
        m = mpmath_value(a, z, surface)
        v = mpf(r["value"])
        if v == 0 and m == 0:
            rel = mpf(0)
        elif v == 0 or m == 0:
            problems.append(f"one oracle returned zero and the other did not "
                            f"at {a}/{z}/{surface}")
            continue
        else:
            rel = abs(m - v) / abs(v)
        if rel > AGREEMENT_TOLERANCE:
            problems.append(f"disagreement {mp.nstr(rel, 6)} at "
                            f"{a}/{z}/{surface}")
        worst = max(worst, rel)
        results.append((a, z, surface, m, v, rel, r["bits"]))
    return results, worst, problems


def write_record(results, worst, r_version: str) -> None:
    digits = int(-mp.log10(worst)) if worst > 0 else ">40"
    lines = [
        "# Oracle cross-validation — mpmath against MPFR",
        "",
        "Generated by `cross_validate_oracle.py`. **Do not hand-edit.**",
        "",
        "The positive-ratio study (#13) measures VBA against mpmath. mpmath's",
        "self-consistency under increasing precision proves convergence, not",
        "correctness, so its references are cross-checked against MPFR — an",
        "independent C implementation with no shared lineage.",
        "",
        f"- generated: `{datetime.now(timezone.utc).isoformat(timespec='seconds')}`",
        f"- mpmath at {mp.dps} decimal digits",
        f"- {r_version}",
        f"- points compared: **{len(results)}**",
        f"- worst relative disagreement: **{mp.nstr(worst, 6)}**",
        f"- agreeing significant digits: **~{digits}**",
        f"- tolerance: {mp.nstr(AGREEMENT_TOLERANCE, 3)}",
        "",
        "MPFR's `igamma` is the upper incomplete gamma, so the lower tail is",
        "`1 - igamma(a,z)/gamma(a)`. When P is tiny that subtraction cancels:",
        "at 200 bits a true P of 1E-157 returns exactly zero. Working precision",
        "is therefore chosen from the expected magnitude of P and every value is",
        "recomputed at double that precision; an unstable value is reported, not",
        "trusted. This is the same complement cancellation #13 addresses with",
        "`-Expm1(LogP)` in place of `1 - Exp(LogP)`.",
        "",
        "## Worst 12 disagreements",
        "",
        "| shape | z | surface | mpmath | MPFR | rel diff | bits |",
        "| --- | --- | --- | --- | --- | --- | --- |",
    ]
    for a, z, s, m, v, rel, bits in sorted(results, key=lambda r: -r[5])[:12]:
        # keep the exponent: a truncated mantissa makes 2^-1074 and 2^-1023
        # look identical, which would make the record unreadable
        lines.append(f"| {a} | {mp.nstr(mpf(z), 6)} | {s} | {mp.nstr(m, 12)} | "
                     f"{mp.nstr(v, 12)} | {mp.nstr(rel, 4)} | {bits} |")
    lines += ["", "All points agree within tolerance." if worst <= AGREEMENT_TOLERANCE
              else "", ""]
    RECORD.write_text("\n".join(lines))


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[3])
    ap.add_argument("--check", action="store_true",
                    help="verify without rewriting the record")
    args = ap.parse_args()

    points = study_points()
    r_rows = run_r(points)
    if r_rows is None:
        print("Rscript or Rmpfr unavailable -- the committed record stands.")
        print(f"Install with: apt-get install r-base-core r-cran-rmpfr")
        return 0

    results, worst, problems = compare(points, r_rows)
    if problems:
        print(f"ORACLE CROSS-VALIDATION FAILED ({len(problems)} problem(s)):\n")
        for p in problems[:20]:
            print(f"   {p}")
        return 1

    digits = int(-mp.log10(worst)) if worst > 0 else 40
    print(f"{len(results)} points compared, worst disagreement "
          f"{mp.nstr(worst, 6)} (~{digits} agreeing digits)")
    if not args.check:
        banner = r_rows[0].get("_banner", "Rmpfr / MPFR")
        write_record(results, worst, banner)
        print(f"wrote {RECORD.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
