"""
Audit the study references that the five #12/#15 contracts will be frozen from.

The main grid's 658 stale references were not the product of sloppy formulas.
They were computed at the intended decimal rather than at the Double the
exporter delivers, and that distinction was invisible until it was measured.
These two studies were written before that lesson, so their references get the
same treatment before anything is promoted:

  * every reference recomputed at the exact Val(token) Double, obtained as an
    exact rational via float.as_integer_ratio(), so no decimal conversion
    enters the oracle;
  * scored against an INDEPENDENT high-precision oracle - mpmath primitives
    here, never the study's own generator - so the audit cannot confirm the
    thing it is checking;
  * certified at 160 and 260 digits, the disagreement between them recorded as
    the oracle's own uncertainty.

The regimes study carries subnormal arguments and seam neighbours one ulp
apart, which is exactly where the decimal-vs-Double distinction bites hardest.

Classification matches audit_references.py:
  REFERENCE_CORRECT   committed reference agrees with the oracle
  REFERENCE_STALE     it does not, by more than the serialisation floor
  ORACLE_AMBIGUOUS    the oracle could not be certified

Report only. Writes nothing but its own CSV.
"""
import argparse, csv, math, os, sys
from collections import Counter

import mpmath as mp

DPS_LOW, DPS_HIGH = 160, 260
SER_FLOOR = mp.mpf(10) ** -24          # study references carry 25+ digits
MARGIN = mp.mpf(4)


def q(text):
    """Exact rational value of the Double the token denotes."""
    n, d = float(text).as_integer_ratio()
    return mp.mpf(n) / d


def loggamma1p(x):
    """Log(Gamma(1 + x)), independent of the VBA kernel and of the study.

    mp.loggamma(1 + x) collapses for small x: at 160 digits 1 + 1E-300 rounds
    to 1 and the answer becomes 0. The precision must rise with the magnitude
    of x - the same adaptive rule the kernel itself exists to embody.
    """
    if x == 0:
        return mp.mpf(0)
    need = mp.mp.dps + max(0, int(-mp.log10(abs(x))) + 20)
    with mp.workdps(need):
        return +mp.loggamma(1 + mp.mpf(x))


def loggamma(z):
    with mp.workdps(mp.mp.dps):
        return +mp.loggamma(mp.mpf(z))



def _selfcheck_oracle():
    """The study generators and this audit both call mp.loggamma, so agreement
    is not independent confirmation unless the primitive itself is checked
    against a different formulation. Runs once, at import.

      LogGamma      vs mp.log(mp.gamma(z)), valid where gamma does not overflow
      LogGamma1p    vs the Maclaurin series -EulerGamma*x + sum zeta(k)(-x)^k/k
    """
    old = mp.mp.dps
    try:
        mp.mp.dps = 200
        for z in ("0.1", "0.5", "1.75", "2.5", "10", "0.25", "0.9"):
            Z = mp.mpf(float(z))
            a, b = mp.loggamma(Z), mp.log(mp.gamma(Z))
            if abs(a - b) > abs(a) * mp.mpf(10) ** -150:
                raise AssertionError(f"mp.loggamma disagrees with log(gamma) at z={z}")
        for x in ("1e-13", "1e-16", "0.01", "0.25", "1e-100"):
            X = mp.mpf(float(x))
            need = mp.mp.dps + max(0, int(-mp.log10(abs(X))) + 20)
            with mp.workdps(need):
                a = +mp.loggamma(1 + X)
                s = mp.mpf(0)
                for k in range(40, 1, -1):
                    s = (s + ((-1) ** k) * mp.zeta(k) / k) * X
                b = +((s - mp.euler) * X)
            if abs(a - b) > abs(a) * mp.mpf(10) ** -20:
                raise AssertionError(f"loggamma(1+x) disagrees with the series at x={x}")
    finally:
        mp.mp.dps = old


_selfcheck_oracle()


ORACLE = {
    "LogGamma": ("mp.loggamma", loggamma),
    "LogGamma1p": ("adaptive mp.loggamma(1+x)", loggamma1p),
    "LogGamma1pOverX": ("adaptive mp.loggamma(1+x) / x",
                        lambda x: loggamma1p(x) / mp.mpf(x)),
}


def audit(path, out_rows):
    rows = list(csv.DictReader(open(path)))
    name = os.path.basename(path)
    for r in rows:
        qty = r["quantity"]
        if qty not in ORACLE:
            continue                     # EchoZ/EchoX are parser probes; LogGammaNaive is the defect
        method, f = ORACLE[qty]
        arg = r["arg1"]
        vals = []
        try:
            for dps in (DPS_LOW, DPS_HIGH):
                mp.mp.dps = dps
                vals.append(f(q(arg)))
            mp.mp.dps = DPS_HIGH
            lo, hi = vals
            if not (mp.isfinite(lo) and mp.isfinite(hi)):
                raise ValueError("non-finite oracle")
            delta = abs(hi - lo)
            o = hi
            c = mp.mpf(r["reference"])
            scale = abs(o) if o != 0 else mp.mpf(1)
            err = abs(c - o)
            unc = max(delta, abs(o) * mp.mpf(10) ** -(DPS_HIGH - 20))
            floor_ = max(SER_FLOOR * scale, unc * MARGIN)
            cls = "REFERENCE_CORRECT" if err <= floor_ else "REFERENCE_STALE"
            out_rows.append({
                "study": name, "quantity": qty, "regime": r.get("regime", ""),
                "arg1": arg, "committed_reference": r["reference"],
                "independent_oracle": mp.nstr(o, 30),
                "abs_error": mp.nstr(err, 6),
                "rel_error": mp.nstr(err / scale, 6),
                "oracle_convergence_delta": mp.nstr(delta, 6),
                "classification": cls, "oracle_method": method,
            })
        except Exception as e:
            out_rows.append({
                "study": name, "quantity": qty, "regime": r.get("regime", ""),
                "arg1": arg, "committed_reference": r["reference"],
                "independent_oracle": "", "abs_error": "", "rel_error": "",
                "oracle_convergence_delta": "",
                "classification": "ORACLE_AMBIGUOUS",
                "oracle_method": f"{method} FAILED: {type(e).__name__}",
            })
        finally:
            mp.mp.dps = 50


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="study_reference_audit.csv")
    ap.add_argument("--study", action="append", default=[
        "loggamma_regimes_study/loggamma_regimes_grid.csv",
        "loggamma_regimes_study/loggamma_regimes_holdout.csv",
        "loggamma1p_study/loggamma1p_grid.csv",
        # The LogGamma1p holdout is evidence for a frozen threshold, so its
        # references need the same independent certification as the fitting
        # grid's. It found the worst case of the whole contract, at X=1E-307,
        # which makes certifying it more important than the fitting set, not
        # less.
        "loggamma1p_study/loggamma1p_holdout.csv",
    ])
    a = ap.parse_args()
    out = []
    for p in a.study:
        if os.path.exists(p):
            audit(p, out)
        else:
            print(f"  missing: {p}")
    with open(a.out, "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(out[0].keys()), lineterminator="\n")
        w.writeheader(); w.writerows(out)
    print(f"audited {len(out)} study references at 160 and 260 digits\n")
    for s in sorted({r["study"] for r in out}):
        sub = [r for r in out if r["study"] == s]
        c = Counter(r["classification"] for r in sub)
        print(f"  {s:38s} {len(sub):4d}  {dict(c)}")
    print()
    for k, v in Counter(r["classification"] for r in out).most_common():
        print(f"  {k:22s} {v:5d}")
    stale = [r for r in out if r["classification"] == "REFERENCE_STALE"]
    if stale:
        print(f"\n  {'study':34s} {'quantity':18s} {'arg1':>22s} {'rel err':>10s}")
        for r in sorted(stale, key=lambda r: -float(r["rel_error"] or 0))[:12]:
            print(f"  {r['study']:34s} {r['quantity']:18s} {r['arg1']:>22s} {r['rel_error']:>10s}")
    print(f"\nwrote {a.out}")


if __name__ == "__main__":
    main()
