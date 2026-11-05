"""
Build benchmark/cdf_large_shape/cdf_large_shape_grid.csv.

WHY THIS STUDY EXISTS
    CR-P1-02 removed the cancelling prefactors from the regularized incomplete
    gamma and beta kernels. Those kernels carry every cumulative and survival
    probability outside the normal family, yet the frozen contracts reached only
    shape 3 for Gamma and 1E5 for Beta, so the repaired regime was unmeasured.
    This grid measures it directly and supplies the evidence for contracts.

REFERENCES
    mpmath, at a working precision that grows with the shape (#159). The log
    prefactor -X + A*Log(X) - LogGamma(A) subtracts terms of size A*Log(A), so
    the fixed 30 digits used before kept only 17 to 24 significant digits at
    shapes of 1E6 and above. Each reference is now computed at 40 + log10(A)
    digits and again 15 digits higher, and kept only when the two agree to 32
    significant digits.

    Gamma: the route that converges, the ascending series below A + 1 and the
    Lentz continued fraction above it. The continued fraction gives the
    survival directly and the series gives the cumulative directly; the other
    side is formed as a complement only where it is at least about one half.
    Beta (balanced shapes): the continued fraction on the side where it
    converges, which gives the upper tail directly above the mean, and the
    symmetry I_x(a,a) = 1 - I_{1-x}(a,a) as a cross-check. A reference that
    fails a check is dropped rather than trusted.

    Each X is a Double and its 17-digit token parses back to it, so every
    reference is evaluated at the exact argument the exporter passes.

    The committed grid holds Excel observations, which this script cannot
    reproduce. Re-reference it, changing reference cells only, with

        python generate_cdf_large_shape.py --merge-references          # report
        python generate_cdf_large_shape.py --merge-references --write  # patch

REACHABILITY
    The lower-tail series needs O(sqrt(A)) terms, so the VBA kernel reaches its
    100000-iteration cap somewhere above A = 1E8 and returns #NUM! rather than a
    partial sum. Points beyond that are retained deliberately: an ERROR cell is
    the measurement, and it documents where the validated domain stops.
"""
import argparse
import csv
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
# Design precision: the test points X are chosen at 30 digits, as when the grid
# was built, so the argument tokens do not move. Every reference is computed
# under its own, higher working precision (_at below).
mp.mp.dps = 30
BASE_DPS = 40               # plus log10(shape)
CHECK_EXTRA = 15            # the check runs this many digits higher
AGREE = mp.mpf(10) ** -32   # relative agreement required between the two

SHAPES = [1e2, 1e4, 1e6, 1e8, 1e10, 1e12]
Z = [-3, -1, 0, 1, 3]


def _eps():
    """Convergence tolerance tied to the working precision."""
    return mp.mpf(10) ** -(mp.mp.dps - 4)


def _tiny():
    return mp.mpf(10) ** -(3 * mp.mp.dps)


def _gamma_log_prefactor(A, X):
    return -X + A * mp.log(X) - mp.loggamma(A)


def gamma_P_series(A, X):
    """Regularized lower incomplete gamma by the ascending series."""
    A = mp.mpf(A); X = mp.mpf(X)
    eps = _eps()
    Ap = A; S = 1 / A; Del = S; n = 0
    while n < 3000000:
        Ap += 1; Del *= X / Ap; S += Del; n += 1
        if abs(Del) <= abs(S) * eps:
            break
    return S * mp.e ** _gamma_log_prefactor(A, X)


def gamma_Q_cf(A, X):
    """Regularized upper incomplete gamma by the Lentz continued fraction."""
    A = mp.mpf(A); X = mp.mpf(X)
    tiny, eps = _tiny(), _eps()
    b = X + 1 - A; c = 1 / tiny; d = 1 / b; h = d
    for i in range(1, 200000):
        an = -mp.mpf(i) * (mp.mpf(i) - A)
        b += 2
        d = an * d + b
        if abs(d) < tiny: d = tiny
        c = b + an / c
        if abs(c) < tiny: c = tiny
        d = 1 / d; de = d * c; h *= de
        if abs(de - 1) <= eps:
            break
    return h * mp.e ** _gamma_log_prefactor(A, X)


def gamma_PQ(A, X):
    """(P, Q) at the active precision, each from the side that converges."""
    if X < A + 1:
        P = gamma_P_series(A, X)
        return P, 1 - P
    Q = gamma_Q_cf(A, X)
    return 1 - Q, Q


def beta_cf(A, B, X):
    """Lentz continued fraction for the incomplete beta, direct orientation."""
    A = mp.mpf(A); B = mp.mpf(B); X = mp.mpf(X)
    tiny, eps = _tiny(), _eps()
    qab = A + B; qap = A + 1; qam = A - 1
    c = mp.mpf(1); d = 1 - qab * X / qap
    if abs(d) < tiny: d = tiny
    d = 1 / d; h = d
    for m in range(1, 200000):
        m2 = 2 * m
        aa = m * (B - m) * X / ((qam + m2) * (A + m2))
        d = 1 + aa * d
        if abs(d) < tiny: d = tiny
        c = 1 + aa / c
        if abs(c) < tiny: c = tiny
        d = 1 / d; h *= d * c
        aa = -(A + m) * (qab + m) * X / ((A + m2) * (qap + m2))
        d = 1 + aa * d
        if abs(d) < tiny: d = tiny
        c = 1 + aa / c
        if abs(c) < tiny: c = tiny
        d = 1 / d; de = d * c; h *= de
        if abs(de - 1) <= eps:
            break
    return h


def _beta_front(A, B, X):
    return mp.e ** (A * mp.log(X) + B * mp.log1p(-X)
                    - (mp.loggamma(A) + mp.loggamma(B) - mp.loggamma(A + B)))


def beta_I(A, B, X):
    """Regularized incomplete beta, using whichever orientation converges."""
    A = mp.mpf(A); B = mp.mpf(B); X = mp.mpf(X)
    if X <= 0: return mp.mpf(0)
    if X >= 1: return mp.mpf(1)
    front = _beta_front(A, B, X)
    if X < (A + 1) / (A + B + 2):
        return front * beta_cf(A, B, X) / A
    return 1 - front * beta_cf(B, A, 1 - X) / B


def beta_S(A, B, X):
    """Upper tail 1 - I_X(A, B), direct above the switch point near the mean."""
    A = mp.mpf(A); B = mp.mpf(B); X = mp.mpf(X)
    if X <= 0: return mp.mpf(1)
    if X >= 1: return mp.mpf(0)
    front = _beta_front(A, B, X)
    if X < (A + 1) / (A + B + 2):
        return 1 - front * beta_cf(A, B, X) / A
    return front * beta_cf(B, A, 1 - X) / B


def _at(dps, f, *args):
    with mp.workdps(dps):
        v = f(*args)
        return tuple(+x for x in v) if isinstance(v, tuple) else +v


def _agree(a, b):
    return abs(a - b) <= AGREE * abs(b)


def build():
    rows = []
    dropped = 0
    worst = mp.mpf(0)

    def add(fn, a1, a2, a3, ref, regime):
        rows.append({
            "function": fn, "vba_kernel": f"K_STATS_{fn}", "claim": "characterization",
            "metric": "rel", "arg1": mp.nstr(mp.mpf(a1), 17), "arg2": mp.nstr(mp.mpf(a2), 17),
            "arg3": ("" if a3 is None else mp.nstr(mp.mpf(a3), 17)),
            "reference": mp.nstr(ref, 34), "observed_vba": "", "regime": regime,
            "evidence_set": "cdf_large_shape"})

    for a in SHAPES:
        A = mp.mpf(a)
        sd = mp.sqrt(A)
        dps = BASE_DPS + int(mp.ceil(mp.log10(A)))
        for z in Z:
            X = mp.mpf(float(A + z * sd))
            if X <= 0:
                continue
            # Self-check by recomputation at higher precision: a value that has
            # genuinely converged does not move; one that hit an iteration cap
            # or ran out of digits does.
            P, Q = _at(dps, gamma_PQ, A, X)
            P2, Q2 = _at(dps + CHECK_EXTRA, gamma_PQ, A, X)
            if not (0 < P2 < 1) or not (_agree(P, P2) and _agree(Q, Q2)):
                dropped += 1
                continue
            add("Gamma_Cumulative", X, A, 1.0, P2, f"z{z:+d}")
            add("Gamma_Survival", X, A, 1.0, Q2, f"z{z:+d}")
            worst = max(worst, abs(P - P2) / P2, abs(Q - Q2) / Q2)

        # Beta, balanced shapes, x near the mode 0.5
        bsd = 1 / (2 * mp.sqrt(2 * A))
        for z in Z:
            X = mp.mpf(float(mp.mpf("0.5") + z * bsd))
            if not (0 < X < 1):
                continue
            I, S = _at(dps, beta_I, A, A, X), _at(dps, beta_S, A, A, X)
            I2 = _at(dps + CHECK_EXTRA, beta_I, A, A, X)
            S2 = _at(dps + CHECK_EXTRA, beta_S, A, A, X)
            sym = _at(dps + CHECK_EXTRA, beta_S, A, A, 1 - X)   # = I_x(a,a)
            if not (0 < I2 < 1) or not (_agree(I, I2) and _agree(S, S2) and _agree(sym, I2)):
                dropped += 1
                continue
            add("Beta_Cumulative", X, A, A, I2, f"bal_z{z:+d}")
            add("Beta_Survival", X, A, A, S2, f"bal_z{z:+d}")
            worst = max(worst, abs(I - I2) / I2, abs(S - S2) / S2, abs(sym - I2) / I2)

    return rows, dropped, worst


FIELDS = ["function", "vba_kernel", "claim", "metric", "arg1", "arg2", "arg3",
          "reference", "observed_vba", "regime", "evidence_set"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--merge-references", action="store_true",
                    help="re-reference the committed grid, keeping every observation")
    ap.add_argument("--write", action="store_true",
                    help="with --merge-references: patch the committed grid")
    args = ap.parse_args()
    rows, dropped, worst = build()
    # Neither check is an independent route. Each reference is one route
    # (series or continued fraction, whichever converges) at two precisions;
    # Beta adds the symmetry, evaluated with the same continued fraction.
    print("self-check (two precisions per reference, plus the Beta symmetry) "
          f"worst relative disagreement: {mp.nstr(worst, 6)}")
    if worst > AGREE:
        raise SystemExit("ABORT: reference self-check failed; no grid written")
    grid = os.path.join(HERE, "cdf_large_shape_grid.csv")
    if args.merge_references:
        sys.path.insert(0, os.path.dirname(HERE))
        from _study_merge import merge_references, print_report, study_describe, study_key
        with open(grid, newline="", encoding="utf-8") as fh:
            committed = {study_key(r) for r in csv.DictReader(fh)}
        # Adding a row needs an Excel export, so the merge keeps the committed
        # design and reports any point the generator would now add.
        extra = [r for r in rows if study_key(r) not in committed]
        if extra:
            print(f"  {len(extra)} generated row(s) are outside the committed design "
                  "and are not added (adding one needs an Excel export):")
            for r in extra:
                print(f"     {r['function']} A={r['arg2']} regime={r['regime']}")
        rows = [r for r in rows if study_key(r) in committed]
        changes, held = merge_references(grid, rows, study_key, {"observed_vba"},
                                         write=args.write, describe=study_describe)
        print_report("cdf_large_shape_grid.csv", changes, held, args.write)
        return
    with open(grid, "w", newline="\n", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS, lineterminator="\n")
        w.writeheader()
        for r in rows:
            w.writerow(r)
    print(f"wrote cdf_large_shape_grid.csv: {len(rows)} rows, {dropped} dropped")
    from collections import Counter
    for fn, c in sorted(Counter(r["function"] for r in rows).items()):
        print(f"  {fn}: {c}")


if __name__ == "__main__":
    main()
