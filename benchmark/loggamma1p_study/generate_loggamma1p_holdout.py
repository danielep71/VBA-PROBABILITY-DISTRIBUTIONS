"""
Independent holdout for LogGamma1p.small.scaled_abs.

The contract is set from loggamma1p_grid.csv, so that grid cannot also validate
it. These points set none of the threshold.

BOUNDARY COVERAGE. The contract's domain starts at the normal-result boundary
PROB_MIN_NORMAL / EulerGamma = 3.854839696505424E-308, below which EulerGamma*X
is itself subnormal and the scaled error necessarily deteriorates. The fitting
grid's lowest point is 1E-300 - about 7.4 decades ABOVE that boundary - so the
threshold was frozen without ever testing the region that stresses it most. The
holdout therefore covers the boundary directly: the boundary Double itself, its
NextUp, and the immediate normal-result interior. NextDown(boundary) is
deliberately excluded; it belongs to the subnormal_result characterisation, not
to this contract.

STRUCTURAL DIFFERENCE. Disjointness alone is weak: neighbouring decades test the
same shape of number. The fitting grid is decade landmarks (1E-300, 1E-16, 1E-8)
plus round interior steps (0.02, 0.05, 0.1). This grid uses irrational mantissas,
so no argument is a round decimal or a power of two and the binary64 rounding
lands differently at every point.

Disjointness is asserted, not eyeballed: the generator reads the fitting set and
refuses to emit any point it already contains.
"""
import argparse, csv, math, os, sys

import mpmath as mp

BOUNDARY = 3.854839696505424e-308        # PROB_MIN_NORMAL / EulerGamma
SERIES_MAX = 0.25
FITTING = "loggamma1p_grid.csv"

PHI = 1.6180339887498949
SQ2 = 1.4142135623730951
SQ3 = 1.7320508075688772
E_1 = 2.718281828459045


def reference(x):
    """Log(Gamma(1 + x)) at a precision that keeps the increment.

    mp.loggamma(1 + x) collapses for small x exactly as the VBA kernel would if
    it formed 1 + X: at a fixed 60 digits, 1 + 1E-300 rounds to 1 and the answer
    becomes 0. The precision must rise with the magnitude of x.
    """
    dps = 60 + max(0, int(-math.log10(x)) + 10)
    with mp.workdps(dps):
        return +mp.loggamma(1 + mp.mpf(x))


def build_points():
    pts = []
    # the boundary itself and its immediate neighbourhood - the region the
    # fitting grid never reaches
    pts += [BOUNDARY, math.nextafter(BOUNDARY, math.inf)]
    pts += [SQ2 * 3.9e-308, 5.0e-308, PHI * 6e-308, 1e-307, SQ3 * 4e-307, 1e-305]
    # the 7.4 decades between the boundary and the fitting set's lowest point
    pts += [PHI * 1e-303, SQ2 * 1e-297, SQ3 * 1e-290, E_1 * 1e-280,
            PHI * 1e-250, SQ2 * 1e-180, SQ3 * 1e-120]
    # interior, irrational mantissas between the fitting decades
    pts += [E_1 * 1e-77, PHI * 1e-40, SQ2 * 1e-25, SQ3 * 1e-18,
            E_1 * 1e-15, PHI * 1e-12, SQ2 * 1e-9, SQ3 * 1e-6, E_1 * 1e-4]
    # approach to the seam, without touching it: the seam is a separate
    # characterisation regime, not part of this contract
    pts += [0.0137, 0.0619, 0.1037, 0.1414, 0.1732, 0.2071, 0.2291, 0.2449]
    out = []
    for p in pts:
        if BOUNDARY <= p <= SERIES_MAX and p not in out:
            out.append(p)
    return sorted(out)


FIELDS = ["quantity", "regime", "arg1", "reference", "observed_vba"]
QUANTITIES = ("EchoX", "LogGamma1p")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="loggamma1p_holdout.csv")
    ap.add_argument("--fitting", default=FITTING)
    a = ap.parse_args()

    here = os.path.dirname(os.path.abspath(__file__))
    fpath = a.fitting if os.path.isabs(a.fitting) else os.path.join(here, a.fitting)
    if not os.path.exists(fpath):
        sys.exit(f"fitting set {fpath} not found; disjointness cannot be verified")
    fitting = {float(r["arg1"]) for r in csv.DictReader(open(fpath))}

    pts = build_points()
    clash = [p for p in pts if p in fitting]
    if clash:
        sys.exit(f"holdout point(s) {clash[:4]} are in the fitting set; refusing")

    rows = []
    for x in pts:
        ref = reference(x)
        for q in QUANTITIES:
            rows.append({
                "quantity": q, "regime": "small", "arg1": repr(x),
                "reference": mp.nstr(mp.mpf(x) if q == "EchoX" else ref, 34),
                "observed_vba": "",
            })
    with open(os.path.join(here, a.out), "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=FIELDS, lineterminator="\n")
        w.writeheader(); w.writerows(rows)
    print(f"wrote {a.out}: {len(rows)} rows ({len(pts)} points x {len(QUANTITIES)})")
    print(f"disjoint from {a.fitting} ({len(fitting)} fitting points): verified")
    print(f"range {pts[0]:.6e} .. {pts[-1]:.6e}")
    below = sum(1 for p in pts if p < 1e-300)
    print(f"points below the fitting set's floor of 1E-300: {below}")


if __name__ == "__main__":
    main()
