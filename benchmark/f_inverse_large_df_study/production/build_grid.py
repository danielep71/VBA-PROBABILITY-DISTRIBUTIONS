"""
#35: the production holdout of the beta inverse (PRODUCTION_PREREGISTRATION.md).

One design, two files with the same rows: f_inverse_pair_baseline.csv, which
Excel fills from the source before the change, and f_inverse_pair_grid.csv,
which Excel fills from the candidate.

  holdout_f     the sixteen points reserved by FIX_COMPARISON_PREREGISTRATION.md
                and never evaluated: p in {0.1, 0.4, 0.6, 0.98} crossed with
                (d1, d2) in {(3E5, 7), (7, 3E5), (50, 1E8), (1E8, 50)};
  far_f         F at p in {1E-11, 1 - 1E-11} and five new pairs of degrees of
                freedom;
  seam_f        F at the orientation boundary of (8, 14) and (14, 8): the
                Double the model gives for the lower tail at 1/2, moved by 0,
                1 and 4 units in the last place either way;
  t_invariance  Student-t at df in {1, 2, 3, 5, 10, 30} and the fitting
                probabilities of the #22 study: 0.5 +- 2^-53, 2^-43, 2^-33,
                2^-23, and 0.625, 0.75, 0.875;
  beta          Beta at six pairs of shapes, one with both below one, and
                p in {1E-9, 0.2, 0.8, 1 - 1E-9}.

Columns (8): point_id, kind, func, p_bits, a_bits, b_bits, ref, value. The
first seven are the design; M_STATS_PROBDIST_FIP.bas writes value. The
arguments are the 16 hexadecimal digits of the Doubles. b_bits is NA for
Student-t. ref is the quantile to 40 digits, from 120-digit arithmetic.

Run:
    python3 build_grid.py --write   (only when the files are absent)
    python3 build_grid.py --check   (the committed files are the design)
"""
import argparse
import csv
import io
import itertools
import os
import struct
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
STUDY = os.path.join(os.path.dirname(HERE), "fix_comparison")
BENCH = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, STUDY)
sys.path.insert(0, BENCH)
import _ibeta as oracle  # noqa: E402

FILES = {"baseline": os.path.join(HERE, "f_inverse_pair_baseline.csv"),
         "candidate": os.path.join(HERE, "f_inverse_pair_grid.csv")}
FIELDS = ["point_id", "kind", "func", "p_bits", "a_bits", "b_bits", "ref", "value"]
KEY_COLUMNS = 7
DIGITS = 120


def bits(x):
    return f"{struct.unpack('<Q', struct.pack('<d', x))[0]:016X}"


def from_bits(text):
    return struct.unpack("<d", struct.pack("<Q", int(text, 16)))[0]


def step(x, n):
    return struct.unpack("<d", struct.pack("<Q", struct.unpack("<Q", struct.pack("<d", x))[0] + n))[0]


def reference(func, p, a, b):
    """The quantile as an mpf: F(p; d1, d2), t(p; df) or Beta(p; a, b)."""
    with mp.workdps(DIGITS):
        fp = mp.mpf(p)
        if func == "F":
            x = oracle.beta_invcdf(fp, mp.mpf(a) / 2, mp.mpf(b) / 2)
            return mp.mpf(b) / mp.mpf(a) * x / (1 - x)
        if func == "Beta":
            return oracle.beta_invcdf(fp, mp.mpf(a), mp.mpf(b))
        # Student-t: with y = t^2 / (df + t^2), I_y(1/2, df/2) = |2p - 1|
        if fp == mp.mpf(1) / 2:
            return mp.mpf(0)
        y = oracle.beta_invcdf(abs(2 * fp - 1), mp.mpf(1) / 2, mp.mpf(a) / 2)
        t = mp.sqrt(mp.mpf(a) * y / (1 - y))
        return t if fp > mp.mpf(1) / 2 else -t


def points():
    rows = [("holdout_f", "F", p, d1, d2) for p, (d1, d2) in itertools.product(
        (0.1, 0.4, 0.6, 0.98), ((3e5, 7.0), (7.0, 3e5), (50.0, 1e8), (1e8, 50.0)))]
    rows += [("far_f", "F", p, d1, d2) for p in (1e-11, 1 - 1e-11) for d1, d2 in
             ((9.0, 4e5), (4e5, 9.0), (20.0, 3e9), (3e9, 20.0), (150.0, 260.0))]
    import tail_aware_model as p4
    for d1, d2 in ((8.0, 14.0), (14.0, 8.0)):
        h = p4.tails(0.5, d1 / 2.0, d2 / 2.0)[0]
        rows += [("seam_f", "F", step(h, n), d1, d2) for n in (-4, -1, 0, 1, 4)]
    t_probabilities = sorted([0.5 + s * 2.0 ** -k for k in (53, 43, 33, 23) for s in (-1, 1)]
                             + [0.625, 0.75, 0.875])
    rows += [("t_invariance", "T", p, df, None) for df in (1.0, 2.0, 3.0, 5.0, 10.0, 30.0)
             for p in t_probabilities]
    rows += [("beta", "Beta", p, a, b) for a, b in
             ((0.3, 0.4), (0.7, 25.0), (25.0, 0.7), (3.0, 4000.0), (4000.0, 3.0), (60.0, 90.0))
             for p in (1e-9, 0.2, 0.8, 1 - 1e-9)]
    return rows


def build():
    out, counts = [], {}
    for kind, func, p, a, b in points():
        n = counts.get(kind, 0)
        counts[kind] = n + 1
        with mp.workdps(DIGITS):
            ref = mp.nstr(reference(func, p, a, b), 40)
        out.append({"point_id": f"{kind[:2]}{n:03d}", "kind": kind, "func": func,
                    "p_bits": bits(p), "a_bits": bits(a),
                    "b_bits": "NA" if b is None else bits(b), "ref": ref})
    return out


def render(rows):
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=FIELDS, lineterminator="\n", restval="")
    writer.writeheader()
    writer.writerows(rows)
    return buf.getvalue()


def identity(text):
    return [r[:KEY_COLUMNS] for r in csv.reader(io.StringIO(text))]


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    mode = ap.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true")
    mode.add_argument("--check", action="store_true")
    a = ap.parse_args(argv)
    text = render(build())
    for name, path in FILES.items():
        if a.check:
            with open(path, encoding="utf-8") as handle:
                if identity(handle.read()) != identity(text):
                    print(f"FAIL: the identity columns of the {name} file differ from a fresh build")
                    return 1
        elif os.path.exists(path):
            print(f"refusing: {path} exists; it may hold observations. Use --check.")
            return 1
        else:
            with open(path, "w", encoding="utf-8", newline="") as handle:
                handle.write(text)
    print(f"{'PASS' if a.check else 'wrote'}: {len(identity(text)) - 1} rows in each of two files")
    return 0


if __name__ == "__main__":
    sys.exit(main())
