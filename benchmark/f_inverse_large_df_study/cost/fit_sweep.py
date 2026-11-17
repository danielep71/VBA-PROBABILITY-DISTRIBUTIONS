"""Fitting sweep for #359: candidate S against the accepted rule P4, in the model.

Every point here is revealed: it informs the registration and is not evidence
for it. The points are a fixed product of shapes and probabilities, not a
random draw, so the sweep is reproducible without a seed. References are
computed with mpmath at 60 digits by solving the regularized incomplete beta
on the smaller probability tail.

Candidate S was adopted on 2026-11-16 (RESULT.md); the names "candidate" and
"accepted rule" here are those of the time of the sweep.

Writes fit_result.json beside this file. Run: python fit_sweep.py
"""
import json
import math
import sys
from pathlib import Path

from mpmath import betainc, exp, log, loggamma, mp, mpf

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import selective_model as S  # noqa: E402

p4 = S.p4
mp.dps = 60

BETA_SHAPES = [0.5, 0.75, 1.0, 1.5, 2.0, 5.0, 12.5, 50.0, 80.0, 300.0, 2500.0]
F_DF = [(1, 1), (1, 30), (2, 9), (5, 10), (10, 5), (30, 1), (100, 100), (7, 3e5), (3e5, 7),
        (2.5, 1e6), (1e6, 3), (40, 2e7), (2e7, 40), (50, 1e8), (1e8, 50), (100, 1e10), (1e10, 100),
        (1e3, 1e3), (1e5, 1e5), (3, 3e3)]
PROBABILITIES = [1e-12, 1e-8, 1e-4, 0.01, 0.1, 0.3, 0.5, 0.7, 0.9, 0.99, 1 - 1e-4, 1 - 1e-8, 1 - 1e-12]


def reference(p, complement, a, b, start):
    """x with I_x(a, b) = p: Newton at 60 digits on the smaller tail, from the
    model's own answer, which is already within about 1E-14 of the root."""
    a, b, x = mpf(a), mpf(b), mpf(start)
    log_beta = loggamma(a) + loggamma(b) - loggamma(a + b)
    for _ in range(8):
        if p <= complement:
            value = betainc(a, b, 0, x, regularized=True) - mpf(p)
        else:
            value = mpf(complement) - betainc(a, b, x, 1, regularized=True)
        density = exp((a - 1) * log(x) + (b - 1) * log(1 - x) - log_beta)
        step = value / density
        x -= step
        if abs(step) <= abs(x) * mpf(10) ** -45:
            return x
    raise ValueError("reference did not converge")


def measure(kind, p, s1, s2):
    complement = 1.0 - p
    a, b = (s1, s2) if kind == "Beta" else (s1 / 2.0, s2 / 2.0)
    row = {"kind": kind, "p": p, "shape1": s1, "shape2": s2}
    try:
        S.reset_counts()
        x4, y4, _ = p4.inverse_pair(p, complement, a, b)
        row["p4_pair"] = S.COUNTS["pair"]
        row["p4_pair_terms"] = S.COUNTS["pair_terms"]
    except Exception as error:  # the model does not cover every branch of the source
        row["skipped"] = "P4: " + str(error)[:80]
        return row
    try:
        S.reset_counts()
        xs, ys, used_pair = S.inverse_pair(p, complement, a, b)
        row.update(s_pair=S.COUNTS["pair"], s_pair_terms=S.COUNTS["pair_terms"],
                   s_binary64=S.COUNTS["binary64"], s_used_pair=used_pair)
    except Exception as error:
        row["skipped"] = "S: " + str(error)[:80]
        return row
    # Solve for the member of the pair below 1/2, so that the other member is
    # one minus a small number and keeps its digits at 60 dps.
    try:
        if x4 <= 0.5:
            small = reference(p, complement, a, b, x4)
            exact_x, exact_y = small, 1 - small
        else:
            small = reference(complement, p, b, a, y4)
            exact_x, exact_y = 1 - small, small
    except Exception as error:
        row["skipped"] = "reference: " + str(error)[:80]
        return row
    if kind == "F":
        truth = mpf(s2) / mpf(s1) * exact_x / exact_y
        value4, value_s = S.f_quantile(p, s1, s2, x4, y4), S.f_quantile(p, s1, s2, xs, ys)
    else:
        truth, value4, value_s = exact_x, x4, xs
    row["p4_error"] = float(abs((mpf(value4) - truth) / truth))
    row["s_error"] = float(abs((mpf(value_s) - truth) / truth))
    row["same_double"] = value4 == value_s
    return row


def main():
    points = [("Beta", p, a, b) for a in BETA_SHAPES for b in BETA_SHAPES for p in PROBABILITIES]
    points += [("F", p, d1, d2) for d1, d2 in F_DF for p in PROBABILITIES]
    rows = [measure(*point) for point in points]
    done = [row for row in rows if "skipped" not in row]
    summary = {
        "points": len(rows), "measured": len(done), "skipped": len(rows) - len(done),
        "same_double": sum(row["same_double"] for row in done),
        "p4_worst": max(row["p4_error"] for row in done),
        "s_worst": max(row["s_error"] for row in done),
        "largest_move_away": max(row["s_error"] - row["p4_error"] for row in done),
        "rows_farther_by_more_than_1e-14": sum(row["s_error"] - row["p4_error"] > 1e-14 for row in done),
        "rows_above_1e-12": sum(row["s_error"] > 1e-12 for row in done),
        "s_without_pair": sum(not row["s_used_pair"] for row in done),
        "p4_pair_evaluations": sum(row["p4_pair"] for row in done),
        "s_pair_evaluations": sum(row["s_pair"] for row in done),
        "s_binary64_evaluations": sum(row["s_binary64"] for row in done),
    }
    (HERE / "fit_result.json").write_text(json.dumps({"summary": summary, "rows": rows}, indent=1) + "\n")
    print(json.dumps(summary, indent=1))
    assert math.isfinite(summary["s_worst"])


if __name__ == "__main__":
    main()
