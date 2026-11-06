"""
Fit the #72 candidate normal-tail kernel exactly as PREREGISTRATION.md §2
fixes it, decide whether the fit converged (Amendment 3, #79), and record the
result.

The rational P(z)/D(z), degree (7, 7) with D(0) = 1, approximates

    f(z) = Q(z) * exp(z^2 / 2),    0 <= z <= 3,

where Q is the standard-normal upper tail. The fit is deterministic:
linearised least squares for relative error at the 500 Chebyshev nodes of
[0, 3], weights 1 / (f * D_prev), Lawson reweighting by the node errors,
120 iterations at 60 digits, keeping the iterate with the smallest node
maximum.

Convergence (Amendment 3). Every iterate's signed relative error at the nodes
is recorded: its node maximum, the number of sign-alternating extrema, and
the levelling ratio node_max / L - 1, where L is the de la Vallee Poussin
lower bound on the best node maximum any (7, 7) rational can reach (the
largest t such that the extrema of size >= t still alternate at 7 + 7 + 2 =
16 points). The selected iterate is

  CONVERGED      if it is finite, D > 0 at every node, its error alternates at
                 16 or more points and its levelling ratio is <= 1%;
  CAP_EXHAUSTED  if it is not, and it is the last iterate: the node maximum
                 was still falling when the frozen cap of 120 stopped the fit;
  NOT_CONVERGED  otherwise: the iteration stalled or oscillated without
                 levelling, or produced a non-finite or nonpositive iterate.

Only a CONVERGED fit is recommended. Any other status, or a missing one, is
refused before kernel_coefficients.json is written or accepted, and exits 1.

Every constant is also recorded the way VBA must spell it. The VBE keeps a
numeric literal to 15 significant digits, so a constant is the Double sum
hi + lo of two 15-digit literals. hi is the target Double rounded to 15
digits, and lo is the exact remainder rounded to 15 digits. The pair is
accepted only if the sum is bit-identical to the target Double.

    python derive_kernel.py
        Fits, decides convergence and records it in fit_convergence.json.
        Writes kernel_coefficients.json only when it does not exist. When it
        exists, the fitted coefficients must reproduce it exactly (every
        field but runtime_s); the file is then left byte-identical. A fit
        that would change it is refused and nothing is written. A
        convergence record that reproduces apart from runtime_s is also left
        unchanged. A fit that is not CONVERGED writes its convergence record,
        bound to no coefficient file, and nothing else.
    python derive_kernel.py --verify
        Read-only: fits again and compares the coefficients with
        kernel_coefficients.json and the convergence record with
        fit_convergence.json. Writes nothing.

Exit status: 0 converged and recorded (or reproduced), 1 not converged,
missing status, or a mismatch.
"""
import argparse
import hashlib
import json
import math
import os
import struct
import sys
import time
from fractions import Fraction

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "kernel_coefficients.json")
CONVERGENCE_OUT = os.path.join(HERE, "fit_convergence.json")

LOWER, UPPER = 0, 3
DEGREE_P = DEGREE_D = 7
NODES = 500
ITERATIONS = 120
DPS = 60

# Amendment 3 (#79): the convergence criterion.
CONVERGED, NOT_CONVERGED, CAP_EXHAUSTED = "CONVERGED", "NOT_CONVERGED", "CAP_EXHAUSTED"
STATUSES = (CONVERGED, NOT_CONVERGED, CAP_EXHAUSTED)
LEVELLING_TOLERANCE = 0.01
CRITERION = (
    "The selected iterate (smallest node maximum of the frozen 120) is CONVERGED when its signed relative "
    "error at the 500 nodes is finite, D > 0 at every node, the error alternates in sign at >= m + n + 2 = 16 "
    "extrema, and node_max / L - 1 <= 0.01, where L is the de la Vallee Poussin lower bound: the largest t "
    "such that the run extrema of size >= t still alternate at 16 points. Otherwise it is CAP_EXHAUSTED when "
    "the selected iterate is the last one (the node maximum was still falling at the cap), and NOT_CONVERGED "
    "when it is not. Only CONVERGED is recommended (PREREGISTRATION.md Amendment 3).")


def target(z):
    """f(z) = Q(z) * exp(z^2/2) at the working precision."""
    z = mp.mpf(z)
    return mp.erfc(z / mp.sqrt(2)) / 2 * mp.exp(z * z / 2)


def chebyshev_nodes(a, b, n):
    a, b = mp.mpf(a), mp.mpf(b)
    return [(a + b) / 2 + (b - a) / 2 * mp.cos(mp.pi * (i + mp.mpf(1) / 2) / n)
            for i in range(n)]


def horner(coefficients, x):
    """coefficients in ascending powers."""
    acc = mp.mpf(0)
    for c in reversed(coefficients):
        acc = acc * x + c
    return acc


# ---------------------------------------------------------------------------
# Convergence measures (Amendment 3)
# ---------------------------------------------------------------------------

def alternant(xs, errors):
    """The run extrema of a signed error, in increasing x.

    The nodes are split into maximal runs of one sign (a zero error belongs to
    no run); each run contributes its largest |error|. Consecutive extrema
    therefore alternate in sign.
    """
    runs = []
    for i in sorted(range(len(xs)), key=lambda k: xs[k]):
        e = errors[i]
        if e == 0:
            continue
        if runs and (runs[-1] > 0) == (e > 0):
            if abs(e) > abs(runs[-1]):
                runs[-1] = e
        else:
            runs.append(e)
    return runs


def levelled_bound(runs, points):
    """de la Vallee Poussin: the largest t such that the extrema of size >= t
    alternate at `points` points, or None if they never do.

    If the error of r alternates in sign at m + n + 2 points with sizes >= t,
    no rational of type (m, n) has a smaller node maximum than t; so t bounds
    the best node maximum from below.
    """
    for t in sorted({abs(e) for e in runs}, reverse=True):
        kept = [e > 0 for e in runs if abs(e) >= t]
        count = 1 + sum(1 for a, b in zip(kept, kept[1:]) if a != b) if kept else 0
        if count >= points:
            return t
    return None


def _finite(x):
    try:
        return math.isfinite(float(x))
    except (OverflowError, TypeError, ValueError):
        return False


def measure(xs, signed, denominators, points):
    """Convergence evidence of one iterate from its signed node errors."""
    finite = all(_finite(e) for e in signed) and all(_finite(d) for d in denominators)
    node_max = max(abs(e) for e in signed) if finite else None
    runs = alternant(xs, signed) if finite else []
    bound = levelled_bound(runs, points) if finite else None
    return {
        "node_max": float(node_max) if node_max is not None else None,
        "alternation_points": len(runs),
        "levelled_min": float(bound) if bound is not None else None,
        "levelling": float(node_max / bound - 1) if bound else None,
        "denominator_min": float(min(denominators)) if finite else None,
    }


def levelled(m, points, tolerance=LEVELLING_TOLERANCE):
    """The criterion of Amendment 3 on one iterate's evidence."""
    return (isinstance(m, dict) and m.get("node_max") is not None and m.get("levelling") is not None
            and m.get("denominator_min") is not None and m["denominator_min"] > 0
            and isinstance(m.get("alternation_points"), int) and m["alternation_points"] >= points
            and 0 <= m["levelling"] <= tolerance)


def decide(history, selected, points, tolerance=LEVELLING_TOLERANCE):
    """CONVERGED, CAP_EXHAUSTED or NOT_CONVERGED for the selected iterate.

    history is the list of per-iterate evidence, in iteration order; selected
    is the 1-based iteration the fit returns.
    """
    if not history or not isinstance(selected, int) or not 1 <= selected <= len(history):
        return NOT_CONVERGED
    if levelled(history[selected - 1], points, tolerance):
        return CONVERGED
    return CAP_EXHAUSTED if selected == len(history) else NOT_CONVERGED


def recommendable(record):
    """Reasons a convergence record does not allow recommending its candidate; [] if it does."""
    if not isinstance(record, dict):
        return ["no convergence record"]
    status = record.get("status")
    if status is None:
        return ["the convergence record has no status"]
    if status not in STATUSES:
        return [f"convergence status {status!r} is not one of {', '.join(STATUSES)}"]
    if status != CONVERGED:
        return [f"the fit is {status}: {record.get('reason', 'no reason recorded')}"]
    return []


# ---------------------------------------------------------------------------
# The fit (§2, unchanged)
# ---------------------------------------------------------------------------

def fit(lower=LOWER, upper=UPPER, degree_p=DEGREE_P, degree_d=DEGREE_D, nodes=NODES,
        iterations=ITERATIONS, f=target):
    """Linearised relative least squares with Lawson reweighting.

    Returns (node_max, p, d, convergence). The arithmetic and the selection
    are those of §2; the convergence evidence is only observed, never used to
    steer or stop the iteration.
    """
    xs = chebyshev_nodes(lower, upper, nodes)
    fs = [f(x) for x in xs]
    lam = [mp.mpf(1)] * nodes
    d_prev = [mp.mpf(1)] * nodes
    points = degree_p + degree_d + 2
    best, history = None, []
    for iteration in range(1, iterations + 1):
        rows, rhs = [], []
        for x, fx, weight, dp in zip(xs, fs, lam, d_prev):
            w = mp.sqrt(weight) / (fx * dp)
            # P(x) - f(x) * (D(x) - 1) = f(x), with D(0) = 1.
            rows.append([w * x ** i for i in range(degree_p + 1)]
                        + [-w * fx * x ** j for j in range(1, degree_d + 1)])
            rhs.append(w * fx)
        solution = mp.qr_solve(mp.matrix(rows), mp.matrix(rhs))[0]
        p = [solution[i] for i in range(degree_p + 1)]
        d = [mp.mpf(1)] + [solution[degree_p + j] for j in range(1, degree_d + 1)]
        denominators = [horner(d, x) for x in xs]
        signed = [horner(p, x) / dx / fx - 1 for x, dx, fx in zip(xs, denominators, fs)]
        errors = [abs(e) for e in signed]
        history.append({"iteration": iteration, **measure(xs, signed, denominators, points)})
        worst = max(errors)
        if best is None or worst < best[0]:
            best = (worst, p, d, iteration)
        d_prev = denominators
        total = sum(weight * e for weight, e in zip(lam, errors))
        lam = [weight * e / total for weight, e in zip(lam, errors)]
    node_max, p, d, selected = best
    return node_max, p, d, convergence(history, selected, points, iterations)


def convergence(history, selected, points, iterations):
    status = decide(history, selected, points)
    first = next((h["iteration"] for h in history if levelled(h, points)), None)
    chosen = history[selected - 1] if 1 <= selected <= len(history) else None
    if status == CONVERGED:
        reason = (f"iterate {selected} alternates at {chosen['alternation_points']} points and is levelled to "
                  f"{chosen['levelling']:.3e} <= {LEVELLING_TOLERANCE}")
    elif status == CAP_EXHAUSTED:
        reason = f"the node maximum was still falling at the cap of {iterations} iterations without levelling"
    else:
        reason = "the selected iterate is not levelled, and the iteration had stopped improving before the cap"
    tail = [h["node_max"] for h in history[-10:]]
    spread = max(tail) / min(tail) - 1 if tail and None not in tail and min(tail) > 0 else None
    return {
        "status": status,
        "reason": reason,
        "criterion": {"rule": CRITERION, "alternation_points_required": points,
                      "levelling_tolerance": LEVELLING_TOLERANCE, "iteration_cap": iterations,
                      "selection": "smallest node maximum; the earlier iterate on a tie"},
        "selected_iteration": selected,
        "first_levelled_iteration": first,
        "selected": chosen,
        "final_ten_node_max_spread": spread,
        "history": history,
    }


# ---------------------------------------------------------------------------
# Records
# ---------------------------------------------------------------------------

def double_hex(x):
    return struct.pack(">d", x).hex()


def vba_literal_pair(x):
    """(hi, lo) 15-significant-digit literals whose Double sum is exactly x."""
    hi_text = format(x, ".15g")
    remainder = Fraction(x) - Fraction(float(hi_text))
    lo_text = "0" if remainder == 0 else format(float(remainder), ".15g")
    return hi_text, lo_text, float(hi_text) + float(lo_text) == x


def constant_record(value):
    x = float(value)                       # mpmath rounds to nearest
    hi, lo, exact = vba_literal_pair(x)
    return {"value": mp.nstr(value, 40), "double": repr(x), "double_hex": double_hex(x),
            "hi": hi, "lo": lo, "pair_is_exact": exact}


def coefficient_record(node_max, p, d, runtime):
    inv_sqrt_two_pi = 1 / mp.sqrt(2 * mp.pi)
    return {
        "study": "normal upper-tail kernel (#72)",
        "preregistration": "PREREGISTRATION.md",
        "target": "f(z) = Q(z) * exp(z^2/2) on [0, 3]",
        "fit": {"interval": [LOWER, UPPER], "degree": [DEGREE_P, DEGREE_D], "nodes": NODES,
                "iterations": ITERATIONS, "dps": DPS, "node_max_relative_error": mp.nstr(node_max, 6)},
        "numerator": [constant_record(c) for c in p],
        "denominator": [constant_record(c) for c in d],
        "inv_sqrt_two_pi": constant_record(inv_sqrt_two_pi),
        "split": 3.0,
        "continued_fraction_levels": 60,
        "exponent_split_scale": 256,
        "runtime_s": runtime,
    }


def without_runtime(record):
    return {k: v for k, v in record.items() if k != "runtime_s"}


def file_sha256(path):
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


def dump(record):
    return json.dumps(record, indent=1) + "\n"


def write(path, text):
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)


def convergence_record(conv, coefficients_sha256, runtime):
    return {"study": "normal upper-tail kernel (#72)", "preregistration": "PREREGISTRATION.md (Amendment 3)",
            "kernel_coefficients_sha256": coefficients_sha256, **conv, "runtime_s": runtime}


def main(argv=None, out=OUT, convergence_out=CONVERGENCE_OUT, fitter=fit, log=print):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--verify", action="store_true",
                        help="fit again and compare with the committed files; write nothing")
    args = parser.parse_args(argv)
    mp.mp.dps = DPS
    started = time.time()
    node_max, p, d, decision = fitter()
    runtime = round(time.time() - started, 1)
    record = coefficient_record(node_max, p, d, runtime)
    refusals = recommendable(decision)
    conv = decision if isinstance(decision, dict) else {}
    log(f"node max relative error {mp.nstr(node_max, 4)}; fit {conv.get('status', 'WITHOUT STATUS')} at iterate "
        f"{conv.get('selected_iteration')} ({conv.get('reason', 'no reason recorded')}); {runtime} s")
    existing = None
    if os.path.exists(out):
        with open(out, encoding="utf-8") as f:
            existing = json.load(f)
        if without_runtime(existing) != without_runtime(record):
            refusals.append(f"the fitted coefficients differ from the recorded {os.path.basename(out)}; "
                            "a rerun must not replace a recorded candidate")
    elif args.verify:
        refusals.append(f"{os.path.basename(out)} does not exist")

    if args.verify:
        if not refusals:
            sha = file_sha256(out)
            try:
                with open(convergence_out, encoding="utf-8") as f:
                    recorded = json.load(f)
            except (OSError, ValueError) as e:
                recorded = None
                refusals.append(f"{os.path.basename(convergence_out)} is missing or unreadable ({e})")
            expected = without_runtime(convergence_record(conv, sha, None))
            if recorded is not None and without_runtime(recorded) != expected:
                refusals.append(f"{os.path.basename(convergence_out)} does not match the convergence of this fit")
        for r in refusals:
            log(f"  REFUSED: {r}")
        log("VERIFIED: the fit reproduces the recorded coefficients and convergence record; nothing written"
            if not refusals else "NOT VERIFIED; nothing written")
        return 0 if not refusals else 1

    if refusals:
        if recommendable(decision):
            # Record the failure for inspection, bound to no candidate.
            write(convergence_out, dump(convergence_record(conv, None, runtime)))
            log(f"  wrote {os.path.basename(convergence_out)} with status {conv.get('status')}")
        for r in refusals:
            log(f"  REFUSED: {r}")
        log(f"no candidate recommended; {os.path.basename(out)} was not written")
        return 1
    if existing is None:
        write(out, dump(record))
        log(f"wrote {os.path.basename(out)}")
    else:
        log(f"{os.path.basename(out)} reproduced exactly; left unchanged")
    fresh = convergence_record(conv, file_sha256(out), runtime)
    try:
        with open(convergence_out, encoding="utf-8") as f:
            unchanged = without_runtime(json.load(f)) == without_runtime(fresh)
    except (OSError, ValueError):
        unchanged = False
    if unchanged:
        log(f"{os.path.basename(convergence_out)} reproduced exactly; left unchanged")
    else:
        write(convergence_out, dump(fresh))
        log(f"wrote {os.path.basename(convergence_out)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
