"""
#314: development replay of the routed Gamma P and Q on revealed data only.

It never reads holdout_design.py or the routing holdout grid. Four sets:

  fit      the 392 points of the #307 fit grid, with their references;
  main     the Gamma, Chi-square and Poisson cumulative and survival rows of
           the main grid, each against the guarantee of its contract;
  holdout  the same functions in the consolidated holdout grid;
  sweep    2,200 seeded random points over the whole domain, shapes from
           1E-300 to 1E5, with references computed here at 50 digits.

Every value is scored as the relative error of the model's routed result
(fast_route.routed) against the reference. The rows of the positive-ratio
regimes of #13 are left out: their dispatch is in the public surfaces, above
the routed entry points.

Run:
    python3 replay_fast_route.py --write    writes fast_route_development.json
    python3 replay_fast_route.py --check    recomputes the fit replay and
                                            compares it with the record; the
                                            fit grid is frozen. The two live
                                            grids are replayed afresh and only
                                            required to stay within their
                                            guarantees: they gain rows over
                                            time, so their recorded summary is
                                            a snapshot, not a binding (#302).
                                            The sweep is checked by its
                                            recorded summary only.
Exit 0 only when no value exceeds its bound.
"""
import argparse
import csv
import hashlib
import json
import math
import os
import random
import sys

from mpmath import gammainc, mp, mpf

HERE = os.path.dirname(os.path.abspath(__file__))
BENCH = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import fast_route as F  # noqa: E402

OUT = os.path.join(HERE, "fast_route_development.json")
FLOOR_BOUND = 1e-13
SWEEP_SEED = 314
FUNCTIONS = {f + s for f in ("Gamma_", "ChiSquare_", "Poisson_") for s in ("Cumulative", "Survival")}
INPUTS = ("fast_route.py", "route_model.py", "b64_fit_grid.csv", "replay_fast_route.py")


def relative(value, ref):
    return float(abs(mpf(value) - ref) / abs(ref))


def score(worst, counts, key, error, where):
    counts[key] = counts.get(key, 0) + 1
    if error > worst.get(key, {"error": -1.0})["error"]:
        worst[key] = {"error": error, **where}


def fit():
    worst, counts, failures = {}, {}, []
    with open(os.path.join(HERE, "b64_fit_grid.csv"), newline="") as handle:
        for r in csv.DictReader(handle):
            a, x = float(r["shape"]), float(r["x"])
            p, q = F.routed(a, x)
            sp, sq = F.paths(a, x)
            for tail, value, ref, source in (("P", p, mpf(r["ref_p"]), sp), ("Q", q, mpf(r["ref_q"]), sq)):
                if ref < mpf("1e-300"):
                    continue
                e = relative(value, ref)
                score(worst, counts, f"{tail}:{source}", e, {"a": a.hex(), "x": x.hex()})
                if e > FLOOR_BOUND:
                    failures.append({"a": a.hex(), "x": x.hex(), "tail": tail, "error": e})
    return {"worst": worst, "counts": counts, "failures": failures}


def _guarantees():
    with open(os.path.join(BENCH, "accuracy_guarantees.csv"), newline="") as handle:
        g = {r["contract_id"]: float(r["guarantee"]) for r in csv.DictReader(handle)
             if r["guarantee"].strip()}
    with open(os.path.join(BENCH, "accuracy_contracts.csv"), newline="") as handle:
        return {(r["function"], r["regime"]): g[r["contract_id"]] for r in csv.DictReader(handle)
                if r["measure"] == "output_error" and r["contract_id"] in g}


def kernel(function, row):
    """(A, X, wants P) of the routed call behind a public row."""
    a1, a2 = float(row["arg1"]), float(row["arg2"])
    if function.startswith("Gamma"):
        return a2, a1 / float(row["arg3"]), function.endswith("Cumulative")
    if function.startswith("ChiSquare"):
        return 0.5 * a2, 0.5 * a1, function.endswith("Cumulative")
    # Poisson: CDF(k, mean) = Q(k + 1, mean), survival = P(k + 1, mean).
    return math.floor(a1) + 1.0, a2, not function.endswith("Cumulative")


def grid(rel):
    guarantees = _guarantees()
    worst, counts, failures = {}, {}, []
    with open(os.path.join(BENCH, rel), newline="") as handle:
        for r in csv.DictReader(handle):
            fn, regime = r["function"], r.get("regime") or "all"
            if fn not in FUNCTIONS or (r.get("expected_error") or "").strip():
                continue
            if "positive_ratio" in regime:
                continue
            ref = mpf(r["reference"])
            if ref == 0:
                continue
            a, x, wants_p = kernel(fn, r)
            if not (a > 0 and x > 0):
                continue
            p, q = F.routed(a, x)
            e = relative(p if wants_p else q, ref)
            bound = guarantees.get((fn, regime), FLOOR_BOUND)
            score(worst, counts, f"{fn}.{regime}", e, {"a": a.hex(), "x": x.hex(), "bound": bound})
            if e > bound:
                failures.append({"function": fn, "regime": regime, "a": a.hex(), "x": x.hex(),
                                 "error": e, "bound": bound})
    return {"worst": worst, "counts": counts, "failures": failures}


def sweep_points():
    rng = random.Random(SWEEP_SEED)
    pts = []
    for _ in range(900):                       # log-uniform over the whole domain
        pts.append((10 ** rng.uniform(-300, 5), 10 ** rng.uniform(-300, 6)))
    for _ in range(900):                       # ordinary shapes, around the body
        a = 10 ** rng.uniform(-8, 5)
        z = rng.uniform(-12, 12)
        x = a + z * math.sqrt(a) if a >= 1 else 10 ** rng.uniform(-12, 2.5)
        if x > 0:
            pts.append((a, x))
    for _ in range(400):                       # deep tails of ordinary shapes
        a = 10 ** rng.uniform(-3, 4)
        pts.append((a, a * 10 ** rng.uniform(-6, 3)))
    return pts


def sweep():
    mp.dps = 50
    worst, counts, failures = {}, {}, []
    for a, x in sweep_points():
        sources = F.paths(a, x)
        if sources == ("dd", "dd"):
            counts["delegated"] = counts.get("delegated", 0) + 1
            continue
        routine, value = F.fast(a, x)
        fa, fx = mpf(a), mpf(x)
        refs = {"P": gammainc(fa, 0, fx, regularized=True),
                "Q": gammainc(fa, fx, mp.inf, regularized=True)}
        for tail, source in zip("PQ", sources):
            if source == "dd":
                continue
            v = value if source in ("series", "cf") else F.rm.sub(1.0, value)
            ref = refs[tail]
            if ref < mpf("1e-300"):
                continue
            e = relative(v, ref)
            score(worst, counts, f"{tail}:{source}", e, {"a": a.hex(), "x": x.hex()})
            if e > FLOOR_BOUND:
                failures.append({"a": a.hex(), "x": x.hex(), "tail": tail, "error": e})
    return {"seed": SWEEP_SEED, "points": len(sweep_points()), "worst": worst, "counts": counts,
            "failures": failures}


def hashes():
    out = {}
    for name in INPUTS:
        with open(os.path.join(HERE, name), "rb") as handle:
            out[name] = hashlib.sha256(handle.read().replace(b"\r\n", b"\n")).hexdigest()
    return out


def light():
    mp.dps = 60
    return {"fit": fit(), "main": grid("probability_accuracy_grid.csv"),
            "holdout": grid(os.path.join("holdout", "holdout_grid.csv"))}


def failed(report):
    return sum(len(report[k]["failures"]) for k in ("fit", "main", "holdout", "sweep"))


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    mode = ap.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true")
    mode.add_argument("--check", action="store_true")
    a = ap.parse_args(argv)
    if a.write:
        report = {"kind": "development replay on revealed data; not the routing holdout",
                  "rule": {"floor": F.FLOOR, "complement_max": F.COMPLEMENT_MAX,
                           "upper_max": F.UPPER_MAX},
                  **light(), "sweep": sweep(), "sha256": hashes()}
        with open(OUT, "w", encoding="utf-8") as handle:
            json.dump(report, handle, indent=2, sort_keys=True)
            handle.write("\n")
    else:
        with open(OUT, encoding="utf-8") as handle:
            report = json.load(handle)
        fresh = json.loads(json.dumps(light(), sort_keys=True))
        if report["fit"] != fresh["fit"]:
            print("FAIL: the committed fit replay differs from a fresh one")
            return 1
        report = {**report, "main": fresh["main"], "holdout": fresh["holdout"]}
    n = failed(report)
    top = max(v["error"] for k in ("fit", "main", "holdout", "sweep") for v in report[k]["worst"].values())
    print(f"{'FAIL' if n else 'PASS'}: {n} value(s) beyond their bound; worst relative error {top:.2e}")
    return 1 if n else 0


if __name__ == "__main__":
    sys.exit(main())
