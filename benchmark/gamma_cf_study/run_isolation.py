"""Reproduce the #26 decomposition from committed Excel observations only."""

import argparse
import csv
import hashlib
import json
import math
import mpmath as mp
from model import cf, pref, read, backward, X, D, ROOT, HERE

GRID = ROOT / "benchmark/positive_ratio_subnormal/gamma_q_complement.csv"


def check_model_guards():
    """A capped or out-of-domain evaluation must never become a measurement."""
    for kwargs, expected in [
        ({"a": 1e-4, "x": 2.0, "budget": 1}, RuntimeError),
        ({"a": 1.0, "x": 1.0}, ValueError),
        ({"a": 1.0, "x": math.nan}, ValueError),
        ({"a": 1.0, "x": 2.0, "steps": 0}, ValueError),
    ]:
        try:
            cf(**kwargs)
        except expected:
            continue
        raise AssertionError(f"model silently scored invalid input: {kwargs}")
    for exact in (False, True):
        value, iterations = cf(1.0, 2.0, exact=exact)
        if value != 0.5 or iterations != 1:
            raise AssertionError("shape-one terminating-fraction identity failed")


def expanded_fitting():
    """Exploratory fitting, never labelled independent holdout evidence."""
    rows = []
    with mp.workdps(90):
        for a in [
            1e-6,
            3e-6,
            1e-5,
            3e-5,
            1e-4,
            3e-4,
            0.001,
            0.003,
            0.01,
            0.03,
            0.1,
            0.5,
            0.6,
            1,
            2,
            5,
            10,
            100,
            1000,
            1e4,
            1e6,
        ]:
            for x in [
                a + 1,
                math.nextafter(a + 1, math.inf),
                (a + 1) * 1.25,
                (a + 1) * 2,
                (a + 1) * 4,
                (a + 1) * 16,
            ]:
                q = mp.gammainc(mp.mpf(a), mp.mpf(x), mp.inf, regularized=True)
                if q < mp.mpf(2) ** -1022:
                    rows.append(
                        dict(shape=a, x=x, status="OUTSIDE_NORMAL_OUTPUT_COMPARISON")
                    )
                    continue
                factor, _ = pref(a, x)
                old, old_n = cf(a, x)
                tight, tight_n = cf(a, x, eps=2**-52)
                back = backward(a, x, tight_n)
                errors = {
                    name: mp.nstr(mp.mpf(D(X(factor) * X(h))) / q - 1, 30)
                    for name, h in [
                        ("production", old),
                        ("tight_stop_only", tight),
                        ("backward_candidate", back),
                    ]
                }
                rows.append(
                    dict(
                        shape=a,
                        x=x,
                        status="SCORED",
                        old_iterations=old_n,
                        tight_iterations=tight_n,
                        relative_errors=errors,
                    )
                )
    return rows


def measure():
    result = []
    with mp.workdps(90), GRID.open() as source:
        for row in csv.DictReader(source):
            if not row["point_id"].startswith("x2_"):
                continue
            a, x = read(row["shape"]), read(row["x"])
            reference = mp.gammainc(mp.mpf(a), mp.mpf(x), mp.inf, regularized=True)
            with mp.workdps(60):
                low_precision = mp.gammainc(
                    mp.mpf(a), mp.mpf(x), mp.inf, regularized=True
                )
            quadrature = mp.quad(
                lambda t: mp.exp(-t) * t ** (mp.mpf(a) - 1), [mp.mpf(x), mp.inf]
            ) / mp.gamma(a)
            stability = abs(low_precision / reference - 1)
            independent = abs(quadrature / reference - 1)
            if stability > mp.mpf("1e-55") or independent > mp.mpf("1e-75"):
                raise RuntimeError(
                    "reference precision or independent quadrature check failed"
                )
            factor_ref = mp.exp(mp.mpf(a) * mp.log(x) - x - mp.loggamma(a))
            h, n = cf(a, x)
            finite, _ = cf(a, x, steps=n, exact=True)
            factor, _ = pref(a, x)
            model = D(X(factor) * X(h))
            observed = read(row["survival"])
            tight_h, tight_n = cf(a, x, eps=2**-52)
            tight = D(X(factor) * X(tight_h))
            parts = {
                "observed_relative_error": mp.mpf(observed) / reference - 1,
                "prefactor_relative_error": mp.mpf(factor) / factor_ref - 1,
                "finite_fraction_truncation": finite / (reference / factor_ref) - 1,
                "recurrence_rounding": mp.mpf(h) / finite - 1,
                "final_multiply_rounding": mp.mpf(model) / (mp.mpf(factor) * h) - 1,
                "tight_stop_only_relative_error": mp.mpf(tight) / reference - 1,
            }
            reconstructed = (
                mp.fprod(
                    1 + parts[k]
                    for k in (
                        "prefactor_relative_error",
                        "finite_fraction_truncation",
                        "recurrence_rounding",
                        "final_multiply_rounding",
                    )
                )
                - 1
            )
            if abs(reconstructed - parts["observed_relative_error"]) > mp.mpf("1e-75"):
                raise RuntimeError(
                    "component reconstruction does not match observed error"
                )
            result.append(
                dict(
                    point_id=row["point_id"],
                    shape=a,
                    x=x,
                    iterations=n,
                    tight_iterations=tight_n,
                    model_matches_excel=model == observed,
                    reference_60_90_relative_difference=mp.nstr(stability, 15),
                    quadrature_relative_difference=mp.nstr(independent, 15),
                    **{k: mp.nstr(v, 30) for k, v in parts.items()},
                )
            )
    if len(result) != 12 or not all(r["model_matches_excel"] for r in result):
        raise RuntimeError("model must reproduce all twelve Excel observations exactly")
    paths = [
        GRID,
        HERE / "model.py",
        HERE / "run_isolation.py",
        HERE.parent / "beta_balanced_survival/beta_route_model.py",
        ROOT / "src/M_STATS_PROBDIST_SPECIALFUNCS.bas",
        ROOT / "src/M_STATS_PROBDIST_CORE.bas",
    ]
    return dict(
        status="DIAGNOSIS ONLY - NOT A FIX OR CLOSURE",
        precision_dps=90,
        bindings={
            str(p.relative_to(ROOT)): hashlib.sha256(
                p.read_bytes().replace(b"\r\n", b"\n")
            ).hexdigest()
            for p in paths
        },
        rows=result,
        expanded_fitting=expanded_fitting(),
    )


def main():
    check_model_guards()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args()
    expected = json.dumps(measure(), indent=2) + "\n"
    output = HERE / "isolation.json"
    if args.write:
        output.write_text(expected)
    elif not output.exists() or output.read_text() != expected:
        raise SystemExit("FAIL: isolation report is absent or stale; run --write")
    print(
        "PASS: 12/12 exact Excel reproductions; multiplicative error decomposition reconstructs all observations"
    )


if __name__ == "__main__":
    main()
