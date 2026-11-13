"""Main-grid promotion of the modelled worst balanced Beta points (#327).

#327 sets the balanced Beta CDF and survival guarantee at 2E-13. The bound
rests on beta_route_model.py, which reproduces the route bit for bit (368 of
368 Excel observations, #174) and reaches 1.38E-13 at shapes 20 to 100. This
promotion puts the points where the model is worst into the main grid, so
Excel measures them and the strict gate holds them to the guarantee on every
change from then on.

Points, chosen from results.json by modelled error, never by an Excel
measurement:

    (0.99, 14, 19)        domain sweep, worst survival      6.70E-14
    (0.7, 19, 14)         domain sweep, survival            5.71E-14
    (0.3, 14, 19)         domain sweep, worst CDF           5.94E-14
    (0.410058, 50, 100)   large sweep, worst survival       1.38E-13
    (0.333333, 50, 100)   large sweep, CDF                  1.30E-13
    (0.666667, 100, 50)   large sweep, worst CDF            1.34E-13

Each point gives a Beta_Survival and a Beta_Cumulative row in the balanced
regime (shape ratio >= 0.1; large_shape needs both shapes >= 100). The large
sweep's x values are its own six-decimal tokens. References are certified at
the exact binary64 arguments: the survival by _survival_oracle.certify, the
CDF by the same construction (measure_balanced_survival.cdf_reference).

Use with benchmark/promote_grid_rows.py --generator; never copy observations.

Run: python3 promote_worst_points.py   (prints the rows and modelled errors)
"""
from pathlib import Path
import sys

import mpmath as mp

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent))
import measure_balanced_survival as S  # noqa: E402

POINTS = (
    ("0.99", "14.0", "19.0"),
    ("0.7", "19.0", "14.0"),
    ("0.3", "14.0", "19.0"),
    ("0.410058", "50.0", "100.0"),
    ("0.333333", "50.0", "100.0"),
    ("0.666667", "100.0", "50.0"),
)
CLAIMS = {"Beta_Survival": "rel<=5E-15", "Beta_Cumulative": "rel<=2E-14"}
GUARANTEE = mp.mpf("2E-13")
DIGITS = 30


def reference(function, x, a, b):
    xs, av, bv = float(x), float(a), float(b)
    value = (S.survival_reference(xs, av, bv) if function == "Beta_Survival"
             else S.cdf_reference(xs, av, bv))
    with mp.workdps(60):
        return mp.nstr(value, DIGITS, strip_zeros=False)


def modelled(function, x, a, b):
    """The binary64 value beta_route_model.py predicts Excel returns."""
    route = S.M.beta_survival if function == "Beta_Survival" else S.M.beta_cumulative
    return route(float(x), float(a), float(b), trace={})


def build_rows():
    rows = []
    for x, a, b in POINTS:
        for function in ("Beta_Cumulative", "Beta_Survival"):
            rows.append({"function": function, "vba_kernel": "K_STATS_" + function,
                         "claim": CLAIMS[function], "metric": "rel", "arg1": x, "arg2": a,
                         "arg3": b, "arg4": "", "reference": reference(function, x, a, b),
                         "observed_vba": "", "regime": "balanced",
                         "evidence_set": "main grid", "expected_error": ""})
    return rows


if __name__ == "__main__":
    for r in build_rows():
        m = modelled(r["function"], r["arg1"], r["arg2"], r["arg3"])
        with mp.workdps(60):
            ref = mp.mpf(r["reference"])
            err = abs(mp.mpf(m) - ref) / abs(ref)
        print(f"{r['function']:<16} ({r['arg1']}, {r['arg2']}, {r['arg3']})  "
              f"ref {r['reference'][:22]}  modelled error {mp.nstr(err, 3)}")
