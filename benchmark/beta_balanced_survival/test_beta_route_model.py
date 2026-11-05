"""
Fixture for the balanced Beta_Survival measurement (#163, README.md).

  0. Binding: the VBA procedures beta_route_model.py ports hash to the values
     results.json records. A change to any of them fails this fixture until
     the study is re-measured (measure_balanced_survival.py) and the README
     updated.
  1. Semantics and fidelity: recomputed, they equal results.json. The model
     reproduces every committed Excel observation it evaluates within
     FIDELITY_LIMIT_ULPS, and all of them with 64-bit intermediates only.
  2. Observations: every committed balanced-domain Beta_Survival and
     Beta_Cumulative error, recomputed against the binary64-argument
     references, equals results.json.
  3. The #163 point: the model returns the committed observation of
     Beta_Survival(0.82, 2.7, 3.4) bit for bit; its error exceeds 5E-15; with
     PROB_LogBeta exact, or with the Loader prefactor, it does not.
  4. Each sweep's recorded worst point for each function and variant
     reproduces. The full sweeps take about two minutes and are re-run by
     `python3 measure_balanced_survival.py --check`, not here.

Run: python3 test_beta_route_model.py   (exit 0 = pass, nonzero = fail)
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import beta_route_model as M                                      # noqa: E402
import measure_balanced_survival as S                              # noqa: E402

fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


with open(S.RESULTS, encoding="utf-8") as f:
    R = json.load(f)

# 0. Binding.
binding = M.source_binding()
for key in sorted(set(binding) | set(R["binding"])):
    check(binding.get(key) == R["binding"].get(key),
          f"binding: {key} changed since the measurement; re-run "
          "measure_balanced_survival.py and review README.md")

# 1. Semantics and fidelity.
sem = S.semantics()
check(sem == R["semantics"], f"semantics differ from results.json: {sem}")
ext = sem["64_bit_intermediates"]
check(ext["exact"] == ext["observations"] > 0,
      f"64-bit intermediates reproduce only {ext['exact']}/{ext['observations']} LogGamma observations")
fid = S.fidelity()
check(fid == R["fidelity"], "fidelity differs from results.json")
check(fid["max_ulps"] <= S.FIDELITY_LIMIT_ULPS,
      f"model differs from Excel by {fid['max_ulps']} ULPs (limit {S.FIDELITY_LIMIT_ULPS})")
for fn in ("Beta_Survival", "Beta_Cumulative", "Beta_Density", "LogGamma", "PROB_LogBeta"):
    check(fid["by_function"].get(fn, {}).get("observations", 0) > 0,
          f"no committed {fn} observation was modelled")

# 2. Observations.
obs = S.observations()
check(obs == R["observations"], "balanced-domain observation errors differ from results.json")

# 3. The #163 point.
row = [r for r in obs["rows"] if r["grid"].startswith("holdout_older/")
       and r["function"] == "Beta_Survival" and r["x"] == "0.82"]
check(len(row) == 1, "Beta_Survival(0.82, 2.7, 3.4) not found in holdout_older")
if row:
    check(row[0]["gate_error"] > S.THRESHOLD and row[0]["binary64_error"] > S.THRESHOLD,
          f"#163 point no longer exceeds 5E-15: {row[0]}")
    observed = [r for r in S.read_rows("holdout_older/holdout_older_grid.csv")
                if r["function"] == "Beta_Survival" and r["arg1"] == "0.82"][0]
    check(M.beta_survival(0.82, 2.7, 3.4) == float(S.parse_observed(observed["observed_vba"])),
          "model does not return the committed observation at the #163 point")
    ref = S.survival_reference(0.82, 2.7, 3.4)
    for name in ("exact_lb", "loader_prefactor"):
        e = S.rel(M.beta_survival(0.82, 2.7, 3.4, **S.VARIANTS[name]), ref)
        check(e < S.THRESHOLD, f"#163 point with {name}: {e:.3e} is not below 5E-15")

# 4. Recorded worst points.
for sweep, result in R["sweeps"].items():
    for fn, model, reference in (("Beta_Survival", M.beta_survival, S.survival_reference),
                                 ("Beta_Cumulative", M.beta_cumulative, S.cdf_reference)):
        for name, summary in result[fn].items():
            w = summary["worst"][0]
            p = (w["x"], w["a"], w["b"])
            e = S.sig(S.rel(model(*p, **S.VARIANTS[name]), reference(*p)))
            check(e == w["error"] == summary["max"],
                  f"{sweep}/{fn}/{name}: worst point {p} measures {e}, recorded {w['error']}")

if fails:
    print("FAIL: beta_balanced_survival")
    for msg in fails:
        print("  - " + msg)
    sys.exit(1)
print(f"PASS: beta_balanced_survival ({fid['observations']} observations reproduced, "
      f"{len(obs['rows'])} balanced-domain errors, {len(R['sweeps'])} sweeps' worst points)")
