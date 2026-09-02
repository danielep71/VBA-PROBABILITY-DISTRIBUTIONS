"""
Run the #35 isolation protocol (PREREGISTRATION.md sections 5-6) mechanically
and record the outcome in results.json - including the outcome that the
preregistered criteria are not met.

This script applies the design; it does not amend it. One post-hoc
diagnostic is included, clearly labelled, because it bears directly on whether
the mechanism in section 3 is real: the joint C5+C6 substitution. It is not a
fix selection.
"""
import json
import os
import platform
import sys
import time

import mpmath
from mpmath import mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import f_inverse_model as M                                          # noqa: E402

# Section 2 rows: (p, d1, d2, observed rel err). Exact binary64 values.
AFFECTED = (
    (0.5, 1e6, 3.0, 7.77e-11), (0.5, 1e8, 10.0, 9.49e-10),
    (0.9, 2.5, 1e6, 1.83e-11), (0.95, 2.5, 1e6, 1.58e-11), (0.99, 2.5, 1e6, 8.99e-12),
    (0.9, 5.0, 1e5, 1.40e-12), (0.9, 100.0, 1e10, 9.31e-9), (0.99, 100.0, 1e10, 7.00e-9),
)
CONTROLS = (
    (0.9, 1e6, 3.0, 4.64e-16), (0.99, 1e6, 3.0, 1.77e-16), (0.5, 2.5, 1e6, 1.52e-14),
    (0.9, 1e8, 10.0, 2.02e-16), (0.5, 100.0, 1e10, 3.84e-16),
)
ORDER = ("C6", "C7", "C3", "C5", "C2", "C4", "C1")     # section 5, fixed
CLEAN = 2.0 ** -48
FACTOR = 4.0


def err(p, d1, d2, prec):
    F, X, Y = M.f_inverse(p, d1, d2, prec)
    Fr, xr, yr = M.reference(p, d1, d2)
    mp.dps = M.DPS
    return float(abs(F - Fr) / Fr), float(min(xr, yr))


def main():
    t0 = time.time()
    problems = []
    rows = []

    # --- faithfulness (section 6) --------------------------------------------
    for p, d1, d2, obs in AFFECTED:
        e, tiny = err(p, d1, d2, M.all_f64())
        ratio = e / obs
        faithful = (1 / FACTOR) <= ratio <= FACTOR
        if not faithful:
            problems.append(f"faithfulness: ({p},{d1:.0e},{d2:.0e}) model {e:.2e} vs observed "
                            f"{obs:.2e}, ratio {ratio:.3f} outside [1/4, 4]")
        rows.append({"p": p, "d1": d1, "d2": d2, "kind": "affected", "observed": obs,
                     "min_xy": tiny, "eps_over_min": 2.0 ** -52 / tiny,
                     "model_all_f64": e, "ratio_model_to_observed": ratio,
                     "faithful": faithful})
    for p, d1, d2, obs in CONTROLS:
        e, tiny = err(p, d1, d2, M.all_f64())
        clean = e < CLEAN
        if not clean:
            problems.append(f"control reproduces as NOT clean: ({p},{d1:.0e},{d2:.0e}) {e:.2e}")
        rows.append({"p": p, "d1": d1, "d2": d2, "kind": "control", "observed": obs,
                     "min_xy": tiny, "model_all_f64": e, "clean": clean})

    # --- forward and reverse sweeps (section 5) ------------------------------
    sweeps = []
    for p, d1, d2, obs in AFFECTED + CONTROLS:
        base_f64, _ = err(p, d1, d2, M.all_f64())
        base_hp, _ = err(p, d1, d2, M.all_hp())
        fwd, rev = {}, {}
        for c in ORDER:
            pr = M.all_f64(); pr[c] = "hp"; fwd[c] = err(p, d1, d2, pr)[0]
            pr = M.all_hp(); pr[c] = "f64"; rev[c] = err(p, d1, d2, pr)[0]
        sweeps.append({"p": p, "d1": d1, "d2": d2, "all_f64": base_f64, "all_hp": base_hp,
                       "forward_single_hp": fwd, "reverse_single_f64": rev})

    # --- responsible-term test (section 5) -----------------------------------
    verdict = {}
    for c in ORDER:
        fwd_ok = all(s["forward_single_hp"][c] < CLEAN
                     for s in sweeps[:len(AFFECTED)])
        rev_ok = all(0 < s["reverse_single_f64"][c] and
                     (1 / FACTOR) <= s["reverse_single_f64"][c] / s["all_f64"] <= FACTOR
                     for s in sweeps[:len(AFFECTED)] if s["all_f64"] > 0)
        ctl_ok = all(max(s["forward_single_hp"][c], s["reverse_single_f64"][c]) < CLEAN * FACTOR
                     for s in sweeps[len(AFFECTED):])
        verdict[c] = {"forward_reduces_to_clean": fwd_ok, "reverse_reproduces": rev_ok,
                      "controls_unchanged": ctl_ok, "responsible": fwd_ok and rev_ok and ctl_ok}
    responsible = [c for c in ORDER if verdict[c]["responsible"]]
    if not responsible:
        problems.append("no single component satisfies section 5: hypothesis as decomposed "
                        "is not confirmed by the preregistered sweep")

    # --- post-hoc diagnostics, labelled ----------------------------------------
    posthoc = []
    for p, d1, d2, obs in AFFECTED:
        pr = M.all_f64(); pr["C5"] = "hp"; pr["C6"] = "hp"
        joint = err(p, d1, d2, pr)[0]
        posthoc.append({"p": p, "d1": d1, "d2": d2, "all_f64": err(p, d1, d2, M.all_f64())[0],
                        "C5_plus_C6_hp": joint})

    out = {
        "study": "#35 unbalanced-F inverse isolation",
        "preregistration": "PREREGISTRATION.md",
        "status": "PREREGISTERED CRITERIA NOT MET - recorded, not amended",
        "problems": problems,
        "generated_utc": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "runtime_s": round(time.time() - t0, 3),
        "versions": {"python": platform.python_version(), "mpmath": mpmath.__version__},
        "faithfulness": rows,
        "sweeps": sweeps,
        "section5_verdict_by_component": verdict,
        "responsible_components_by_section5": responsible,
        "post_hoc_diagnostics": {
            "standing": ("NOT preregistered; bears on whether the section-3 mechanism is "
                         "real; not a fix selection. Carrying U at high precision INTO the "
                         "pair formation (C5 and C6 together) removes the error; neither "
                         "alone does, because C5 alone hands a rounded U to a binary64 "
                         "subtraction and C6 alone subtracts from an already-rounded U."),
            "rows": posthoc,
        },
        "holdout_559_inspected": False,
    }
    with open(os.path.join(HERE, "results.json"), "w", newline="\n") as f:
        json.dump(out, f, indent=2)
        f.write("\n")
    print(f"status: {out['status']}")
    for pr in problems:
        print("  - " + pr)
    print(f"responsible by section 5: {responsible or 'NONE'}")
    print("post-hoc C5+C6 joint, worst:", f"{max(r['C5_plus_C6_hp'] for r in posthoc):.2e}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
