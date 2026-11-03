"""
Fixture for Amendment 2 of the #72 normal-tail study (#78): the conditioning
term of the §7 P2 and P3 bounds.

Amendment 2 restates P2 as error <= B + phi(z) * eta / F(z), where F is the
survival Q or the cumulative Phi at the exact standardised argument z, and eta
bounds |z~ - z|: 2^-52 |z| for Normal, 2^-52 (|z| + |ln x| / sigma) for
Lognormal. Wherever z != 0 that is the frozen B + kappa * delta with
kappa = |z| phi / F.

  1. The amendment is appended: PREREGISTRATION.md still begins with the bytes
     the holdout record binds, Amendment 2 follows them, and results.json
     lists it.
  2. Every recorded P2 row follows the amended rule: z, kappa and eta are
     recomputed; the bound is B + phi * eta / F, finite and at least B; where
     z != 0 it agrees with the frozen B + kappa * delta taken with |z|; the
     neglected second-order terms are below 1E-14 of the first-order term.
  3. Negative z: on the recorded negative-z survival rows the frozen signed
     expression z phi / Q gives a bound below B that falls as delta grows; the
     amended bound is at least B and rises with eta. On a sweep of Normal and
     Lognormal survival and cumulative rows over z in [-8, 8], the binary64
     model stays within the amended bound.
  4. Exact centre, x = 1 and mu = 0: the frozen delta is 0/0; the model's z is
     exactly 0, eta = 0, the bound is B and the model meets it. Normal at
     x = mu likewise.
  5. Near centre: x = 1 +/- a few ulps with mu = 0, and x = fl(exp(mu)) and
     its neighbours with mu != 0. The bound is finite, tends to B as x -> 1
     with mu = 0, approaches its limit sqrt(2/pi) 2^-52 |mu| / sigma as
     z -> 0, agrees with the frozen product wherever that is defined, and the
     model meets it.
  6. Seams: the split at z = 3, the 37.5 end of V, the last normal Q and the
     38.5 underflow cut-off, for survival and cumulative of both signs and for
     Lognormal at x = fl(exp(3)). The bound is finite and at least B; the model
     meets it wherever the reference is a normal Double (§4).

It never evaluates the holdout H and writes nothing.

Run: python3 test_conditioning.py   (exit 0 = pass, nonzero = fail)
"""
import csv
import json
import math
import os
import sys
from decimal import Decimal

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import kernel_model as km                                        # noqa: E402
import validate_kernel as vk                                     # noqa: E402

fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


B = 2.0 ** -50
U = mp.mpf(2) ** -52
DBL_MIN = 2.0 ** -1022


def reference(fn, z):
    return vk.q_ref(z) if fn.endswith("Survival") else vk.cdf_ref(z)


def frozen_kappa(fn, z, f):
    """§7 P2 as frozen: z phi / Q for survival (signed), |z| phi / Phi for cumulative."""
    with mp.workdps(60):
        lead = z if fn.endswith("Survival") else abs(z)
        return lead * vk.density(z) / f


def frozen_delta(fn, args):
    """§7 P2 as frozen; None where it is 0/0 (Lognormal at ln x = mu)."""
    with mp.workdps(60):
        if fn.startswith("Lognormal"):
            lx = mp.log(mp.mpf(args[0]))
            den = abs(lx - mp.mpf(args[1]))
            return None if den == 0 else U + U * abs(lx) / den
        return U


def amended(fn, args, f=None):
    """(z, F, eta, term, bound, model error, model z) for a P2 row; F defaults to the 60-digit oracle."""
    z = vk.exact_z(fn, args)
    f = reference(fn, z) if f is None else f
    eta = vk.z_error(fn, args, z)
    term = vk.standardisation_term(z, f, eta)
    value, model_z = km.public(fn, args, "candidate")
    return z, f, eta, term, B + term, vk.rel(value, f), model_z


def committed_references():
    """The committed reference of every P2 row, as validate_kernel.predictions reads it."""
    out = {}
    for grid in vk.GRIDS:
        with open(os.path.join(vk.BENCH, grid), encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                if row["function"] in vk.STANDARDISED:
                    args = [a for a in vk.args_of(row) if a is not None]
                    out[(grid, row["function"], tuple(args))] = Decimal(row["reference"].strip())
    return out


def close(a, b, tol):
    with mp.workdps(60):
        a, b = mp.mpf(a), mp.mpf(b)
        return abs(a - b) <= tol * max(abs(a), abs(b))


# 1. Appended amendment -----------------------------------------------------
with open(os.path.join(HERE, "results.json"), encoding="utf-8") as f:
    res = json.load(f)
with open(vk.PREREGISTRATION, "rb") as f:
    prereg = f.read()
frozen_bytes = res["holdout_exposure"]["design"]["bytes"]
heading = b"## Amendment 2"
check(heading not in prereg[:frozen_bytes], "Amendment 2 appears inside the frozen text the holdout record binds")
check(prereg.find(heading) > frozen_bytes, "Amendment 2 is not appended after the frozen text")
check(any(a.startswith("Amendment 2 (#78)") for a in res.get("amendments", [])),
      "results.json does not list Amendment 2")
check(any(r.get("reason", "").startswith("#78") and r.get("holdout_evaluated") is False
          for r in res.get("regenerations", [])), "no #78 regeneration recorded without the holdout")

# 2. Recorded P2 rows follow the amended rule -------------------------------
references = committed_references()
p2 = [r for r in res["predictions"]["rows"] if r["prediction"] == "P2"]
check(len(p2) == 27 and len(references) == 27, f"expected 27 P2 rows, recorded {len(p2)}, committed {len(references)}")
negative_survival, centre, second_order = [], [], 0.0
for r in p2:
    fn, args = r["function"], r["args"]
    where = f"{fn} {args}"
    committed = references.get((r["grid"], fn, tuple(args)))
    if committed is None:
        check(False, f"{where}: no committed reference")
        continue
    z, fz, eta, term, bound, _, _ = amended(fn, args, committed)
    fz = mp.mpf(str(committed))
    kappa = vk.condition(z, committed)
    check(r["z"] == float(z) and r["condition"] == kappa and r["z_error"] == float(eta),
          f"{where}: recorded z, kappa or eta differ from the amended rule")
    check(r["bound"] == bound, f"{where}: recorded bound {r['bound']} is not B + phi eta / F = {bound}")
    check(math.isfinite(r["bound"]) and r["bound"] >= B and kappa >= 0 and eta >= 0,
          f"{where}: bound {r['bound']}, kappa {kappa} or eta {eta} is not finite and nonnegative")
    check(r["within_bound"] is True and r["predicted_error"] <= r["bound"], f"{where}: outside its bound")
    delta = frozen_delta(fn, args)
    if z != 0:
        check(delta is not None and close(abs(frozen_kappa(fn, z, fz)) * delta, term, 1e-13),
              f"{where}: amended term differs from the frozen |kappa| * delta")
        second_order = max(second_order, float(abs(z) * eta + U))
    else:
        centre.append(r)
    if z < 0 and fn.endswith("Survival"):
        negative_survival.append(r)
        signed = B + float(frozen_kappa(fn, z, fz) * delta)
        signed_wider = B + float(frozen_kappa(fn, z, fz) * 2 * delta)
        check(signed_wider < signed < B < bound,
              f"{where}: expected the frozen signed bound below B and falling with delta, the amended one above B")
        check(r["predicted_error"] <= signed, f"{where}: the frozen signed bound would change this row's outcome")
check(second_order < 1e-14, f"neglected second-order terms reach {second_order:.2e} of the first-order term")
check(sorted(r["function"] + " " + str(r["args"]) for r in negative_survival)
      == sorted(["Normal_Survival [3.0, 5.0, 2.0]", "Normal_Survival [-2.1, 0.4, 1.7]",
                 "Lognormal_Survival [0.5, 0.0, 1.0]", "Lognormal_Survival [0.45, 0.2, 0.6]"]),
      f"negative-z survival rows changed: {[r['args'] for r in negative_survival]}")
check(sorted(r["function"] for r in centre) == ["Lognormal_Cumulative", "Lognormal_Survival"]
      and all(r["args"] == [1.0, 0.0, 1.0] and r["bound"] == B for r in centre),
      f"centre rows changed: {[(r['function'], r['args'], r['bound']) for r in centre]}")
for tag in ("P1", "P2", "P3", "P4"):
    summary = res["predictions"]["summary"][tag]
    check(summary.get("bounds_finite_nonnegative") is True and summary["outside_bound"] == [],
          f"{tag}: a bound is not finite and nonnegative, or a row is outside its bound")

# 3. Negative z sweep -------------------------------------------------------
swept, worst_ratio = 0, 0.0
for k in range(-64, 65):
    zt = k / 8
    for fn, args in (("Normal_Survival", [0.4 + 1.7 * zt, 0.4, 1.7]),
                     ("Normal_Cumulative", [0.4 + 1.7 * zt, 0.4, 1.7]),
                     ("Lognormal_Survival", [math.exp(0.2 + 0.6 * zt), 0.2, 0.6]),
                     ("Lognormal_Cumulative", [math.exp(0.2 + 0.6 * zt), 0.2, 0.6])):
        z, fz, eta, term, bound, err, _ = amended(fn, args)
        swept += 1
        worst_ratio = max(worst_ratio, err / bound)
        check(math.isfinite(bound) and bound >= B, f"{fn} {args}: bound {bound} not finite and >= B")
        check(err <= bound, f"{fn} {args} (z = {float(z):.4f}): model error {err:.3e} > bound {bound:.3e}")

# 4. Exact centre -----------------------------------------------------------
for fn in ("Lognormal_Survival", "Lognormal_Cumulative"):
    for sigma in (0.25, 1.0, 3.0):
        args = [1.0, 0.0, sigma]
        z, fz, eta, term, bound, err, model_z = amended(fn, args)
        check(frozen_delta(fn, args) is None, f"{fn} {args}: the frozen delta should be 0/0 here")
        check(z == 0 and eta == 0 and model_z == 0.0 and term == 0 and bound == B and err <= B,
              f"{fn} {args}: centre gives z {z}, eta {eta}, model z {model_z}, bound {bound}, error {err}")
for fn in ("Normal_Survival", "Normal_Cumulative"):
    z, fz, eta, term, bound, err, model_z = amended(fn, [0.4, 0.4, 1.7])
    check(z == 0 and model_z == 0.0 and bound == B and err <= B, f"{fn} at x = mu: bound {bound}, error {err}")

# 5. Near centre ------------------------------------------------------------
near = 0
for fn in ("Lognormal_Survival", "Lognormal_Cumulative"):
    for side in (1, -1):
        previous = math.inf
        for steps in (2 ** 20, 2 ** 10, 16, 2, 1):
            x = 1.0 + side * steps * (2.0 ** -52 if side > 0 else 2.0 ** -53)
            args = [x, 0.0, 0.25]
            z, fz, eta, term, bound, err, _ = amended(fn, args)
            near += 1
            check(math.isfinite(bound) and bound >= B and err <= bound, f"{fn} {args}: bound {bound}, error {err}")
            check(close(abs(frozen_kappa(fn, z, fz)) * frozen_delta(fn, args), term, 1e-13),
                  f"{fn} {args}: amended term differs from the frozen product")
            check(0 < term <= 4 * float(U * abs(mp.log(x))) / 0.25 and term < previous,
                  f"{fn} {args}: term {term} does not fall to 0 with ln x")
            previous = term
    for mu in (0.5, 1.0, -2.0, 10.0):
        x0 = math.exp(mu)
        for x in (math.nextafter(math.nextafter(x0, 0.0), 0.0), math.nextafter(x0, 0.0), x0,
                  math.nextafter(x0, math.inf), math.nextafter(math.nextafter(x0, math.inf), math.inf)):
            args = [x, mu, 0.25]
            z, fz, eta, term, bound, err, _ = amended(fn, args)
            near += 1
            delta = frozen_delta(fn, args)
            limit = mp.sqrt(2 / mp.pi) * U * abs(mu) / 0.25
            check(z != 0 and delta is not None and delta > 1e-3,
                  f"{fn} {args}: expected a near-centre row with a large frozen delta, got {delta}")
            check(close(abs(frozen_kappa(fn, z, fz)) * delta, term, 1e-13),
                  f"{fn} {args}: amended term differs from the frozen product")
            check(close(term, limit, 1e-9), f"{fn} {args}: term {term} is not near its limit {float(limit)}")
            check(math.isfinite(bound) and err <= bound, f"{fn} {args}: model error {err:.3e} > bound {bound:.3e}")

# 6. Seams ------------------------------------------------------------------
seams, characterised = 0, 0
for zs in (3.0, 37.5, vk.last_normal_q(), 38.5):
    for zt in vk.around(zs):
        for fn, args in (("Normal_Survival", [zt, 0.0, 1.0]), ("Normal_Cumulative", [-zt, 0.0, 1.0]),
                         ("Normal_Survival", [-zt, 0.0, 1.0]), ("Normal_Cumulative", [zt, 0.0, 1.0])):
            z, fz, eta, term, bound, err, _ = amended(fn, args)
            seams += 1
            check(math.isfinite(bound) and bound >= B, f"{fn} {args}: bound {bound} not finite and >= B")
            if fz >= DBL_MIN:
                check(err <= bound, f"{fn} {args}: model error {err:.3e} > bound {bound:.3e}")
            else:
                characterised += 1
for x in vk.around(math.exp(3.0)):
    for fn in ("Lognormal_Survival", "Lognormal_Cumulative"):
        z, fz, eta, term, bound, err, model_z = amended(fn, [x, 0.0, 1.0])
        seams += 1
        check(math.isfinite(bound) and bound >= B and err <= bound,
              f"{fn} [{x}, 0, 1] (model z {model_z!r}): error {err:.3e}, bound {bound:.3e}")
# Above the last normal Q (one probe) and around 38.5 (three), the survival of
# z and the cumulative of -z have a subnormal reference.
check(characterised == 8, f"expected 8 seam probes with a subnormal reference, found {characterised}")

if fails:
    print("FAIL: #78 conditioning of the #72 P2/P3 bounds")
    for msg in fails:
        print("  -", msg)
    sys.exit(1)
print(f"PASS: #78 conditioning: Amendment 2 appended after the {frozen_bytes} frozen bytes; 27 recorded P2 bounds "
      f"follow B + phi eta / F (4 negative-z survival rows, 2 centre rows at B); {swept} swept rows on z in [-8, 8] "
      f"within bound (worst error/bound {worst_ratio:.2f}); exact centre and {near} near-centre rows finite and "
      f"within bound; {seams} seam probes finite, {characterised} with a subnormal reference characterised only")
