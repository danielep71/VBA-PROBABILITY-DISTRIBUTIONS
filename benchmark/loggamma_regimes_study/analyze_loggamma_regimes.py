"""
Analyze the PROB_LogGamma regime study.

  0. Argument integrity: did VBA evaluate the Double the grid intended?
  1. Accuracy by regime, on the metric that regime actually supports.
  2. Baseline diff, classified row by row (needs --baseline).
  3. Subfinding A: the subnormal reflection defect, before and after.
  4. Subfinding B: why a global relative contract cannot work.
  5. Contract recommendation for the three replacement regimes.

Metric. Absolute error in the logarithm is primary everywhere except `general`,
because downstream callers propagate it: the relative error of Exp(v) is
approximately the absolute error of v. The two exact zeros are absolute-only --
a relative error against zero is not a number.
"""
import argparse
import csv
from collections import defaultdict
from decimal import Decimal, getcontext

getcontext().prec = 60

SERIES_MAX = Decimal("0.25")
EPS = Decimal(2) ** -52
ABS_REGIMES = ("small_positive", "reflection", "near_zero", "exact_zero")


def parse(token):
    token = (token or "").strip()
    if not token or token.upper() == "ERROR":
        return None
    return sum(Decimal(p) for p in token.split(";"))


def load(path):
    pts, order = {}, []
    for r in csv.DictReader(open(path)):
        z = r["arg1"]
        if z not in pts:
            pts[z] = {"regime": r["regime"], "metric": r["metric"], "bits": r["bits"]}
            order.append(z)
        pts[z][r["quantity"]] = (parse(r["observed_vba"]), Decimal(r["reference"]))
    return order, pts


def fmt(v, w=10):
    return f"{float(v):{w}.2e}" if v is not None else f"{'--':>{w}}"


def abs_err(obs, ref):
    return None if obs is None else abs(obs - ref)


def rel_err(obs, ref):
    if obs is None or ref == 0:
        return None
    return abs(obs - ref) / abs(ref)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grid", default="loggamma_regimes_grid.csv")
    ap.add_argument("--baseline", default=None,
                    help="pre-Phase-1 export of the same grid")
    a = ap.parse_args()
    order, pts = load(a.grid)

    filled = [z for z in order if pts[z].get("LogGamma", (None,))[0] is not None]
    print(f"Loaded {len(order)} points; {len(order) - len(filled)} unfilled or ERROR")
    if not filled:
        print("  Grid is empty -- run Export_LogGammaRegimes in the workbook first.")
        return

    # ---- 0 --------------------------------------------------------------
    print("\n0) ARGUMENT INTEGRITY")
    bad = [z for z in order
           if pts[z].get("EchoZ", (None, None))[0] is not None
           and float(pts[z]["EchoZ"][0]) != float(pts[z]["EchoZ"][1])]
    print(f"   {'All arguments round-tripped exactly.' if not bad else str(len(bad)) + ' MISMATCH: ' + ', '.join(bad[:6])}")

    # ---- 1 --------------------------------------------------------------
    print("\n1) ACCURACY BY REGIME")
    print(f"   {'regime':>15} {'n':>3} {'worst |abs|':>12} {'at Z':>24} {'worst rel':>11}")
    worst = {}
    for z in order:
        d = pts[z]
        obs, ref = d.get("LogGamma", (None, None))
        if obs is None:
            continue
        ae, re_ = abs_err(obs, ref), rel_err(obs, ref)
        w = worst.setdefault(d["regime"], [Decimal(0), None, Decimal(0), 0])
        w[3] += 1
        if ae > w[0]:
            w[0], w[1] = ae, z
        if re_ is not None and re_ > w[2]:
            w[2] = re_
    for reg in ("small_positive", "reflection", "near_zero", "exact_zero", "general"):
        if reg not in worst:
            continue
        ae, at, re_, n = worst[reg]
        rs = fmt(re_, 11) if reg == "general" else f"{'n/a':>11}"
        print(f"   {reg:>15} {n:3d} {fmt(ae, 12)} {at:>24} {rs}")
    print("   'general' is the only regime where a relative contract is meaningful;")
    print("   elsewhere LogGamma passes through or approaches zero.")

    # ---- 2 --------------------------------------------------------------
    if a.baseline:
        border, bpts = load(a.baseline)
        print("\n2) BASELINE DIFF, row by row")
        buckets = defaultdict(list)
        for z in order:
            n_obs, ref = pts[z].get("LogGamma", (None, None))
            b_obs = bpts.get(z, {}).get("LogGamma", (None, None))[0]
            if n_obs is None or b_obs is None:
                continue
            reachable = Decimal(z) <= SERIES_MAX
            if n_obs == b_obs:
                buckets["unchanged"].append((z, reachable))
                continue
            oe, ne = abs_err(b_obs, ref), abs_err(n_obs, ref)
            tol = 2 * EPS * abs(ref)
            if ne < oe:
                k = "expected improvement"
            elif ne - oe <= tol:
                k = "expected neutral rounding"
            else:
                k = "UNEXPECTED"
            buckets[k].append((z, reachable, oe, ne))
        for k in ("expected improvement", "expected neutral rounding", "UNEXPECTED", "unchanged"):
            v = buckets.get(k, [])
            print(f"   {k:>27}: {len(v)}")
        # structural assertion: only Z <= 0.25 may move, and it is the only branch edited
        moved_out = [z for k in ("expected improvement", "expected neutral rounding", "UNEXPECTED")
                     for (z, reach, *_) in buckets.get(k, []) if not reach]
        print(f"\n   Points that moved but are OUTSIDE Z <= {float(SERIES_MAX)}: "
              f"{'none' if not moved_out else moved_out}")
        print("   Any such point is a defect: Phase 1 edits only the small-positive branch.")
        if buckets.get("UNEXPECTED"):
            print(f"\n   {'Z':>24} {'old |abs|':>11} {'new |abs|':>11}")
            for z, _, oe, ne in buckets["UNEXPECTED"]:
                print(f"   {z:>24} {fmt(oe, 11)} {fmt(ne, 11)}")
        top = sorted(buckets.get("expected improvement", []), key=lambda t: -(t[2] / t[3]) if t[3] else 0)[:8]
        if top:
            print(f"\n   Largest improvements")
            print(f"   {'Z':>24} {'old |abs|':>11} {'new |abs|':>11} {'factor':>10}")
            for z, _, oe, ne in top:
                fac = (oe / ne) if ne else None
                print(f"   {z:>24} {fmt(oe, 11)} {fmt(ne, 11)} {fmt(fac, 10)}")

    # ---- 3 --------------------------------------------------------------
    print("\n3) SUBFINDING A -- subnormal reflection defect")
    print(f"   {'Z':>24} {'bits':>5} {'|abs| in log':>13} {'relative':>11}")
    for z in order:
        if pts[z]["regime"] != "small_positive":
            continue
        if Decimal(z) >= Decimal("1e-307"):
            continue
        obs, ref = pts[z].get("LogGamma", (None, None))
        if obs is None:
            continue
        print(f"   {z:>24} {pts[z]['bits']:>5} {fmt(abs_err(obs, ref), 13)} {fmt(rel_err(obs, ref), 11)}")
    print("   Pre-Phase-1 the reflection route reaches ~4.6E-02 of absolute log error")
    print("   here, because PROB_PI * Z is subnormal before Sin() ever sees it.")

    # ---- 4 --------------------------------------------------------------
    print("\n4) SUBFINDING B -- why a global relative contract cannot work")
    print(f"   {'Z':>24} {'LogGamma(Z)':>16} {'|abs|':>11} {'relative':>11}")
    for z in order:
        if pts[z]["regime"] not in ("near_zero", "exact_zero"):
            continue
        obs, ref = pts[z].get("LogGamma", (None, None))
        if obs is None:
            continue
        r = rel_err(obs, ref)
        rs = fmt(r, 11) if r is not None else f"{'undefined':>11}"
        print(f"   {z:>24} {float(ref):16.8e} {fmt(abs_err(obs, ref), 11)} {rs}")
    print("   Absolute error is flat through the zeros; the relative column diverges")
    print("   because the denominator does, not because the kernel degrades.")

    # ---- 5 --------------------------------------------------------------
    print("\n5) CONTRACT RECOMMENDATION (1-2-5 rule, main grid only)")
    def rec(label, keys, kind):
        vals = [worst[k][0 if kind == "abs" else 2] for k in keys if k in worst]
        if not vals:
            return
        w = max(vals)
        if w == 0:
            print(f"   {label:>34}  exact")
            return
        exp = w.adjusted()
        # 1-2-5 must be allowed to cross into the next decade, or a worst case
        # just above half a decade silently yields no candidate at all.
        for m in (1, 2, 5, 10, 20, 50, 100):
            c = Decimal(m) * (Decimal(10) ** exp)
            if c >= w * 2:
                print(f"   {label:>34} <= {float(c):.0e}   (worst {float(w):.2e}, headroom {float(c / w):.1f}x)")
                return
        print(f"   {label:>34}   no candidate found (worst {float(w):.2e})")
    rec("LogGamma.small_positive.log_abs", ["small_positive"], "abs")
    print("     floor here is EPS * Abs(Log(Z)), irreducible: 1.65E-13 at the smallest")
    print("     positive subnormal, 5.1E-14 at 1E-100, 3.1E-16 at Z = 0.25. The absolute")
    print("     error is dominated by Log(Z), not by the series.")
    rec("LogGamma.near_zero.log_abs", ["near_zero", "exact_zero"], "abs")
    rec("LogGamma.general.output_rel", ["general"], "rel")
    print("   The reflection regime (0.25 < Z < 0.5) is unchanged by Phase 1 and is")
    print("   covered by LogGamma.small_positive.log_abs or its own row, as preferred.")
    print("   Freeze only after the independent holdout is populated.")


if __name__ == "__main__":
    main()
