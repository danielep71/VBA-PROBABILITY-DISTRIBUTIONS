"""
Analyze the LogGamma1p study (ICR-P1-01 prerequisite, issue #12).

  0. Argument integrity: did VBA evaluate the Double the grid intended?
  1. Kernel validation by regime, on the SCALED metric.
  2. Head-to-head against PROB_LogGamma(1# + X), the spelling being replaced.
  3. Seam continuity at PROB_LG1P_SERIES_MAX, one ulp either side.
  4. Subnormal-result degradation against its closed form, 2^-1075 / X.
  5. Recommended contract threshold under the 1-2-5 rule.

Metric note. The contract metric is the SCALED absolute error,

    Abs(observed - reference) / X

not the ordinary absolute or relative error. The scaled Gamma inverse computes
[LogProbability + LogGamma1p(Shape)] / Shape, so an absolute error in this
kernel is divided by Shape before it reaches the quantile: the scaled error IS
the relative error of the quantile it produces. Ordinary absolute error would
look flattering at small X and prove nothing about that caller.
"""
import argparse
import csv
import os
from collections import defaultdict
from decimal import Decimal, getcontext

getcontext().prec = 60

CONTRACT = Decimal("2.4E-16")          # provisional; frozen from holdout, not here
LOGGAMMA_CONTRACT = Decimal("6.1E-14")  # PROB_LogGamma's own published relative bound
SERIES_MAX = Decimal("0.25")
TWO_POW_M1075 = Decimal(2) ** -1075
ULP = Decimal(2) ** -52

# The scaled contract governs the series branch only. Points above the seam run
# PROB_LogGamma and are judged by ITS relative contract; points with a subnormal
# result are a documented representability limit and are judged by neither.
SCALED_CONTRACT_REGIMES = ("small",)


def parse(token):
    """Decode the exporter's hi;lo full-precision token."""
    token = (token or "").strip()
    if not token or token.upper() == "ERROR":
        return None
    return sum(Decimal(p) for p in token.split(";"))


def load(path):
    by_point = defaultdict(dict)
    order = []
    for r in csv.DictReader(open(path)):
        key = r["arg1"]
        if key not in by_point:
            order.append(key)
        by_point[key]["regime"] = r["regime"]
        by_point[key][r["quantity"]] = (parse(r["observed_vba"]), Decimal(r["reference"]))
    return order, by_point


def scaled(obs, ref, x):
    if obs is None or x == 0:
        return None
    return abs(obs - ref) / abs(x)


def fmt(v, width=10):
    return f"{float(v):{width}.2e}" if v is not None else f"{'--':>{width}}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grid", default="loggamma1p_grid.csv")
    ap.add_argument("--holdout", default=None,
                    help="independent holdout export; validates the threshold on "
                         "data that set none of it")
    a = ap.parse_args()
    order, pts = load(a.grid)

    missing = [x for x in order if pts[x].get("LogGamma1p", (None,))[0] is None]
    print(f"Loaded {len(order)} points.  Unfilled or ERROR rows: {len(missing)}")
    if len(missing) == len(order):
        print("  Grid is empty -- run Export_LogGamma1p in the workbook first.")
        return

    # ---- 0. argument integrity ------------------------------------------------
    print("\n0) ARGUMENT INTEGRITY (Val() round-trip of the grid literal)")
    bad = []
    for x in order:
        obs, ref = pts[x].get("EchoX", (None, None))
        if obs is None:
            continue
        # The question is only ever "did VBA evaluate the same Double", so
        # compare as Doubles. Decimal equality would be wrong twice over: the
        # hi;lo token is a 15+15 digit rendering, and at subnormal magnitudes
        # the residual X - hi is itself below the representable grid, so the
        # token cannot round-trip exactly there however it is written.
        if float(obs) != float(ref):
            bad.append((x, obs, ref))
    if bad:
        print(f"   {len(bad)} point(s) where VBA did NOT evaluate the intended Double:")
        for x, o, r in bad[:8]:
            print(f"     arg1={x:>24}  VBA={float(o):.17e}  intended={float(r):.17e}")
        print("   Every other number at these points is measuring the wrong argument.")
    else:
        print("   All arguments round-tripped exactly, including the subnormals.")

    # ---- 1. kernel validation by regime ---------------------------------------
    print("\n1) PROB_TryLogGamma1p -- scaled absolute error by regime")
    worst = defaultdict(lambda: (Decimal(0), None))
    for x in order:
        d = pts[x]
        obs, ref = d.get("LogGamma1p", (None, None))
        e = scaled(obs, ref, Decimal(x))
        if e is None:
            continue
        if e > worst[d["regime"]][0]:
            worst[d["regime"]] = (e, x)
    print(f"   {'regime':>18} {'worst scaled':>13}  at X")
    for reg in ("subnormal_result", "small", "series_seam", "lanczos_handover"):
        if reg not in worst:
            continue
        w, at = worst[reg]
        if reg in SCALED_CONTRACT_REGIMES:
            flag = "  OK" if w <= CONTRACT else "  OVER CONTRACT"
        else:
            flag = "  (not scaled-contracted)"
        print(f"   {reg:>18} {fmt(w, 13)}  {at}{flag}")
    overall = max([worst[r][0] for r in SCALED_CONTRACT_REGIMES if r in worst] or [Decimal(0)])
    print(f"\n   Small-series branch worst: {fmt(overall)}   (provisional contract {float(CONTRACT):.1e})")

    # The hand-over branch is PROB_LogGamma. It is reported on the SAME scaled
    # metric, not on relative error: Log(Gamma(1+X)) has zeros at X = 0 and
    # X = 1, so relative error there is unbounded by construction and says
    # nothing about the kernel. Absolute error stays near 1E-14 throughout.
    hw, hat = Decimal(0), None
    rw, rat = Decimal(0), None
    for x in order:
        if pts[x]["regime"] != "lanczos_handover":
            continue
        obs, ref = pts[x].get("LogGamma1p", (None, None))
        if obs is None:
            continue
        e = scaled(obs, ref, Decimal(x))
        if e is not None and e > hw:
            hw, hat = e, x
        if ref != 0:
            q = abs(obs - ref) / abs(ref)
            if q > rw:
                rw, rat = q, x
    if hat is not None:
        print(f"   Hand-over branch worst scaled: {fmt(hw)} at X={hat}")
        if rw > LOGGAMMA_CONTRACT:
            print(f"   NOTE: relative error there reaches {float(rw):.2e} at X={rat}, above")
            print(f"   PROB_LogGamma's published {float(LOGGAMMA_CONTRACT):.1e}. That is the zero of")
            print(f"   Log(Gamma) at X = 1, not a kernel defect -- but the published claim is")
            print(f"   stated as relative over Z in [1E-8, 1E+50] and does not hold there.")

    # ---- 2. head to head ------------------------------------------------------
    print("\n2) KERNEL vs PROB_LogGamma(1# + X) -- the defect being removed")
    print(f"   {'X':>24} {'kernel':>11} {'naive':>11} {'naive/kernel':>13}")
    for x in order:
        d = pts[x]
        if d["regime"] == "subnormal_result":
            continue
        ko, kr = d.get("LogGamma1p", (None, None))
        no, nr = d.get("LogGammaNaive", (None, None))
        ke, ne = scaled(ko, kr, Decimal(x)), scaled(no, nr, Decimal(x))
        if ke is None or ne is None:
            continue
        if ne <= CONTRACT and Decimal(x) > Decimal("1E-6"):
            continue                       # unremarkable, keep the table readable
        ratio = (ne / ke) if ke > 0 else None
        print(f"   {x:>24} {fmt(ke, 11)} {fmt(ne, 11)} {fmt(ratio, 13)}")

    # ---- 3. seam continuity ---------------------------------------------------
    print(f"\n3) SEAM CONTINUITY at PROB_LG1P_SERIES_MAX = {float(SERIES_MAX)}")
    seam = sorted((x for x in order
                   if Decimal("0.24") < Decimal(x) < Decimal("0.27")),
                  key=lambda v: Decimal(v))
    prev = None
    for x in seam:
        obs, ref = pts[x].get("LogGamma1p", (None, None))
        if obs is None:
            continue
        rel = abs(obs - ref) / abs(ref) if ref != 0 else None
        jump = abs(obs - prev) / abs(obs) if prev is not None and obs != 0 else None
        print(f"   X={x:>22}  rel={fmt(rel, 10)}  step from previous={fmt(jump, 10)}")
        prev = obs
    print("   A step of order 1E-14 here is PROB_LogGamma's own 6.1E-14 contract,")
    print("   not a defect in the series; it shrinks only if that kernel improves.")

    # ---- 4. subnormal-result limitation ---------------------------------------
    print("\n4) SUBNORMAL-RESULT LIMITATION -- measured vs half-ulp bound 2^-1075 / X")
    print(f"   {'X':>24} {'measured':>11} {'bound':>11} {'under':>6}")
    for x in order:
        if pts[x]["regime"] != "subnormal_result":
            continue
        obs, ref = pts[x].get("LogGamma1p", (None, None))
        e = scaled(obs, ref, Decimal(x))
        if e is None:
            continue
        b = TWO_POW_M1075 / Decimal(x)
        print(f"   {x:>24} {fmt(e, 11)} {fmt(b, 11)} {str(e <= b):>6}")
    print("   EulerGamma * X is itself subnormal here, so the returned Double has no")
    print("   grid point near the answer. This is a binary64 limit of the OUTPUT and")
    print("   belongs in numerical_limitations.csv, not in a contract. The closed form")
    print("   is a half-ulp UPPER BOUND, not a prediction: where the quantization lands")
    print("   favourably the measured error is smaller. What matters is that no point")
    print("   exceeds it -- an excess would mean the series, not the grid, is at fault.")

    # ---- 5. threshold recommendation ------------------------------------------
    print("\n5) CONTRACT RECOMMENDATION (1-2-5 rule, main grid only)")
    if overall > 0:
        exp = overall.adjusted()
        for mant in (1, 2, 5, 10):
            cand = Decimal(mant) * (Decimal(10) ** exp)
            if cand >= overall * 2:
                print(f"   LogGamma1p.small.scaled_abs  <=  {float(cand):.0e}"
                      f"   (measured {float(overall):.2e}, headroom {float(cand / overall):.1f}x)")
                break
    print("   Scope to X >= 3.855E-308 (PROB_MIN_NORMAL / EulerGamma); below that the")
    print("   result is itself subnormal and LogGamma1p.subnormal_result applies.")
    print("   Freeze only after the independent holdout is populated.")


    # ---- 6 --------------------------------------------------------------
    if not a.holdout:
        return
    print("\n6) INDEPENDENT HOLDOUT -- freeze decision for "
          "LogGamma1p.small.scaled_abs")
    if not os.path.exists(a.holdout):
        print(f"   {a.holdout} not found.")
        return
    horder, hpts = load(a.holdout)
    filled = [z for z in horder if hpts[z].get("LogGamma1p", (None,))[0] is not None]
    print(f"   {a.holdout}: {len(filled)}/{len(horder)} points filled")
    if not filled:
        print("   Not exported yet -- run Export_LogGamma1p and pick the holdout.")
        return
    overlap = set(horder) & set(order)
    print(f"   points shared with the fitting set: {len(overlap)}"
          f"{'' if not overlap else '  *** NOT INDEPENDENT'}")
    bad = [z for z in horder
           if hpts[z].get("EchoX", (None, None))[0] is not None
           and float(hpts[z]["EchoX"][0]) != float(hpts[z]["EchoX"][1])]
    print(f"   argument integrity: {'all exact' if not bad else str(len(bad)) + ' MISMATCH'}")

    worstH = Decimal(0); atH = None; n = 0
    below = Decimal(0); n_below = 0
    for z in horder:
        obs, ref = hpts[z].get("LogGamma1p", (None, None))
        e = scaled(obs, ref, Decimal(z))
        if e is None:
            continue
        n += 1
        if e > worstH:
            worstH, atH = e, z
        if Decimal(z) < Decimal("1e-300"):     # below the fitting set's floor
            n_below += 1
            below = max(below, e)
    thr = CONTRACT
    print(f"\n   {'contract':34s} {'threshold':>10} {'holdout worst':>14} {'pts':>4} "
          f"{'margin':>8}  verdict")
    margin = thr / worstH if worstH > 0 else None
    ok = worstH <= thr
    print(f"   {'LogGamma1p.small.scaled_abs':34s} {float(thr):10.0e} "
          f"{float(worstH):14.2e} {n:4d} "
          f"{(f'{float(margin):.1f}x' if margin else 'inf'):>8}  "
          f"{'PASS' if ok else 'FAIL'}   worst at X={atH}")
    print(f"\n   {n_below} point(s) lie below the fitting set's floor of 1E-300, "
          f"where the threshold was frozen")
    print(f"   without evidence. Worst there: {float(below):.2e}")
    if ok:
        print("\n   The threshold holds on data that set none of it, including the")
        print("   normal-result boundary the fitting grid never reached.")
    else:
        print("\n   The threshold is exceeded. Do NOT freeze: adjust it to the honest")
        print("   holdout-inclusive worst and record why.")


if __name__ == "__main__":
    main()