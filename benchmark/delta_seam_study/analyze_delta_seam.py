"""
Analyze the delta seam study.

  1. Delta validation: relative error of PROB_LogGammaDelta across the grid.
  2. Crossover: LogBeta error via the identity vs via the stable difference, per
     ratio, to choose PROB_LOGBETA_STABLE_RATIO from measured VBA data.
"""
import argparse, csv
from collections import defaultdict
from decimal import Decimal, getcontext
getcontext().prec = 50
CLAIM = Decimal("5E-15")

def parse(s):
    s = s.strip()
    if not s or s.upper() == "ERROR":
        return None
    return sum(Decimal(p) for p in s.split(";"))

def rel(obs, ref):
    o = parse(obs)
    if o is None:
        return None
    r = Decimal(ref)
    return abs(o - r) / abs(r) if r != 0 else Decimal(0)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grid", default="delta_seam_grid.csv")
    a = ap.parse_args()
    rows = list(csv.DictReader(open(a.grid)))

    # 1. delta validation, split by whether the delta is actually used in production
    SWITCH = Decimal("0.1")   # PROB_LOGBETA_STABLE_RATIO (delta used when ratio < SWITCH)
    prod_worst = Decimal(0); prod_at = None
    full_worst = Decimal(0); full_at = None; derr = 0
    for r in rows:
        if r["quantity"] != "LogGammaDelta":
            continue
        e = rel(r["observed_vba"], r["reference"])
        if e is None:
            derr += 1; continue
        ratio = Decimal(r["arg2"]) / Decimal(r["arg1"])
        if e > full_worst:
            full_worst = e; full_at = (r["arg1"], r["arg2"])
        if ratio < SWITCH and e > prod_worst:
            prod_worst = e; prod_at = (r["arg1"], r["arg2"], float(ratio))
    print("1) PROB_LogGammaDelta validation")
    print(f"   production regime (ratio < {float(SWITCH)}): worst {float(prod_worst):.2e} "
          f"at (Large,Small)=({prod_at[0]}, {prod_at[1]})  meets 5E-15: {prod_worst <= CLAIM}")
    print(f"   full grid incl. balanced points: worst {float(full_worst):.2e} "
          f"at ({full_at[0]}, {full_at[1]}) -- points at ratio >= {float(SWITCH)} use the")
    print(f"   identity, not the delta, so a larger error there is expected and harmless.")
    print(f"   (ERROR rows: {derr})\n")

    # 2. crossover envelope by ratio
    ident = defaultdict(list); stable = defaultdict(list)
    for r in rows:
        if r["quantity"] not in ("LogBeta_ident", "LogBeta_stable"):
            continue
        large = Decimal(r["arg1"]); small = Decimal(r["arg2"])
        # canonical ratio bucket (round to 1 significant figure of the exponent-mantissa)
        raw = float(small / large)
        ratio = float(f'{raw:.1g}')   # 1-sig-fig bucket
        e = rel(r["observed_vba"], r["reference"])
        if e is None:
            continue
        (ident if r["quantity"] == "LogBeta_ident" else stable)[ratio].append(e)

    print("2) LogBeta crossover envelope (worst over Small at each ratio)")
    print(f"   {'ratio':>10} {'identity':>11} {'stable':>11} {'both<claim':>11} {'better':>9}")
    ratios = sorted(set(ident) | set(stable), reverse=True)
    safe_switch = None
    for ratio in ratios:
        iw = max(ident[ratio]) if ident.get(ratio) else None
        sw = max(stable[ratio]) if stable.get(ratio) else None
        iw_s = f"{float(iw):.2e}" if iw is not None else "—"
        sw_s = f"{float(sw):.2e}" if sw is not None else "—"
        both = (iw is not None and iw <= CLAIM) and (sw is not None and sw <= CLAIM)
        better = "stable" if (sw is not None and (iw is None or sw < iw)) else "identity"
        print(f"   {ratio:>10.0e} {iw_s:>11} {sw_s:>11} {str(both):>11} {better:>9}")

    # recommend: highest ratio where stable is safely within claim AND identity is
    # still fine (clean overlap), i.e. switch below there
    stable_ok = [r for r in ratios if stable.get(r) and max(stable[r]) <= CLAIM]
    ident_ok = [r for r in ratios if ident.get(r) and max(ident[r]) <= CLAIM]
    if stable_ok and ident_ok:
        overlap = [r for r in stable_ok if r in ident_ok]
        if overlap:
            print(f"\n   Clean overlap (both <= 5E-15) up to ratio {max(overlap):.0e}.")
            print(f"   Recommended PROB_LOGBETA_STABLE_RATIO: a value in that overlap,")
            print(f"   e.g. {max(overlap):.0e} (switch to stable below it).")

if __name__ == "__main__":
    main()
