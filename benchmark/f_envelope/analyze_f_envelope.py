"""
Locate the F accuracy envelope: the incomplete-beta shape parameter (= df/2) at
which F_Cumulative / F_Survival cross the 1.1E-10 accuracy contract.

Reports the degradation curve and the largest beta parameter that still meets the
contract, plus a conservative recommended envelope one measured step inside it.
"""
import argparse, csv
from collections import defaultdict
from decimal import Decimal, getcontext
getcontext().prec = 50
CONTRACT = Decimal("1.1E-10")

def parse(s):
    s = s.strip()
    return None if (not s or s.upper() == "ERROR") else sum(Decimal(p) for p in s.split(";"))

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--grid", default="f_envelope_grid.csv")
    a = ap.parse_args()
    rows = list(csv.DictReader(open(a.grid)))

    # group worst relative error by beta parameter (= max(d1,d2)/2)
    by_bp = defaultdict(lambda: {"worst": Decimal(0), "at": "", "err_rows": 0, "n": 0})
    for r in rows:
        o = parse(r["observed_vba"])
        if o is None:
            bp = max(float(r["arg2"]), float(r["arg3"])) / 2
            by_bp[bp]["err_rows"] += 1
            continue
        ref = Decimal(r["reference"])
        rel = abs(o - ref) / abs(ref) if ref != 0 else abs(o - ref)
        bp = max(float(r["arg2"]), float(r["arg3"])) / 2
        d = by_bp[bp]; d["n"] += 1
        if rel > d["worst"]:
            d["worst"] = rel
            d["at"] = f"{r['function']} x={r['arg1']} d1={r['arg2'][:8]} d2={r['arg3'][:10]} [{r['evidence_set'].split(':')[-1]}]"

    print("F accuracy vs incomplete-beta shape parameter (contract: rel <= 1.1E-10)\n")
    print(f"{'beta param (df/2)':>18}{'worst rel err':>15}{'meets contract':>16}   worst case")
    last_ok = None; first_fail = None
    for bp in sorted(by_bp):
        d = by_bp[bp]
        ok = d["worst"] <= CONTRACT and d["err_rows"] == 0
        if ok: last_ok = bp
        elif first_fail is None: first_fail = bp
        flag = "yes" if ok else ("ERROR rows" if d["err_rows"] else "NO")
        print(f"{bp:>18.2e}{float(d['worst']):>15.2e}{flag:>16}   {d['at']}")

    print()
    if last_ok is not None and first_fail is not None:
        print(f"Contract holds up to beta param {last_ok:.2e} (df ~ {2*last_ok:.2e});")
        print(f"first crossing at beta param {first_fail:.2e} (df ~ {2*first_fail:.2e}).")
        print(f"Recommended CONSERVATIVE envelope: reject df/2 > {last_ok:.0e}  (df > {2*last_ok:.0e}).")
    elif last_ok is not None:
        print(f"Contract holds across the whole sweep (up to beta param {last_ok:.2e}).")
    else:
        print("Contract fails across the whole sweep - envelope is below the smallest tested param.")

if __name__ == "__main__":
    main()
