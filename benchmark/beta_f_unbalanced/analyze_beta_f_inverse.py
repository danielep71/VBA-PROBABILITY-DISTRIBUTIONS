"""
Analyze the unbalanced Beta/F INVERSE study.

Two metrics per point (inverse solvers amplify normalization error differently
from the forward functions, so both matter):

  quantile error     abs and relative error in the returned x;
  forward residual    push x_VBA through the TRUE (mpmath) CDF and compare the
                      recovered probability to the target p:
                        abs residual        |I_{x_VBA}(a,b) - p|
                        tail-relative        |I_{x_VBA}(a,b) - p| / min(p, 1-p)

The tail-relative residual is the operationally meaningful one for steep
quantiles: a tiny x error can still be a large probability error, and vice versa.
"""
import argparse, csv
from collections import defaultdict
from decimal import Decimal, getcontext
import mpmath as mp
from _ibeta import ibeta, f_cdf
getcontext().prec = 50
mp.mp.dps = 50

def parse(s):
    s = s.strip()
    if not s or s.upper() == "ERROR":
        return None
    return sum(Decimal(p) for p in s.split(";"))

def true_cdf(fn, x, a2, a3):
    if fn == "Beta_InverseCumulative":
        return ibeta(x, a2, a3)
    return f_cdf(x, a2, a3)   # F_InverseCumulative

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grid", default="beta_f_inverse_grid.csv")
    a = ap.parse_args()
    rows = list(csv.DictReader(open(a.grid)))

    stats = defaultdict(lambda: {"qrel": [], "resid": [], "tailrel": [], "err": 0})
    for r in rows:
        fn = r["function"]
        xobs = parse(r["observed_vba"])
        if xobs is None:
            stats[fn]["err"] += 1
            continue
        p = mp.mpf(r["arg1"]); a2 = mp.mpf(r["arg2"]); a3 = mp.mpf(r["arg3"])
        xref = mp.mpf(r["reference"])
        xo = mp.mpf(str(xobs))
        # quantile relative error
        qrel = abs(xo - xref) / abs(xref) if xref != 0 else abs(xo - xref)
        # forward residual via the true CDF
        prec_p = true_cdf(fn, xo, a2, a3)
        resid = abs(prec_p - p)
        tail = resid / min(p, 1 - p)
        stats[fn]["qrel"].append((qrel, f"p={r['arg1']}, {r['arg2']}, {r['arg3']}"))
        stats[fn]["resid"].append(resid)
        stats[fn]["tailrel"].append((tail, f"p={r['arg1']}, {r['arg2']}, {r['arg3']}"))

    print("Unbalanced Beta/F inverse: quantile error and forward-probability residual\n")
    for fn in sorted(stats):
        s = stats[fn]
        if not s["qrel"]:
            print(f"{fn}: no observations ({s['err']} ERROR)"); continue
        wq, wq_at = max(s["qrel"], key=lambda t: t[0])
        wr = max(s["resid"])
        wt, wt_at = max(s["tailrel"], key=lambda t: t[0])
        print(f"{fn}  ({len(s['qrel'])} points, {s['err']} ERROR)")
        print(f"    worst quantile rel error : {float(wq):.2e}   at {wq_at}")
        print(f"    worst forward residual   : {float(wr):.2e}")
        print(f"    worst tail-rel residual  : {float(wt):.2e}   at {wt_at}")
        print(f"    -> suggested contract: quantile rel <= {2*10**__import__('math').ceil(__import__('math').log10(float(wq))):.0e}, "
              f"tail-rel residual <= {2*10**__import__('math').ceil(__import__('math').log10(float(wt))):.0e}")
    print("\nMark inverse contracts 'active' only after this passes on the final module;")
    print("until then they are PENDING. Freeze thresholds from the tail-relative residual,")
    print("the operationally meaningful metric for steep quantiles.")

if __name__ == "__main__":
    main()
