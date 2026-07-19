"""
Seam study for PROB_LogGammaDelta and the two-regime PROB_LogBeta crossover.

Produces reference values (mpmath, 50+ digits) for three measured quantities per
(Large, Small) point, so the VBA export can validate the delta kernel AND let us
choose the crossover from measured VBA error envelopes:

  LogGammaDelta   = LogGamma(Large+Small) - LogGamma(Large)   (delta kernel)
  LogBeta_ident   = Log(Beta(Large, Small))                   (via the identity)
  LogBeta_stable  = Log(Beta(Large, Small))                   (via LogGamma(Small)-delta)

The two LogBeta rows share the same reference; the point is to measure the VBA
error of each *route* at each ratio, so the crossover is chosen from data.
"""
import argparse, csv
import mpmath as mp
mp.mp.dps = 120

def ref_delta(large, small):
    return mp.loggamma(mp.mpf(large) + mp.mpf(small)) - mp.loggamma(mp.mpf(large))

def ref_logbeta(large, small):
    return mp.log(mp.beta(mp.mpf(large), mp.mpf(small)))

SMALLS = [mp.mpf(s) for s in ("0.25", "0.7", "1.3", "2.5", "5.75", "10.25")]

def build():
    rows = []
    def add(quantity, large, small, ref):
        rows.append({"quantity": quantity, "arg1": mp.nstr(large, 17),
                     "arg2": mp.nstr(small, 17), "reference": mp.nstr(ref, 30),
                     "observed_vba": ""})
    # A) seam region: fine ratios around the switch, Large = Small/ratio
    seam = [mp.mpf(r) for r in ("0.5","0.2","0.15","0.1","0.08","0.05","0.03","0.02","0.01","0.005")]
    # B) deep region: validate the delta far past the switch
    deep = [mp.mpf(10)**(-e) for e in (3,4,6,8,10,12,15,18)]
    for small in SMALLS:
        for ratio in seam + deep:
            large = small / ratio
            add("LogGammaDelta", large, small, ref_delta(large, small))
            add("LogBeta_ident", large, small, ref_logbeta(large, small))
            add("LogBeta_stable", large, small, ref_logbeta(large, small))
    # C) absolute-scale independence: fixed Large, ratio = Small/Large
    for large in [mp.mpf(x) for x in ("1e2","1e4","1e8","1e12","1e20","1e50")]:
        for small in SMALLS:
            add("LogGammaDelta", large, small, ref_delta(large, small))
            add("LogBeta_stable", large, small, ref_logbeta(large, small))
    return rows

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="delta_seam_grid.csv")
    a = ap.parse_args()
    rows = build()
    with open(a.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["quantity","arg1","arg2","reference","observed_vba"])
        w.writeheader(); w.writerows(rows)
    print(f"wrote {a.out}: {len(rows)} rows")

if __name__ == "__main__":
    main()
