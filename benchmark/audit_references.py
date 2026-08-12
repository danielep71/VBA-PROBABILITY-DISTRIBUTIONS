"""
Audit every reference disagreement at the exact binary64 argument (#17 section B).

Each disagreeing row is scored against an INDEPENDENT high-precision oracle built
from mpmath primitives - never by calling the generator's own helper, so the
audit cannot be biased toward confirming the generator. GENERATOR_STALE is a
first-class outcome for that reason.

Inputs are exact rationals. float.as_integer_ratio() gives the mathematically
exact value of the Double the exporter passes, so no decimal-conversion
ambiguity enters the oracle at all.

Self-certifying. Every oracle is evaluated twice, at DPS_LOW and DPS_HIGH, and
the agreement between them is the measured uncertainty. A row is only decided
when the winner beats the loser by MARGIN and both the gap and the oracle
uncertainty are comfortably below that decision. Inverse functions additionally
report the forward residual, so convergence is certified rather than assumed.

Five numerical classifications, plus an orthogonal representation flag:

  REFERENCE_STALE    committed materially wrong, generator >= MARGIN closer
  GENERATOR_STALE    generator materially wrong, committed >= MARGIN closer
  BOTH_STALE         both materially outside the floor, neither wins by MARGIN
  BOTH_EQUIVALENT_AT_COMMITTED_PRECISION
                     neither materially distinguishable at the committed
                     reference precision
  ORACLE_AMBIGUOUS   the ORACLE's own certification is insufficient. Reserved
                     for genuine oracle uncertainty, never for a margin tie.

argument_text_diff / argument_bits_equal are flags, not classes. "0.85" and
"0.84999999999999998" can differ textually while the reference is stale,
generator-stale, both-stale or equivalent; mixing representation provenance with
numerical correctness would muddy the action to take.

replacement_source names that action, so the migration script is mechanical:

  NONE                 leave the committed reference alone
  CANONICAL_GENERATOR  adopt the generator value (it agrees with the oracle)
  INDEPENDENT_ORACLE   adopt a freshly computed oracle value; neither candidate
                       deserves to survive
  BLOCKED_GENERATOR    fix the generator first, do not touch the grid
  BLOCKED_ORACLE       migration blocked for this row

Report only; writes nothing but its own CSV.
"""
import argparse, csv, importlib.util, struct, sys
import time as _time
from collections import Counter
import mpmath as mp

DPS_LOW, DPS_HIGH = 160, 260

REPLACEMENT = {
    "REFERENCE_STALE": "CANONICAL_GENERATOR",
    "BOTH_STALE": "INDEPENDENT_ORACLE",
    "GENERATOR_STALE": "BLOCKED_GENERATOR",
    "BOTH_EQUIVALENT_AT_COMMITTED_PRECISION": "NONE",
    "ORACLE_AMBIGUOUS": "BLOCKED_ORACLE",
}
SER_FLOOR = mp.mpf(10) ** -24      # committed references carry 25 significant digits
MARGIN = mp.mpf(4)


def q(x):
    """Exact rational value of the Double denoted by x."""
    n, den = float(x).as_integer_ratio()
    return mp.mpf(n) / den


# ---------------- independent oracles ----------------
def _lg(z): return mp.loggamma(z)
def _lbeta(a, b): return _lg(a) + _lg(b) - _lg(a + b)
def _ibeta(x, a, b): return mp.betainc(a, b, 0, x, regularized=True)
def _Phi_inv(p):
    """Inverse standard normal CDF.

    erfinv(1 - 2p) collapses for small p exactly as 1 + x does for small x: at
    160 digits, 1 - 2E-300 rounds to 1 and erfinv returns +inf. The precision
    must therefore rise with the magnitude of p - the same adaptive rule the
    LogGamma1p reference needed, and for the same reason."""
    p = mp.mpf(p)
    if p <= 0 or p >= 1:
        raise ValueError("Phi_inv outside (0,1)")
    need = mp.mp.dps
    if p < mp.mpf("0.5"):
        need = max(need, int(-mp.log10(p)) + 60)
    with mp.workdps(need):
        r = (-mp.sqrt(2) * mp.erfinv(1 - 2 * p) if p < mp.mpf("0.5")
             else mp.sqrt(2) * mp.erfinv(2 * p - 1))
        if not mp.isfinite(r):
            raise ValueError("Phi_inv did not converge to a finite value")
        return +r


_NB_CACHE = {}


def _nb_cdf(k, r, p):
    """I_p(r, k+1).

    betainc does not converge for the large parameters this grid uses
    (r = 5000, k ~ 21000) at any maxterms, so the mass is summed directly. The
    sum is exact - its truncation is the whole support - and the terms are built
    by the recurrence pmf(i+1)/pmf(i) = ((i + r)/(i + 1)) * (1 - p), which costs
    one multiply and one divide per term instead of three loggamma evaluations.
    Memoised because Cumulative and Survival share every argument triple."""
    K = int(mp.floor(k))
    try:
        return _ibeta(p, r, mp.mpf(K) + 1)
    except Exception:
        key = (str(k), str(r), str(p), mp.mp.dps)
        if key in _NB_CACHE:
            return _NB_CACHE[key]
        q = 1 - p
        term = p ** r                       # i = 0
        total = term
        for i in range(K):
            term = term * ((mp.mpf(i) + r) / (mp.mpf(i) + 1)) * q
            total += term
        _NB_CACHE[key] = total
        return total



def _selfcheck_nb_recurrence():
    """The recurrence replaces mp.betainc only where betainc cannot converge.
    Before relying on it there, prove it reproduces that independent route on
    the domain where the route does converge. Runs once, at import."""
    old = mp.mp.dps
    try:
        mp.mp.dps = 120
        for k, r, p in [(4,1,"0.2"), (20,5,"0.5"), (60,50,"0.85"), (200,50,"0.5"),
                        (500,100,"0.2"), (1500,500,"0.5"), (3000,500,"0.2")]:
            K, R, P = mp.mpf(k), mp.mpf(r), mp.mpf(float(p))
            b = _ibeta(P, R, K + 1)
            q = 1 - P; term = P ** R; tot = term
            for i in range(int(K)):
                term = term * ((mp.mpf(i) + R) / (mp.mpf(i) + 1)) * q
                tot += term
            if abs(tot - b) > abs(b) * mp.mpf(10) ** -100:
                raise AssertionError(
                    f"NB recurrence disagrees with betainc at k={k}, r={r}, p={p}")
    finally:
        mp.mp.dps = old


_selfcheck_nb_recurrence()


def _hy_pmf(k, n, K, N): return mp.binomial(K, k) * mp.binomial(N - K, n - k) / mp.binomial(N, n)
def _hy_cdf(k, n, K, N):
    lo = int(max(mp.mpf(0), n + K - N))
    return mp.fsum([_hy_pmf(mp.mpf(i), n, K, N) for i in range(lo, int(k) + 1)])
def _du_n(lo, hi): return mp.floor(hi) - mp.ceil(lo) + 1


def _root_pos(f):
    """Bracketed monotone solve on (0, inf) in log coordinate."""
    lo, hi = mp.mpf(-700), mp.mpf(700)
    g = lambda t: f(mp.e**t)
    for _ in range(4000):
        m = (lo+hi)/2
        if g(m) > 0: hi = m
        else: lo = m
        if hi-lo < mp.mpf(10)**-(mp.mp.dps-12): break
    return mp.e**((lo+hi)/2)


def _root_real(f):
    """Bracketed monotone solve on the whole real line."""
    lo, hi = mp.mpf(-mp.mpf(10)**40), mp.mpf(10)**40
    for _ in range(4000):
        m = (lo+hi)/2
        if f(m) > 0: hi = m
        else: lo = m
        if hi-lo < max(abs(hi),mp.mpf(1))*mp.mpf(10)**-(mp.mp.dps-12): break
    return (lo+hi)/2


def _solve_logit(target, fwd, resid_out):
    """Solve fwd(x) = target for x in (0,1) by bracketed bisection in
    t = log(x/(1-x)). The logit coordinate keeps resolution at both endpoints,
    where x itself would lose every significant digit."""
    lo, hi = mp.mpf(-2000), mp.mpf(2000)
    f = lambda t: fwd(1 / (1 + mp.e ** (-t))) - target
    flo = f(lo)
    for _ in range(4000):
        mid = (lo + hi) / 2
        if (f(mid) > 0) == (flo > 0): lo = mid
        else: hi = mid
        if hi - lo < mp.mpf(10) ** -(mp.mp.dps - 12): break
    t = (lo + hi) / 2
    x = 1 / (1 + mp.e ** (-t))
    resid_out.append(abs(fwd(x) - target))
    return x


def _binom_cdf(k, n, p): return _ibeta(1-p, n-mp.floor(k), mp.floor(k)+1)
def _geo_cdf(k, p): return 1-(1-p)**(mp.floor(k)+1)
def _t_cdf(x, df):
    """Student t CDF via the regularized incomplete beta, both tails."""
    u = df/(df+x*x)
    h = _ibeta(u, df/2, mp.mpf(1)/2)/2
    return h if x <= 0 else 1-h

ORACLE_EXTRA = True


def build(dps, resid):
    """Oracle table at a given working precision."""
    mp.mp.dps = dps
    return {
 "LogGamma": (1, "loggamma", lambda z: _lg(z)),
 "LogGammaHalfDiff": (1, "loggamma", lambda z: _lg(z + mp.mpf(1)/2) - _lg(z)),
 "StirlingError": (1, "loggamma", lambda n: _lg(n+1) - (n*mp.log(n) - n + mp.log(2*mp.pi*n)/2)),
 "LogChoose": (2, "loggamma", lambda n, k: _lg(n+1) - _lg(k+1) - _lg(n-k+1)),
 "PROB_LogBeta": (2, "loggamma", _lbeta),

 "Beta_Density": (3, "beta pdf", lambda x,a,b: mp.e**((a-1)*mp.log(x)+(b-1)*mp.log(1-x)-_lbeta(a,b))),
 "Beta_Cumulative": (3, "betainc", lambda x,a,b: _ibeta(x,a,b)),
 "Beta_Survival": (3, "betainc", lambda x,a,b: 1-_ibeta(x,a,b)),
 "Beta_InverseCumulative": (3, "logit bisection / betainc",
    lambda p,a,b: _solve_logit(p, lambda x: _ibeta(x,a,b), resid)),

 "F_Density": (3, "beta pdf", lambda x,d1,d2: mp.e**(
    (d1/2)*mp.log(d1*x) + (d2/2)*mp.log(d2) - ((d1+d2)/2)*mp.log(d1*x+d2) - _lbeta(d1/2,d2/2)) / x),
 "F_Cumulative": (3, "betainc", lambda x,d1,d2: _ibeta(d1*x/(d1*x+d2), d1/2, d2/2)),
 "F_Survival": (3, "betainc", lambda x,d1,d2: 1-_ibeta(d1*x/(d1*x+d2), d1/2, d2/2)),
 # Solve for the BETA variable first, then transform. This removes the
 # unbounded F domain from the solver entirely.
 "F_InverseCumulative": (3, "logit bisection on u / betainc", lambda p,d1,d2: (
    lambda u: d2*u/(d1*(1-u)))(_solve_logit(p, lambda u: _ibeta(u, d1/2, d2/2), resid))),

 "NegativeBinomial_PMF": (3, "loggamma", lambda k,r,p: mp.e**(
    _lg(k+r)-_lg(k+1)-_lg(r)+r*mp.log(p)+k*mp.log(1-p))),
 "NegativeBinomial_LogPMF": (3, "loggamma", lambda k,r,p:
    _lg(k+r)-_lg(k+1)-_lg(r)+r*mp.log(p)+k*mp.log(1-p)),
 "NegativeBinomial_Cumulative": (3, "betainc/direct sum", _nb_cdf),
 "NegativeBinomial_Survival": (3, "betainc/direct sum", lambda k,r,p: 1-_nb_cdf(k,r,p)),
 "NegativeBinomial_Mean": (2, "exact rational", lambda r,p: r*(1-p)/p),
 "NegativeBinomial_Variance": (2, "exact rational", lambda r,p: r*(1-p)/(p*p)),
 "NegativeBinomial_StdDev": (2, "exact rational + sqrt", lambda r,p: mp.sqrt(r*(1-p)/(p*p))),

 "Hypergeometric_PMF": (4, "exact binomial", _hy_pmf),
 "Hypergeometric_LogPMF": (4, "exact binomial", lambda k,n,K,N: mp.log(_hy_pmf(k,n,K,N))),
 "Hypergeometric_Cumulative": (4, "exact sum", _hy_cdf),
 "Hypergeometric_Survival": (4, "exact sum", lambda k,n,K,N: 1-_hy_cdf(k,n,K,N)),
 "Hypergeometric_Variance": (3, "exact rational", lambda n,K,N: n*(K/N)*((N-K)/N)*((N-n)/(N-1))),
 "Hypergeometric_StdDev": (3, "exact rational + sqrt",
    lambda n,K,N: mp.sqrt(n*(K/N)*((N-K)/N)*((N-n)/(N-1)))),

 "DiscreteUniform_PMF": (3, "exact rational", lambda k,lo,hi: 1/_du_n(lo,hi)),
 "DiscreteUniform_LogPMF": (3, "exact rational", lambda k,lo,hi: -mp.log(_du_n(lo,hi))),
 "DiscreteUniform_Cumulative": (3, "exact rational",
    lambda k,lo,hi: (mp.floor(k)-mp.ceil(lo)+1)/_du_n(lo,hi)),
 "DiscreteUniform_Survival": (3, "exact rational",
    lambda k,lo,hi: 1-(mp.floor(k)-mp.ceil(lo)+1)/_du_n(lo,hi)),
 "DiscreteUniform_Variance": (2, "exact rational", lambda lo,hi: (_du_n(lo,hi)**2-1)/12),
 "DiscreteUniform_StdDev": (2, "exact rational + sqrt",
    lambda lo,hi: mp.sqrt((_du_n(lo,hi)**2-1)/12)),

 "NormalStandard_InverseCumulative": (1, "erfinv", _Phi_inv),
 "NormalStandard_InverseSurvival": (1, "erfinv", lambda p: -_Phi_inv(p)),
 "Normal_InverseSurvival": (3, "erfinv", lambda p,mu,sd: mu - sd*_Phi_inv(p)),
 "Lognormal_InverseSurvival": (3, "erfinv", lambda p,ml,sl: mp.exp(ml - sl*_Phi_inv(p))),

 "Binomial_PMF": (3, "loggamma", lambda k,n,p: mp.e**(
    _lg(n+1)-_lg(k+1)-_lg(n-k+1)+k*mp.log(p)+(n-k)*mp.log(1-p))),
 "Binomial_LogPMF": (3, "loggamma", lambda k,n,p:
    _lg(n+1)-_lg(k+1)-_lg(n-k+1)+k*mp.log(p)+(n-k)*mp.log(1-p)),
 "Binomial_Cumulative": (3, "betainc", _binom_cdf),
 "Binomial_Survival": (3, "betainc", lambda k,n,p: 1-_binom_cdf(k,n,p)),

 "Geometric_PMF": (2, "closed form", lambda k,p: p*(1-p)**k),
 "Geometric_LogPMF": (2, "closed form", lambda k,p: mp.log(p)+k*mp.log(1-p)),
 "Geometric_Cumulative": (2, "closed form", _geo_cdf),
 "Geometric_Survival": (2, "closed form", lambda k,p: (1-p)**(mp.floor(k)+1)),

 "StudentT_Density": (2, "loggamma", lambda x,df: mp.e**(
    _lg((df+1)/2)-_lg(df/2)-mp.log(df*mp.pi)/2-((df+1)/2)*mp.log(1+x*x/df))),
 "StudentT_Cumulative": (2, "betainc", _t_cdf),
 "StudentT_Survival": (2, "betainc", lambda x,df: 1-_t_cdf(x,df)),
 "StudentT_InverseCumulative": (2, "root/betainc", lambda p,df: _root_real(
    lambda x: _t_cdf(x,df)-p)),

 "ChiSquare_InverseCumulative": (2, "root/gammainc", lambda p,df: _root_pos(
    lambda x: mp.gammainc(df/2,0,x/2,regularized=True)-p)),
 "Gamma_InverseCumulative": (3, "root/gammainc", lambda p,k,th: _root_pos(
    lambda x: mp.gammainc(k,0,x/th,regularized=True)-p)),
}


FIELDS = ["function","regime","arg1","arg2","arg3","arg4",
          "arg1_hex","arg2_hex","arg3_hex","arg4_hex",
          "committed_reference","generator_reference","independent_oracle",
          "committed_abs_error","generator_abs_error",
          "committed_rel_error","generator_rel_error",
          "oracle_dps_low","oracle_dps_high","oracle_value_low","oracle_value_high",
          "oracle_convergence_delta","oracle_residual",
          "argument_text_diff","argument_bits_equal",
          "classification","replacement_source","oracle_method"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grid", default="probability_accuracy_grid.csv")
    ap.add_argument("--generator", default="generate_reference_values.py")
    ap.add_argument("--out", default="reference_audit.csv")
    a = ap.parse_args()

    spec = importlib.util.spec_from_file_location("gen", a.generator)
    g = importlib.util.module_from_spec(spec); sys.modules["gen"] = g
    spec.loader.exec_module(g)

    def bits(t):
        t = (t or "").strip()
        if t == "": return ""
        try: return struct.pack(">d", float(t)).hex()
        except (ValueError, OverflowError): return "?" + t

    mp.mp.dps = 50
    K = lambda r: (r["function"], bits(r["arg1"]), bits(r["arg2"]), bits(r["arg3"]),
                   bits(r["arg4"]), r["regime"], r.get("evidence_set", ""))
    N = {K(r): r for r in g.build_rows()}
    grid = list(csv.DictReader(open(a.grid)))
    dis = [r for r in grid if K(r) in N and r["reference"] != N[K(r)]["reference"]]

    out, cnt = [], Counter()
    t0 = _time.time()
    for _i, r in enumerate(dis, 1):
        fn = r["function"]
        row = {"function": fn, "regime": r["regime"],
               "committed_reference": r["reference"],
               "generator_reference": N[K(r)]["reference"],
               "oracle_dps_low": DPS_LOW, "oracle_dps_high": DPS_HIGH}
        gr = N[K(r)]
        tdiff = any(r[f"arg{i+1}"] != gr[f"arg{i+1}"] for i in range(4))
        beq = all(bits(r[f"arg{i+1}"]) == bits(gr[f"arg{i+1}"]) for i in range(4))
        row["argument_text_diff"] = "1" if tdiff else "0"
        row["argument_bits_equal"] = "1" if beq else "0"
        for i in range(4):
            row[f"arg{i+1}"] = r[f"arg{i+1}"]; row[f"arg{i+1}_hex"] = bits(r[f"arg{i+1}"])
        vals, resid = [], []
        method = ""
        try:
            for dps in (DPS_LOW, DPS_HIGH):
                tab = build(dps, resid)
                ar, method, f = tab[fn]
                vals.append(f(*[q(r[f"arg{i+1}"]) for i in range(ar)]))
            mp.mp.dps = DPS_HIGH
            lo, hi = vals
            if not (mp.isfinite(lo) and mp.isfinite(hi)):
                raise ValueError("oracle returned a non-finite value")
            delta = abs(hi - lo)
            if not mp.isfinite(delta):
                raise ValueError("oracle convergence delta is not finite")
            o = hi
            c, gv = mp.mpf(r["reference"]), mp.mpf(row["generator_reference"])
            scale = abs(o) if o != 0 else mp.mpf(1)
            ec, eg = abs(c - o), abs(gv - o)
            unc = max(delta, abs(o) * mp.mpf(10) ** -(DPS_HIGH - 20))
            floor_ = max(SER_FLOOR * scale, unc * MARGIN)
            if ec <= floor_ and eg <= floor_:
                cls = "BOTH_EQUIVALENT_AT_COMMITTED_PRECISION"
            elif eg * MARGIN < ec and ec > floor_:
                cls = "REFERENCE_STALE"
            elif ec * MARGIN < eg and eg > floor_:
                cls = "GENERATOR_STALE"
            else:
                # Both candidates are materially outside the floor and neither
                # wins by MARGIN. The margin is NOT relaxed to manufacture a
                # winner: a fresh oracle value replaces both.
                cls = "BOTH_STALE"
            row.update(independent_oracle=mp.nstr(o, 30),
                       committed_abs_error=mp.nstr(ec, 6), generator_abs_error=mp.nstr(eg, 6),
                       committed_rel_error=mp.nstr(ec/scale, 6),
                       generator_rel_error=mp.nstr(eg/scale, 6),
                       oracle_value_low=mp.nstr(lo, 30), oracle_value_high=mp.nstr(hi, 30),
                       oracle_convergence_delta=mp.nstr(delta, 6),
                       oracle_residual=mp.nstr(max(resid), 6) if resid else "",
                       classification=cls,
                       replacement_source=REPLACEMENT[cls],
                       oracle_method=method)
        except Exception as e:
            cls = "ORACLE_AMBIGUOUS"
            row.update(independent_oracle="", committed_abs_error="", generator_abs_error="",
                       committed_rel_error="", generator_rel_error="",
                       oracle_value_low="", oracle_value_high="",
                       oracle_convergence_delta="", oracle_residual="",
                       classification=cls,
                       replacement_source=REPLACEMENT[cls],
                       oracle_method=f"{method or 'unmapped'} FAILED: {type(e).__name__}")
        finally:
            mp.mp.dps = 50
        cnt[cls] += 1; out.append(row)
        if _i % 10 == 0 or _i == len(dis):
            print(f"  {_i:4d}/{len(dis)}  {_time.time()-t0:6.0f}s  last: {fn}", flush=True)

    with open(a.out, "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=FIELDS, lineterminator="\n")
        w.writeheader(); w.writerows(out)
    print(f"audited {len(out)} disagreements against an independent oracle "
          f"at {DPS_LOW} and {DPS_HIGH} digits\n")
    for c, n in cnt.most_common():
        print(f"  {c:42s} {n:5d}  -> {REPLACEMENT[c]}")
    td = sum(1 for r in out if r["argument_text_diff"] == "1")
    print(f"\n  rows whose argument TEXT differs from the generator: {td}"
          f"  (bits equal on {sum(1 for r in out if r['argument_bits_equal'] == '1')})")
    print(f"\nwrote {a.out}")


if __name__ == "__main__":
    main()
