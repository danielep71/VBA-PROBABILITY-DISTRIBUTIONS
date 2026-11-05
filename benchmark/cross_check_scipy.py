"""
Independent-oracle cross-check of the benchmark REFERENCE values against SciPy.

The reference values in the extreme Beta/F grids come from a 50-digit
continued-fraction incomplete beta (_ibeta.py). That is high precision but shares
algorithmic structure with the production kernel, so it is not a fully INDEPENDENT
oracle. SciPy (Cephes/Boost lineage) is independent, but loses reliability at very
large shape parameters - exactly the extreme regime the grids probe.

This script compares each reference row with SciPy and puts every row in exactly
one of three classes:
  - confirmed     : the script maps the function to an exact SciPy equivalent,
                    the point lies inside SciPy's reliable range and SciPy
                    returns a finite value. The relative difference is recorded;
                    the record gives the worst per function.
  - beyond range  : the function is mapped, but the incomplete-beta shape
                    parameter is above SCIPY_SHAPE_LIMIT (Beta and F only), or
                    SciPy returns no finite value. The record lists each such row
                    with its reason. These rows are NOT independently confirmed.
                    No committed tool records a 50-vs-80-digit self-consistency
                    check for them; f_envelope/README.md describes one for the
                    F envelope region only.
  - not mapped    : the script has no exact SciPy equivalent for the function,
                    so the row was never compared. The record names each such
                    function with its row count and the reason.
A function without a mapping is always "not mapped", whatever its arguments;
it is never counted as beyond SciPy's range.

Parameterisation of the mapped discrete families (library argument order on the
left, SciPy call on the right):
  Binomial(k, n, p)               binom(n, p) at k
  Poisson(k, mu)                  poisson(mu) at k
  Geometric(k, p)                 geom(p, loc=-1) at k; the library counts
                                  failures before the first success (support
                                  0, 1, ...), SciPy's geom counts trials
                                  (support 1, 2, ...), hence loc=-1
  NegativeBinomial(k, r, p)       nbinom(r, p) at k; both count failures k
                                  before the r-th success
  Hypergeometric(k, n, K, N)      hypergeom(M=N, n=K, N=n) at k; sample
                                  successes k, sample size n, population
                                  successes K, population size N
  DiscreteUniform(x, a, b)        randint(a, b + 1) at x; SciPy's upper bound is
                                  exclusive
Moments take the same arguments without the leading k. Every discrete inverse,
in the library and in SciPy, returns the least integer k with CDF(k) >= p;
survival is P(X > k) in both.

It writes cross_check_scipy.md so the README's "cross-checked against SciPy" claim
is a regenerable artifact, honestly scoped by SciPy's reliable range and by what
the script maps. The record lists every input grid with the SHA-256 of the bytes
it read, so it states which grids it covers, and the SciPy and NumPy versions
that produced it.

The classification and the record rendering do not import SciPy, so
test_cross_check_scipy.py runs them in CI without it; only evaluating SciPy
needs it.

USAGE, from benchmark/, with SciPy installed. The committed record was generated
with these three grids as they stood at commit ce9a274 (2026-11-05), after
the #159 re-reference of beta_f_unbalanced; its input list carries their
SHA-256 values. If a grid changes, re-run and commit the record with it.

    python3 cross_check_scipy.py --grids
        beta_f_unbalanced/beta_f_unbalanced_grid.csv
        beta_f_unbalanced/beta_f_inverse_grid.csv
        holdout/holdout_grid.csv
"""
import argparse
import csv
import hashlib
import math

import mpmath as mp

mp.mp.dps = 50

SCIPY_SHAPE_LIMIT = 1e7   # above this incomplete-beta shape param, SciPy is unreliable

CONFIRMED = "confirmed"
BEYOND = "beyond"
UNMAPPED = "unmapped"


# Each mapping: function name -> (number of leading numeric arguments, callable
# taking (scipy.stats, scipy.special, *args)). The argument order is the grid's
# arg1..argN, which is the library's own argument order.
MAPPINGS = {
    # Beta and F (incomplete-beta family)
    "Beta_Density": (3, lambda st, sp, x, a, b: st.beta.pdf(x, a, b)),
    "Beta_Cumulative": (3, lambda st, sp, x, a, b: sp.betainc(a, b, x)),
    "Beta_Survival": (3, lambda st, sp, x, a, b: st.beta.sf(x, a, b)),
    "Beta_InverseCumulative": (3, lambda st, sp, p, a, b: sp.betaincinv(a, b, p)),
    "F_Density": (3, lambda st, sp, x, d1, d2: st.f.pdf(x, d1, d2)),
    "F_Cumulative": (3, lambda st, sp, x, d1, d2: st.f.cdf(x, d1, d2)),
    "F_Survival": (3, lambda st, sp, x, d1, d2: st.f.sf(x, d1, d2)),
    "F_InverseCumulative": (3, lambda st, sp, p, d1, d2: st.f.ppf(p, d1, d2)),
    "PROB_LogBeta": (2, lambda st, sp, large, small: sp.betaln(large, small)),
    # Chi-square and normal family
    "ChiSquare_Density": (2, lambda st, sp, x, df: st.chi2.pdf(x, df)),
    "NormalStandard_InverseCumulative": (1, lambda st, sp, p: st.norm.ppf(p)),
    "NormalStandard_InverseSurvival": (1, lambda st, sp, p: st.norm.isf(p)),
    "Normal_InverseSurvival": (3, lambda st, sp, p, mean, sd:
                               st.norm.isf(p, loc=mean, scale=sd)),
    "Lognormal_InverseSurvival": (3, lambda st, sp, p, mulog, sdlog:
                                  st.lognorm.isf(p, sdlog, scale=math.exp(mulog))),
    # Binomial(k, n, p)
    "Binomial_PMF": (3, lambda st, sp, k, n, p: st.binom.pmf(k, n, p)),
    "Binomial_LogPMF": (3, lambda st, sp, k, n, p: st.binom.logpmf(k, n, p)),
    "Binomial_Cumulative": (3, lambda st, sp, k, n, p: st.binom.cdf(k, n, p)),
    "Binomial_Survival": (3, lambda st, sp, k, n, p: st.binom.sf(k, n, p)),
    "Binomial_InverseCumulative": (3, lambda st, sp, q, n, p: st.binom.ppf(q, n, p)),
    "Binomial_Mean": (2, lambda st, sp, n, p: st.binom.mean(n, p)),
    "Binomial_Variance": (2, lambda st, sp, n, p: st.binom.var(n, p)),
    "Binomial_StdDev": (2, lambda st, sp, n, p: st.binom.std(n, p)),
    # Poisson(k, mu)
    "Poisson_PMF": (2, lambda st, sp, k, mu: st.poisson.pmf(k, mu)),
    "Poisson_LogPMF": (2, lambda st, sp, k, mu: st.poisson.logpmf(k, mu)),
    "Poisson_Cumulative": (2, lambda st, sp, k, mu: st.poisson.cdf(k, mu)),
    "Poisson_Survival": (2, lambda st, sp, k, mu: st.poisson.sf(k, mu)),
    "Poisson_InverseCumulative": (2, lambda st, sp, q, mu: st.poisson.ppf(q, mu)),
    "Poisson_Mean": (1, lambda st, sp, mu: st.poisson.mean(mu)),
    "Poisson_Variance": (1, lambda st, sp, mu: st.poisson.var(mu)),
    "Poisson_StdDev": (1, lambda st, sp, mu: st.poisson.std(mu)),
    # Geometric(k failures, p): SciPy's geom counts trials, so loc=-1
    "Geometric_PMF": (2, lambda st, sp, k, p: st.geom.pmf(k, p, loc=-1)),
    "Geometric_LogPMF": (2, lambda st, sp, k, p: st.geom.logpmf(k, p, loc=-1)),
    "Geometric_Cumulative": (2, lambda st, sp, k, p: st.geom.cdf(k, p, loc=-1)),
    "Geometric_Survival": (2, lambda st, sp, k, p: st.geom.sf(k, p, loc=-1)),
    "Geometric_InverseCumulative": (2, lambda st, sp, q, p: st.geom.ppf(q, p, loc=-1)),
    "Geometric_Mean": (1, lambda st, sp, p: st.geom.mean(p, loc=-1)),
    "Geometric_Variance": (1, lambda st, sp, p: st.geom.var(p, loc=-1)),
    "Geometric_StdDev": (1, lambda st, sp, p: st.geom.std(p, loc=-1)),
    # NegativeBinomial(k failures, r successes, p)
    "NegativeBinomial_PMF": (3, lambda st, sp, k, r, p: st.nbinom.pmf(k, r, p)),
    "NegativeBinomial_LogPMF": (3, lambda st, sp, k, r, p: st.nbinom.logpmf(k, r, p)),
    "NegativeBinomial_Cumulative": (3, lambda st, sp, k, r, p: st.nbinom.cdf(k, r, p)),
    "NegativeBinomial_Survival": (3, lambda st, sp, k, r, p: st.nbinom.sf(k, r, p)),
    "NegativeBinomial_InverseCumulative": (3, lambda st, sp, q, r, p:
                                           st.nbinom.ppf(q, r, p)),
    "NegativeBinomial_Mean": (2, lambda st, sp, r, p: st.nbinom.mean(r, p)),
    "NegativeBinomial_Variance": (2, lambda st, sp, r, p: st.nbinom.var(r, p)),
    "NegativeBinomial_StdDev": (2, lambda st, sp, r, p: st.nbinom.std(r, p)),
    # Hypergeometric(k, n sample size, K population successes, N population size)
    # SciPy: hypergeom(M = population size, n = population successes, N = draws)
    "Hypergeometric_PMF": (4, lambda st, sp, k, n, K, N: st.hypergeom.pmf(k, N, K, n)),
    "Hypergeometric_LogPMF": (4, lambda st, sp, k, n, K, N:
                              st.hypergeom.logpmf(k, N, K, n)),
    "Hypergeometric_Cumulative": (4, lambda st, sp, k, n, K, N:
                                  st.hypergeom.cdf(k, N, K, n)),
    "Hypergeometric_Survival": (4, lambda st, sp, k, n, K, N: st.hypergeom.sf(k, N, K, n)),
    "Hypergeometric_InverseCumulative": (4, lambda st, sp, q, n, K, N:
                                         st.hypergeom.ppf(q, N, K, n)),
    "Hypergeometric_Mean": (3, lambda st, sp, n, K, N: st.hypergeom.mean(N, K, n)),
    "Hypergeometric_Variance": (3, lambda st, sp, n, K, N: st.hypergeom.var(N, K, n)),
    "Hypergeometric_StdDev": (3, lambda st, sp, n, K, N: st.hypergeom.std(N, K, n)),
    # DiscreteUniform(x, a, b), support a..b inclusive; SciPy's randint excludes high
    "DiscreteUniform_PMF": (3, lambda st, sp, x, a, b: st.randint.pmf(x, a, b + 1)),
    "DiscreteUniform_LogPMF": (3, lambda st, sp, x, a, b: st.randint.logpmf(x, a, b + 1)),
    "DiscreteUniform_Cumulative": (3, lambda st, sp, x, a, b: st.randint.cdf(x, a, b + 1)),
    "DiscreteUniform_Survival": (3, lambda st, sp, x, a, b: st.randint.sf(x, a, b + 1)),
    "DiscreteUniform_InverseCumulative": (3, lambda st, sp, q, a, b:
                                          st.randint.ppf(q, a, b + 1)),
    "DiscreteUniform_Mean": (2, lambda st, sp, a, b: st.randint.mean(a, b + 1)),
    "DiscreteUniform_Variance": (2, lambda st, sp, a, b: st.randint.var(a, b + 1)),
    "DiscreteUniform_StdDev": (2, lambda st, sp, a, b: st.randint.std(a, b + 1)),
}

# Functions deliberately left without a mapping, with the reason. Any other
# function without a mapping is reported as not mapped with DEFAULT_UNMAPPED_REASON.
UNMAPPED_REASONS = {
    "Normal_IntervalProbability":
        "SciPy has no interval-probability function; a difference of two SciPy "
        "CDFs is a different computation, not an exact equivalent",
    "Lognormal_ParametersFromMeanStdDev":
        "SciPy has no function that converts a mean and standard deviation to "
        "lognormal parameters",
}
DEFAULT_UNMAPPED_REASON = "no SciPy mapping in this script"

# Functions whose SciPy evaluation goes through the incomplete beta, so the
# shape-parameter range limit applies.
BETA_SHAPE_FUNCTIONS = frozenset({
    "Beta_Density", "Beta_Cumulative", "Beta_Survival", "Beta_InverseCumulative",
    "F_Density", "F_Cumulative", "F_Survival", "F_InverseCumulative",
})


def beta_shape_param(fn, args):
    """Incomplete-beta shape parameter behind a Beta or F function."""
    if fn.startswith("F_"):
        return max(args[1], args[2]) / 2
    return max(args[1], args[2])


def unmapped_reason(fn):
    return UNMAPPED_REASONS.get(fn, DEFAULT_UNMAPPED_REASON)


def _parse_args(fn, row, nargs, where):
    out = []
    for i in range(1, nargs + 1):
        text = (row.get(f"arg{i}") or "").strip()
        try:
            out.append(float(text))
        except ValueError:
            raise ValueError(f"{where}: {fn} needs a numeric arg{i}, got {text!r}") from None
    return out


def classify_row(fn, row, evaluate, where="row"):
    """Classify one reference row. Returns (class, detail).

    class is CONFIRMED (detail = relative difference as a float), BEYOND
    (detail = reason) or UNMAPPED (detail = reason). evaluate(fn, args) returns
    SciPy's value or raises; it is never called for an unmapped function or a
    point beyond the shape limit.
    """
    if fn not in MAPPINGS:
        return UNMAPPED, unmapped_reason(fn)
    ref_s = (row.get("reference") or "").strip()
    try:
        ref = mp.mpf(ref_s)
    except (TypeError, ValueError):
        raise ValueError(f"{where}: {fn} has no numeric reference ({ref_s!r})") from None
    nargs = MAPPINGS[fn][0]
    args = _parse_args(fn, row, nargs, where)
    if fn in BETA_SHAPE_FUNCTIONS:
        bp = beta_shape_param(fn, args)
        if bp > SCIPY_SHAPE_LIMIT:
            return BEYOND, (f"incomplete-beta shape parameter {bp:.6g} > "
                            f"{SCIPY_SHAPE_LIMIT:.0e}")
    try:
        sv = float(evaluate(fn, args))
    except Exception as exc:   # SciPy raised: it gives no value for this point
        return BEYOND, f"SciPy raised {type(exc).__name__}"
    if not math.isfinite(sv):
        return BEYOND, f"SciPy returned {sv}"
    diff = abs(mp.mpf(sv) - ref)
    rel = diff / abs(ref) if ref != 0 else diff
    return CONFIRMED, float(rel)


def cross_check(grids, evaluate):
    """Classify every row of every grid. Returns a result dictionary."""
    per_fn = {}
    beyond_rows = []
    total = 0
    worst = 0.0
    worst_at = ""
    for path in grids:
        with open(path, newline="", encoding="utf-8") as f:
            for line_no, r in enumerate(csv.DictReader(f), start=2):
                fn = r["function"]
                total += 1
                cls, detail = classify_row(fn, r, evaluate, where=f"{path}:{line_no}")
                d = per_fn.setdefault(fn, {CONFIRMED: 0, BEYOND: 0, UNMAPPED: 0,
                                           "worst": 0.0})
                d[cls] += 1
                if cls == CONFIRMED:
                    if detail > d["worst"]:
                        d["worst"] = detail
                    if detail > worst:
                        worst = detail
                        worst_at = (f"{fn} x={r['arg1']} a2={r['arg2'][:8]} "
                                    f"a3={r['arg3'][:8]}")
                elif cls == BEYOND:
                    args = ", ".join(r[k] for k in ("arg1", "arg2", "arg3", "arg4")
                                     if (r.get(k) or "").strip())
                    beyond_rows.append((fn, args, detail))
    counts = {c: sum(d[c] for d in per_fn.values()) for c in (CONFIRMED, BEYOND, UNMAPPED)}
    if sum(counts.values()) != total:
        raise AssertionError("a row was not classified exactly once")
    return {"per_fn": per_fn, "beyond_rows": beyond_rows, "counts": counts,
            "total": total, "worst": worst, "worst_at": worst_at}


def render(result, grid_digests, versions):
    """Render the Markdown record. grid_digests: [(path, sha256)];
    versions: [(name, version)]."""
    per_fn = result["per_fn"]
    counts = result["counts"]
    md = ["# SciPy independent cross-check of benchmark references", "",
          "Regenerated by `cross_check_scipy.py`. Each reference row is in exactly "
          "one of three classes. **Confirmed**: the script maps the function to an "
          "exact SciPy equivalent (an independent Cephes/Boost-lineage "
          "implementation), the point lies where SciPy is reliable, i.e. "
          f"incomplete-beta shape parameter <= {SCIPY_SHAPE_LIMIT:.0e} for Beta and "
          "F, and SciPy returns a finite value; the worst relative difference is "
          "shown. **Beyond SciPy range**: the function is mapped but the point is "
          "outside that range or SciPy returns no finite value. **Not mapped**: the "
          "script has no exact SciPy equivalent, so the row was never compared. "
          "Rows beyond SciPy's range and rows not mapped are not independently "
          "confirmed here, and this record cites no other evidence for them.", "",
          "| Function | Confirmed by SciPy | Worst rel. diff | Beyond SciPy range |",
          "|---|---:|---:|---:|"]
    for fn in sorted(per_fn):
        d = per_fn[fn]
        if d[CONFIRMED] == 0 and d[BEYOND] == 0:
            continue
        wtxt = f"{d['worst']:.1e}" if d[CONFIRMED] else "-"
        md.append(f"| {fn} | {d[CONFIRMED]} | {wtxt} | {d[BEYOND]} |")
    md += ["", "Rows beyond SciPy's reliable range, with the reason:", ""]
    if result["beyond_rows"]:
        md += ["| Function | Arguments | Reason |", "|---|---|---|"]
        for fn, args, reason in result["beyond_rows"]:
            md.append(f"| {fn} | {args} | {reason} |")
    else:
        md.append("None.")
    md += ["", "Functions not mapped to SciPy (rows never compared):", ""]
    unmapped = [fn for fn in sorted(per_fn) if per_fn[fn][UNMAPPED]]
    if unmapped:
        md += ["| Function | Rows | Reason |", "|---|---:|---|"]
        for fn in unmapped:
            md.append(f"| {fn} | {per_fn[fn][UNMAPPED]} | {unmapped_reason(fn)} |")
    else:
        md.append("None.")
    worst_at = result["worst_at"]
    md += ["", f"> {result['total']} reference row(s): {counts[CONFIRMED]} independently "
           f"confirmed by SciPy (worst relative difference {result['worst']:.1e}"
           + (f" at {worst_at}" if worst_at else "") + "); "
           f"{counts[BEYOND]} beyond SciPy's reliable range or not evaluated by "
           f"SciPy; {counts[UNMAPPED]} not mapped to SciPy. The "
           f"{counts[BEYOND] + counts[UNMAPPED]} row(s) in the last two classes are "
           "not independently confirmed.",
           "",
           "SciPy works in double precision, so agreement independently "
           "corroborates a reference only to the relative difference shown for its "
           "function, and this record does not say which side carries that "
           "difference. It does not, and cannot, certify the extreme-parameter "
           "points SciPy itself cannot compute, or functions this script does not "
           "map.",
           "",
           "Input grids (path as given, SHA-256 of the bytes read):",
           ""]
    for path, digest in grid_digests:
        md.append(f"- `{path}` `{digest}`")
    md += ["", "Generated with " + ", ".join(f"{name} {ver}" for name, ver in versions)
           + "."]
    return "\n".join(md) + "\n"


def scipy_evaluator():
    """Return (evaluate, versions) backed by SciPy. Imports SciPy lazily."""
    import numpy
    import scipy
    from scipy import special, stats

    def evaluate(fn, args):
        return MAPPINGS[fn][1](stats, special, *args)

    versions = [("SciPy", scipy.__version__), ("NumPy", numpy.__version__),
                ("mpmath", mp.__version__)]
    return evaluate, versions


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grids", nargs="+", required=True, help="reference grids to cross-check")
    ap.add_argument("--out", default="cross_check_scipy.md")
    a = ap.parse_args()

    evaluate, versions = scipy_evaluator()
    result = cross_check(a.grids, evaluate)
    digests = []
    for path in a.grids:
        with open(path, "rb") as g:
            digests.append((path, hashlib.sha256(g.read()).hexdigest()))
    with open(a.out, "w", encoding="utf-8", newline="\n") as f:
        f.write(render(result, digests, versions))
    c = result["counts"]
    print(f"{result['total']} rows: {c[CONFIRMED]} confirmed by SciPy "
          f"(worst {result['worst']:.1e}); {c[BEYOND]} beyond SciPy range; "
          f"{c[UNMAPPED]} not mapped")
    print(f"wrote {a.out}")


if __name__ == "__main__":
    main()
