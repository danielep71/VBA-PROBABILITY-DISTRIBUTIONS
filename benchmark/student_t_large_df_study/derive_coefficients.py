"""
Symbolic derivation of the Student-t large-df upper-tail expansion (#34).

    S(t; nu) = Q(t) + phi(t) * sum_{k>=1} g_k(t) / nu^k

Every g_k is DERIVED here from the exact density, as an exact rational
polynomial in t. Nothing is copied from memory or from a reference table: an
earlier attempt used recalled coefficients for g_2..g_4, and they were wrong
in a way that made higher orders fail to improve on the leading term. That is
why this file exists.

DERIVATION

  1. f(t; nu) = C(nu) (1 + t^2/nu)^(-(nu+1)/2), with
     log C(nu) = loggamma((nu+1)/2) - loggamma(nu/2) - log(nu pi)/2.
     Both loggamma terms are expanded by the Stirling asymptotic series
     (Bernoulli numbers from SymPy, not typed in).
  2. log( f / phi ) is expanded as a series in eps = 1/nu; exponentiating
     gives f/phi = 1 + sum c_k(t) eps^k with c_k polynomial in t.
  3. S(t) = int_t^inf f = Q(t) + sum eps^k int_t^inf phi(u) c_k(u) du.
     Each integral reduces by
         I_n(t) = int_t^inf u^n phi(u) du,  I_0 = Q, I_1 = phi,
         I_n = t^(n-1) phi(t) + (n-1) I_(n-2)
     to  g_k(t) phi(t) + q_k Q(t).
  4. Because int f = 1 for every nu, every q_k must vanish. The derivation
     asserts this; a non-zero q_k means the expansion is wrong.
  5. g_1 must equal the classical (t^3 + t)/4. This is a sanity assertion on
     the derivation, permitted by the preregistration; it is not a source.

OUTPUT

  coefficients.json - for each k: the exact rational polynomial as a list of
  (power, numerator, denominator) triples, plus its factored SymPy string, so a
  fixture can recompute and compare without trusting this file.

Run: python3 derive_coefficients.py [max_order]   (default 8)
"""
import json
import sys
import time

import sympy as sp

MAX_ORDER = int(sys.argv[1]) if len(sys.argv) > 1 else 8


def stirling_loggamma(x, order):
    """loggamma(x) ~ (x-1/2)log x - x + log(2pi)/2 + sum B_2k/(2k(2k-1) x^(2k-1))."""
    s = (x - sp.Rational(1, 2)) * sp.log(x) - x + sp.log(2 * sp.pi) / 2
    for k in range(1, order + 1):
        s += sp.bernoulli(2 * k) / (2 * k * (2 * k - 1) * x ** (2 * k - 1))
    return s


def derive(max_order):
    t, e, u, x = sp.symbols("t epsilon u x", positive=True)
    nu = 1 / e

    # Step 1-2: log(f/phi) as a series in eps, then exponentiate.
    L = (stirling_loggamma(x, max_order).subs(x, (nu + 1) / 2)
         - stirling_loggamma(x, max_order).subs(x, nu / 2)
         - sp.log(nu) / 2)
    K = -(nu + 1) / 2 * sp.log(1 + t ** 2 * e)
    logf = L + K - sp.log(sp.pi) / 2
    logphi = -t ** 2 / 2 - sp.log(2 * sp.pi) / 2
    ratio_log = sp.series(sp.expand(logf - logphi), e, 0, max_order + 1).removeO()
    R = sp.expand(sp.series(sp.exp(ratio_log), e, 0, max_order + 1).removeO())

    assert sp.simplify(R.coeff(e, 0) - 1) == 0, "f/phi must start at 1"

    # Step 3: integrate term by term via the I_n reduction.
    def I_n(n):
        if n == 0:
            return sp.Integer(0), sp.Integer(1)
        if n == 1:
            return sp.Integer(1), sp.Integer(0)
        p, q = I_n(n - 2)
        return sp.expand(t ** (n - 1) + (n - 1) * p), (n - 1) * q

    coeffs = {}
    for k in range(1, max_order + 1):
        ck = sp.Poly(sp.expand(R.coeff(e, k).subs(t, u)), u)
        gp, gq = sp.Integer(0), sp.Integer(0)
        for (n,), c in ck.terms():
            p, q = I_n(n)
            gp += c * p
            gq += c * q
        gp = sp.expand(gp)
        gq = sp.nsimplify(sp.expand(gq))
        # Step 4: the Q coefficient must vanish exactly.
        assert gq == 0, f"q_{k} = {gq} != 0: expansion is inconsistent with mass 1"
        # Every coefficient must be an exact rational.
        poly = sp.Poly(gp, t)
        for c in poly.all_coeffs():
            assert c.is_Rational, f"g_{k} has a non-rational coefficient {c}"
        coeffs[k] = poly

    # Step 5: classical leading term.
    assert sp.simplify(coeffs[1].as_expr() - (t ** 3 + t) / 4) == 0, \
        "g_1 does not match the classical (t^3+t)/4"
    return coeffs


def main():
    t0 = time.time()
    coeffs = derive(MAX_ORDER)
    t = sp.symbols("t", positive=True)
    out = {
        "expansion": "S(t; nu) = Q(t) + phi(t) * sum_k g_k(t) / nu^k",
        "max_order": MAX_ORDER,
        "derivation": ("Stirling series for the normalising constant, series in "
                       "eps = 1/nu, term-by-term integration via I_n reduction; "
                       "every Q coefficient asserted zero; g_1 asserted equal to "
                       "the classical (t^3+t)/4"),
        "sympy_version": sp.__version__,
        "runtime_s": None,
        "g": {},
    }
    for k, poly in coeffs.items():
        terms = []
        for (n,), c in poly.terms():
            c = sp.Rational(c)
            terms.append([int(n), int(c.p), int(c.q)])
        out["g"][str(k)] = {
            "degree": int(poly.degree()),
            "terms": sorted(terms, reverse=True),
            "factored": str(sp.factor(poly.as_expr())),
        }
        print(f"g_{k}(t) = {sp.factor(poly.as_expr())}")
    out["runtime_s"] = round(time.time() - t0, 3)
    with open("coefficients.json", "w", newline="\n") as f:
        json.dump(out, f, indent=2)
        f.write("\n")
    print(f"\nwrote coefficients.json (orders 1..{MAX_ORDER}) in {out['runtime_s']} s")


if __name__ == "__main__":
    main()
