"""
Proofs for the #307 repaired binary64 route model and its VBA block.

  1. The production procedures the model ports (PROB_StirlingError,
     PROB_Log1p, PROB_GammaSmallLogGamma1p) are the source text the port was
     written against: their SHA-256 is pinned here.
  2. The block b64_route_vba.py generates is the one the fit candidate 88ff6ba
     carried, the module ends with exactly that block (#314), and its constants are the
     model's: every log(k / 256) pair, log 2 and 2 pi.
  3. The deviance pair is accurate to 1E-16 absolute on 3,000 seeded random
     points (A from 1.01 to 1E7, X from 1E-3 A to 1E3 A).
  4. On the binary64 route map's 282 points, every tail the routed route
     would return (P from the series, Q = 1 - P where P <= 0.5, Q from the
     continued fraction, P = 1 - Q) meets its governing contract.
  5. Declines: X / A subnormal, and the off-side routine.

The fidelity of the VBA to this model is measured in Excel (FIT_DESIGN.md).

Run: python3 test_route_model.py   (exit 0 = pass)
"""
import csv
import hashlib
import math
import os
import random
import re
import sys
from fractions import Fraction

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import analyze_b64_fit as F  # noqa: E402
import analyze_binary64_map as M  # noqa: E402
import b64_route_vba as V  # noqa: E402
import route_model as RM  # noqa: E402

PINNED = {
    ("special", "PROB_StirlingError"): "80c310ce6519ddecf41d2508970cbc7ee20eed9e43073884772c34ec207ed2ad",
    ("core", "PROB_Log1p"): "80f61ebba4e7f85f16a85c32df3f88ae91dc35348e6a89c389fb543db8a256ec",
    ("special", "PROB_GammaSmallLogGamma1p"): "80ad4449b4d2939784ba251a5598b0e9b4b0c8411b3059a94f2a1b5663484cc2",
}


def main():
    fails = []

    def check(condition, message):
        if not condition:
            fails.append(message)

    # 1. Pinned production procedures.
    for (module, name), digest in PINNED.items():
        text = RM.BRM.procedure_text(module, name)
        check(hashlib.sha256(text.encode()).hexdigest() == digest, f"{name} changed since the port")

    # 2. The generated block and its constants.
    check(V.main(["--check"]) == 0, "the block is not the fit candidate's, or the module does not end with it")
    block = V.block()
    words = re.findall(r"Table(Hi|Lo)\((\d+)\) = PROB_DD_SP_FromWords\(&H(\w{8})&, &H(\w{8})&\)", block)
    check(len(words) == 2 * 129, f"{len(words)} table entries, not 258")
    for part, k, high, low in words:
        hi, lo = RM.LOG_TABLE[int(k) - 128]
        value = hi if part == "Hi" else lo
        check(V.words(value) == f"PROB_DD_SP_FromWords(&H{high}&, &H{low}&)", f"table {part}({k})")
    for value in (RM.LN2[0], RM.LN2[1], RM.TWO_PI):
        check(V.words(value) in block, f"constant {value!r} missing from the block")

    # 3. Deviance accuracy.
    rng = random.Random(307)
    worst = 0.0
    with mp.workdps(60):
        for _ in range(3000):
            a = 10 ** rng.uniform(math.log10(1.01), 7)
            x = a * 10 ** rng.uniform(-3, 3)
            dh, dl = RM.deviance(a, x)
            exact = mp.mpf(a) * mp.log(mp.mpf(a) / mp.mpf(x)) + mp.mpf(x) - mp.mpf(a)
            worst = max(worst, float(abs(mp.mpf(dh) + mp.mpf(dl) - exact)))
    check(worst <= 1e-16, f"deviance error {worst:.2e} exceeds 1E-16")

    # 4. The routed tails on the map.
    misses = []
    with open(M.GRID, encoding="utf-8") as f:
        for r in csv.DictReader(f):
            a, x = float(r["shape"]), float(r["x"])
            series, cf = F.modelled(a, x)
            p, q = F.tails(series and series[0], cf and cf[0])
            bound_p, bound_q = M.BOUNDS[M.regime(a, x)]
            for value, ref, bound in ((p, r["ref_p"], bound_p), (q, r["ref_q"], bound_q)):
                if value is None:
                    continue
                e = F.relative(value, Fraction(ref))
                if e != "underflow" and e > float(bound):
                    misses.append(f"A={a!r} X={x!r}: {e:.2e}")
    check(not misses, f"routed tails miss their contract: {misses[:4]}")

    # 5. Declines.
    check(F.modelled(1.5, 1e-310) == (None, None), "a subnormal X / A is not declined")
    series, cf = F.modelled(2.0, 1.0)
    check(series is not None and cf is None, "the continued fraction answers below A + 1")
    series, cf = F.modelled(2.0, 3.0)
    check(series is None and cf is not None, "the series answers at A + 1")

    if fails:
        print("FAIL: #307 binary64 route model")
        for f in fails:
            print("  - " + f)
        return 1
    print(f"PASS: #307 binary64 route model: pinned sources, historical VBA block, deviance within "
          f"{worst:.1e}, routed tails within contract on the map")
    return 0


if __name__ == "__main__":
    sys.exit(main())
