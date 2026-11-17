"""Regenerate small independent Explorer fixtures using Python's standard library.

Normal uses erf/erfc; Exponential uses exp/expm1; Poisson uses its defining
finite mass sum. No production VBA, Excel or builder evaluation is used.
This educational fixture does not replace the library's accuracy contracts.
"""
import json
import math
from pathlib import Path


def references():
    rows = []
    for mean, sigma in [(0, 1), (2, 3)]:
        for z in [-2, 0, 1, 3]:
            x = mean + z * sigma
            rows.append(dict(family="Normal", mean=mean, sigma=sigma, x=x,
                             density=math.exp(-z*z/2)/(sigma*math.sqrt(2*math.pi)),
                             cdf=math.erfc(-z/math.sqrt(2))/2))
    for rate in [0.5, 2]:
        for x in [0, 0.5, 2, 4]:
            rows.append(dict(family="Exponential", rate=rate, x=x,
                             density=rate*math.exp(-rate*x), cdf=-math.expm1(-rate*x)))
    for mean in [1, 4, 20]:
        for k in [0, 1, 4, 20]:
            rows.append(dict(family="Poisson", poisson_mean=mean, x=k,
                             density=math.exp(-mean)*mean**k/math.factorial(k),
                             cdf=math.fsum(math.exp(-mean)*mean**j/math.factorial(j)
                                           for j in range(k+1))))
    return {"method": __doc__, "absolute_tolerance": 1e-12, "cases": rows}


if __name__ == "__main__":
    Path(__file__).with_name("explorer-references.json").write_text(
        json.dumps(references(), indent=2) + "\n", encoding="utf-8")
