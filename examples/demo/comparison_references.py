"""Write the cases of the Comparison sheet (Excel vs VBA vs Python) with their SciPy values.

The cases are the 29 rows of benchmark/excel_comparison/excel_comparison_grid.csv, the study that
measures Excel's native functions and the library against the same references: a body point of
every family, and the deep tails, large shapes and missing functions where a difference is
expected. The sheet reuses them as they are: the label, the Excel formula ("-" where Excel has no
equivalent), the library call and the reference, which generate_excel_comparison.py computes with
mpmath at 50 digits. Nothing is re-referenced here.

What this script adds is the Python column: for each case, the SciPy call with the same inputs,
the same parameterization (Gamma and Weibull scale, Exponential rate as scale = 1/lambda, Lognormal
mu and sigma as scale = exp(mu) and s = sigma, the counts of the Negative binomial and the
Hypergeometric) and the same tail (sf for an upper tail, isf for an inverse upper tail), and its
value computed now with the installed SciPy. The values are stored, with the versions of Python,
NumPy, SciPy and mpmath, so the sheet works without Python in Excel; the code is written so it can
be pasted into an =PY() cell, where the same SciPy call runs in Microsoft's cloud.

SciPy is a third implementation, not a reference: the sheet grades it against the mpmath value as
it grades Excel and the library.

Run with SciPy installed: python examples/demo/comparison_references.py
"""
import csv
import json
import platform
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
GRID = ROOT / "benchmark" / "excel_comparison" / "excel_comparison_grid.csv"
OUTPUT = HERE / "comparison-cases.json"

# The SciPy call of each library call, with the same inputs, parameterization and tail.
PYTHON = {
    "K_STATS_NormalStandard_Cumulative(1.5)": "from scipy.stats import norm; norm.cdf(1.5)",
    "K_STATS_NormalStandard_Survival(8)": "from scipy.stats import norm; norm.sf(8)",
    "K_STATS_NormalStandard_Survival(15)": "from scipy.stats import norm; norm.sf(15)",
    "K_STATS_NormalStandard_InverseSurvival(1E-15)": "from scipy.stats import norm; norm.isf(1e-15)",
    "K_STATS_Lognormal_Cumulative(2,0,1)": "from scipy.stats import lognorm; import math; lognorm.cdf(2, s=1, scale=math.exp(0))",
    "K_STATS_Lognormal_Survival(2981,0,1)": "from scipy.stats import lognorm; import math; lognorm.sf(2981, s=1, scale=math.exp(0))",
    "K_STATS_StudentT_Survival(1.5,10)": "from scipy.stats import t; t.sf(1.5, 10)",
    "K_STATS_StudentT_Survival(30,10)": "from scipy.stats import t; t.sf(30, 10)",
    "K_STATS_StudentT_Survival(2,1E7)": "from scipy.stats import t; t.sf(2, 1e7)",
    "K_STATS_ChiSquare_Survival(200,3)": "from scipy.stats import chi2; chi2.sf(200, 3)",
    "K_STATS_ChiSquare_Cumulative(1E6,1E6)": "from scipy.stats import chi2; chi2.cdf(1e6, 1e6)",
    "K_STATS_F_Survival(100,5,10)": "from scipy.stats import f; f.sf(100, 5, 10)",
    "K_STATS_F_Survival(1,1E6,3)": "from scipy.stats import f; f.sf(1, 1e6, 3)",
    "K_STATS_Gamma_Cumulative(3,2,1)": "from scipy.stats import gamma; gamma.cdf(3, 2, scale=1)",
    "K_STATS_Gamma_Survival(60,2,1)": "from scipy.stats import gamma; gamma.sf(60, 2, scale=1)",
    "K_STATS_Gamma_Cumulative(1E6,1E6,1)": "from scipy.stats import gamma; gamma.cdf(1e6, 1e6, scale=1)",
    "K_STATS_Beta_Cumulative(0.3,2,3)": "from scipy.stats import beta; beta.cdf(0.3, 2, 3)",
    "K_STATS_Beta_Survival(0.999,2,3)": "from scipy.stats import beta; beta.sf(0.999, 2, 3)",
    "K_STATS_Exponential_Survival(50,1)": "from scipy.stats import expon; expon.sf(50, scale=1/1)",
    "K_STATS_Weibull_Survival(10,2,1)": "from scipy.stats import weibull_min; weibull_min.sf(10, 2, scale=1)",
    "K_STATS_Uniform_Cumulative(0.3,0,1)": "from scipy.stats import uniform; uniform.cdf(0.3, loc=0, scale=1)",
    "K_STATS_Binomial_Cumulative(5,10,0.5)": "from scipy.stats import binom; binom.cdf(5, 10, 0.5)",
    "K_STATS_Binomial_Survival(90,100,0.5)": "from scipy.stats import binom; binom.sf(90, 100, 0.5)",
    "K_STATS_Binomial_LogPMF(900,1000,0.5)": "from scipy.stats import binom; binom.logpmf(900, 1000, 0.5)",
    "K_STATS_Poisson_Cumulative(3,2)": "from scipy.stats import poisson; poisson.cdf(3, 2)",
    "K_STATS_Poisson_Survival(60,2)": "from scipy.stats import poisson; poisson.sf(60, 2)",
    "K_STATS_NegativeBinomial_Cumulative(5,3,0.5)": "from scipy.stats import nbinom; nbinom.cdf(5, 3, 0.5)",
    "K_STATS_Hypergeometric_PMF(2,5,10,50)": "from scipy.stats import hypergeom; hypergeom.pmf(2, M=50, n=10, N=5)",
    "K_STATS_DiscreteUniform_Cumulative(3,1,10)": "from scipy.stats import randint; randint.cdf(3, 1, 10 + 1)",
}

DEFAULT_CASE = "normal survival, z = 8"


def read_grid():
    with GRID.open(encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle, delimiter="|"))


def run(code):
    """The value of the last expression of a one-line =PY() snippet."""
    statements = [part.strip() for part in code.split(";")]
    scope = {}
    exec("\n".join(statements[:-1]), scope)
    return float(eval(statements[-1], scope))


def main():
    import mpmath
    import numpy
    import scipy

    cases = []
    for row in read_grid():
        call = row["kstats_call"].strip()
        code = PYTHON[call]
        excel = row["excel_formula"].strip()
        cases.append({
            "label": row["label"].strip(),
            "excel": None if excel == "-" else excel,
            "library": call,
            "reference": row["reference"].strip(),
            "python": code,
            "scipy": repr(run(code)),
            "why": row["why"].strip(),
        })
    data = {
        "schema": "comparison-cases-1",
        "source": "benchmark/excel_comparison/excel_comparison_grid.csv",
        "default_case": DEFAULT_CASE,
        "reference": {"method": "mpmath at 50 significant digits, "
                                "benchmark/excel_comparison/generate_excel_comparison.py",
                      "digits": 50},
        "scipy": {"python": platform.python_version(), "numpy": numpy.__version__,
                  "scipy": scipy.__version__, "mpmath": mpmath.__version__},
        "grades": {"reference": 12, "report": 6},
        "cases": cases,
    }
    OUTPUT.write_text(json.dumps(data, indent=1) + "\n", encoding="utf-8")
    print(OUTPUT, len(cases), "cases")
    return 0


if __name__ == "__main__":
    sys.exit(main())
