"""
Shared reference and merge code for the three F-envelope sweeps (#159).

Each reference is evaluated at the exact binary64 value of the argument token
the row stores, which is what the exporter passes. Before #159 the base and gap
sweeps evaluated x at its decimal token ("0.3" as three tenths) and the
both-large sweep at the 17-digit decimal of a Double; at these degrees of
freedom the survival is sensitive enough that half an ulp of x cost up to 16 of
the 30 printed digits.

The survival is computed directly as the upper tail
I_{d2/(d1 x + d2)}(d2/2, d1/2) (_ibeta.f_sf), never as 1 - CDF.
"""
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
from _binary64 import exact_binary64_mpf as D  # noqa: E402
from _ibeta import f_cdf, f_sf  # noqa: E402

FIELDS = ["function", "vba_kernel", "claim", "metric", "arg1", "arg2", "arg3",
          "reference", "observed_vba", "regime", "evidence_set"]


def rows_for(xs, d1, d2, tag):
    """The F_Cumulative and F_Survival rows for one point, at its Doubles."""
    a2, a3 = mp.nstr(mp.mpf(d1), 17), mp.nstr(mp.mpf(d2), 17)
    x, n1, n2 = D(xs), D(a2), D(a3)
    out = []
    for fn, ref in (("F_Cumulative", f_cdf(x, n1, n2)), ("F_Survival", f_sf(x, n1, n2))):
        out.append({"function": fn, "vba_kernel": f"K_STATS_{fn}", "claim": "rel<=1.1E-10",
                    "metric": "rel", "arg1": xs, "arg2": a2, "arg3": a3,
                    "reference": mp.nstr(ref, 30), "observed_vba": "",
                    "regime": "extreme_df", "evidence_set": f"f_envelope:{tag}"})
    return out


def merge(grid_name, rows, write):
    """Re-reference the committed grid; observations are never touched."""
    from _study_merge import merge_references, print_report, study_describe, study_key
    changes, held = merge_references(os.path.join(HERE, grid_name), rows, study_key,
                                     {"observed_vba"}, write=write,
                                     describe=study_describe)
    print_report(grid_name, changes, held, write)
