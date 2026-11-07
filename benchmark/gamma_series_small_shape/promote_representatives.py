"""Reference-only #23 promotion: original counterexamples and all scope probes.

Use with benchmark/promote_grid_rows.py; never copy study observations.
The already observed holdout remains in its original evidence set.
"""
from pathlib import Path
import importlib.util
import sys
from decimal import Decimal, DecimalException, localcontext

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from _contract_eval import parse_observed, parse_reference, calculate_error


def meets_frozen_bound(row):
    """Enforce the study's 3E-15 bound even when a family registry is broader."""
    try:
        with localcontext() as context:
            context.prec = 80
            observed = parse_observed(row.get("observed_vba", ""))
            reference = parse_reference(row.get("reference", ""))
            if (observed is None or reference is None or not observed.is_finite()
                    or not reference.is_finite() or reference <= 0):
                return False
            return calculate_error(observed, reference, "relative") <= Decimal("3E-15")
    except (DecimalException, ValueError):
        return False



def build_rows():
    path = Path(__file__).with_name("generate_gamma_series_small_shape.py")
    spec = importlib.util.spec_from_file_location("gsss_reference", path)
    if spec is None or spec.loader is None:
        raise RuntimeError("Cannot load the frozen study generator")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    rows = []
    for source in module.build():
        if source["regime"] not in ("counterexample", "scope_probe"):
            continue
        row = {key: source[key] for key in
               ("function", "vba_kernel", "claim", "metric", "arg1", "arg2", "arg3", "reference")}
        row.update(arg4="", observed_vba="", regime="all", evidence_set="main grid",
                   expected_error="")
        rows.append(row)
    return rows
