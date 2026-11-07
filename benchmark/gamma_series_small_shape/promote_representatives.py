"""Reference-only #23 promotion: original counterexamples and all scope probes.

Use with benchmark/promote_grid_rows.py; never copy study observations.
The already observed holdout remains in its original evidence set.
"""
from pathlib import Path
import importlib.util


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
