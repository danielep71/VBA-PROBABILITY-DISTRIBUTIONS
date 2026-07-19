"""Render the benchmark README accuracy table from accuracy_contracts.csv.

Single source of truth: edit accuracy_contracts.csv, then regenerate the table
with `python3 render_contract_table.py`. Do not hand-edit it in README.md.
Also lists numerical_limitations.csv (limitations that are not accuracy contracts).
"""
import csv, os

HERE = os.path.dirname(os.path.abspath(__file__))


def render(contracts_path=None, limitations_path=None):
    if contracts_path is None:
        contracts_path = os.path.join(HERE, "accuracy_contracts.csv")
    if limitations_path is None:
        limitations_path = os.path.join(HERE, "numerical_limitations.csv")
    with open(contracts_path, newline="") as f:
        contracts = list(csv.DictReader(f))

    lines = ["| Contract | Function | Regime | Measure | Metric | Threshold | Provenance |",
             "|---|---|---|---|---|---|---|"]
    for r in sorted(contracts, key=lambda r: r["contract_id"]):
        lines.append(f"| {r['contract_id']} | {r['function']} | {r['regime']} | "
                     f"{r['measure']} | {r['metric']} | {r['threshold']} | {r['provenance']} |")
    out = "\n".join(lines)

    if os.path.exists(limitations_path):
        with open(limitations_path, newline="") as f:
            lims = list(csv.DictReader(f))
        if lims:
            out += "\n\n**Numerical limitations** (documented, not accuracy contracts)\n\n"
            out += "| Limitation | Affected | Domain | Observed effect | Status |\n"
            out += "|---|---|---|---|---|\n"
            for r in lims:
                out += (f"| {r['limitation_id']} | {r['affected_functions']} | {r['domain']} "
                        f"| {r['observed_effect']} | {r['status']} |\n")
    return out


if __name__ == "__main__":
    print(render())
