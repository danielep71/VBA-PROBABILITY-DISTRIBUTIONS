"""Render the benchmark README accuracy table from accuracy_contracts.csv.

Single source of truth: edit accuracy_contracts.csv, then regenerate the table
below with `python3 render_contract_table.py`. Do not hand-edit it in README.md.
"""
import csv, os

HERE = os.path.dirname(os.path.abspath(__file__))

def render(contracts_path=None):
    if contracts_path is None:
        contracts_path = os.path.join(HERE, "accuracy_contracts.csv")
    with open(contracts_path, newline="") as f:
        contracts = list(csv.DictReader(f))
    active = [r for r in contracts if r["status"] == "active"]
    limits = [r for r in contracts if r["status"] != "active"]

    lines = ["| Function | Metric | Threshold | Domain |", "|---|---|---|---|"]
    for r in sorted(active, key=lambda r: r["function"]):
        lines.append(f"| {r['function']} | {r['metric']} | {r['threshold']} | {r['domain']} |")
    out = "\n".join(lines)
    if limits:
        out += "\n\n**Known limitations**\n\n| Function | Domain | Notes |\n|---|---|---|\n"
        for r in limits:
            out += f"| {r['function']} | {r['domain']} | {r['notes']} |\n"
    return out

if __name__ == "__main__":
    print(render())
