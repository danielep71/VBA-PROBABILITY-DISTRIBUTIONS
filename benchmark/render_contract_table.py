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


def write_into_readme(readme_path=None):
    """Splice the rendered table between the BEGIN/END generated markers in README.md."""
    if readme_path is None:
        readme_path = os.path.join(HERE, "README.md")
    begin = "<!-- BEGIN generated: accuracy_contracts.csv via render_contract_table.py. Do not hand-edit. -->"
    end = "<!-- END generated -->"
    with open(readme_path, "r", newline="") as f:
        raw = f.read()
    text = raw.replace("\r\n", "\n")
    i = text.find(begin); j = text.find(end)
    if i == -1 or j == -1 or j < i:
        raise SystemExit("README markers not found; add the BEGIN/END generated markers first.")
    new_block = begin + "\n\n" + render() + "\n\n" + end
    text = text[:i] + new_block + text[j + len(end):]
    with open(readme_path, "w", newline="") as f:
        f.write(text.replace("\n", "\r\n"))
    print(f"updated {readme_path} between generated markers")


if __name__ == "__main__":
    import sys
    if "--write" in sys.argv:
        write_into_readme()
    else:
        print(render())
