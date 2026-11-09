"""Install a test-only VBA port and exact-output checks on revealed v2 points."""
import json
from pathlib import Path
from install_cap_seam_tests import literal

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]


def main():
    template = (HERE / 'precision_v2_vba.template').read_text()
    template = template.replace('Failed', 'Err_Handler')
    template = template.replace('@LN2HI@', literal(float.fromhex('0x1.62e42fefa39efp-1')))
    template = template.replace('@LN2LO@', literal(float.fromhex('0x1.abc9e3b39803fp-56')))
    coefficients = json.loads((HERE / 'precision_coefficients.json').read_text())['coefficients_hi_lo_hex']
    lines = []
    for i, (hi, lo) in enumerate(reversed(coefficients)):
        const = f'Array({literal(float.fromhex(hi))}, {literal(float.fromhex(lo))})'
        lines.append(f'    V = {const}' if i == 0 else f'    V = PV2_Add(PV2_Mul(V, U), {const})')
    template = template.replace('@COEFFICIENTS@', '\n'.join(lines))
    lines = []
    for line in template.splitlines():
        lines.append(line)
        if line.startswith('Private Function '):
            name = line.split('(')[0].split()[-1]
            lines += ["'==============================================================================",
                      "' PURPOSE", f"'   {name}: test-only arithmetic port of the frozen Precision v2 model.",
                      "' NOTES", "'   Explicit Double stores preserve the operation boundaries under test.",
                      "'   This restricted diagnostic implementation is not a public kernel.",
                      "' ERROR POLICY", "'   Errors propagate to PV2_Pair, which returns False without acceptance.",
                      "' UPDATED", "'   2026-11-08",
                      "'=============================================================================="]
    lines += ['Private Sub Test_PrecisionV2_Fidelity()',
              "'==============================================================================",
              "' PURPOSE", "'   Compare the VBA port with all 250 revealed Python pairs bit for bit.",
              "' NOTES", "'   This is arithmetic fidelity evidence, not a fresh independent holdout.",
              "' UPDATED", "'   2026-11-08",
              "'==============================================================================",
              '    Dim P As Double, Q As Double, OK As Boolean']
    rows = json.loads((HERE / 'precision_v2_screen.json').read_text())['rows']
    assert len(rows) == 250
    for i, row in enumerate(rows):
        a, x = [literal(float.fromhex(row[k])) for k in ['a_hex', 'x_hex']]
        lines.append(f'    OK = PV2_Pair({a}, {x}, P, Q)')
        for j, var in enumerate(['P', 'Q']):
            expected = literal(float.fromhex(row['values_hex'][j]))
            lines += [f'    AssertTrue "PV2 fidelity {i:03d} {var}" & _',
                      f'        " observed=" & GammaCF_FullPrecision({var}), _',
                      f'        (OK And {var} = {expected})']
    lines += ['End Sub', '', '']
    target = ROOT / 'tests/M_STATS_PROBDIST_TEST.bas'
    text = target.read_text()
    assert 'Private Function PV2_H(' not in text
    text = text.replace('Private Sub Test_GammaCF_PrecisionCounterexample()',
                        '\n'.join(lines) + 'Private Sub Test_GammaCF_PrecisionCounterexample()', 1)
    text = text.replace('    Test_GammaCF_PrecisionCounterexample\n',
                        '    Test_GammaCF_PrecisionCounterexample\n    Test_PrecisionV2_Fidelity\n', 1)
    target.write_bytes(text.replace('\n', '\r\n').encode('ascii'))
    policy = ROOT / '.github/excel-evidence-policy.json'
    text = policy.read_text()
    assert '"expected_assertions": 1947' in text
    policy.write_text(text.replace('"expected_assertions": 1947', '"expected_assertions": 2447'))


if __name__ == '__main__':
    main()
