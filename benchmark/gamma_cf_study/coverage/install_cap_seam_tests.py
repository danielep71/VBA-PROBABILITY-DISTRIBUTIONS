"""Install natural-cap and adjacent-X seam checks; leave production budgets intact."""
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]


def literal(value):
    high = format(value, '.14e')
    low = format(value - float(high), '.14e')
    return f'({high.upper()} + ({low.upper()}))'


def main():
    points = json.loads((HERE / 'protocol.json').read_text())['points']
    lines = [
        'Private Sub Test_GammaCF_CapAndXSeam()',
        "'==============================================================================",
        "' PURPOSE",
        "'   Verifies real iteration-cap failure propagation and adjacent-X monotonicity.",
        "' WHY",
        "'   A red accuracy probe must not hide a partial result reported as success.",
        "' BEHAVIOR",
        "'   Uses the unchanged 100000-iteration cap at A=1E14, X=A+1.",
        "'   Reuses the frozen boundary coordinates; this is not a fresh holdout.",
        "' ERROR POLICY",
        "'   Require False with a cap diagnostic in kernels and #NUM! in wrappers.",
        "' UPDATED",
        "'   2026-11-08",
        "'==============================================================================",
        '    Dim A As Double', '    Dim X As Double', '    Dim V As Double',
        '    Dim Diag As String', '    Dim OK As Boolean',
        '    A = 1E+14', '    X = A + 1#',
    ]
    for fn in ['ContinuedFractionQ', 'RegularizedP', 'RegularizedQ']:
        lines += ['    Diag = vbNullString',
                  f'    OK = PROB_TryGamma{fn}(A, X, V, Diag)',
                  f'    AssertTrue "B6 natural cap {fn}", (Not OK)',
                  f'    AssertTrue "B6 cap diagnostic {fn}", _',
                  '        (InStr(1, Diag, "continued fraction failed to converge in 100000 iterations", vbTextCompare) > 0)']
    for fn in ['Cumulative', 'Survival']:
        lines += ['    Diag = vbNullString',
                  f'    AssertErrorCode "B6 cap Gamma {fn}", K_STATS_Gamma_{fn}(X, A, 1#, Diag), xlErrNum',
                  f'    AssertTrue "B6 public cap diagnostic {fn}", _',
                  '        (InStr(1, Diag, "continued fraction failed to converge in 100000 iterations", vbTextCompare) > 0)']
    lines += ['    V = CDbl(K_STATS_Gamma_Survival(2#, 1#, 1#, Diag))',
              '    AssertClose "B6 recovery after cap", V, Exp(-2#), 1E-15',
              '    AssertTrue "B6 recovery clears cap diagnostic", (Len(Diag) = 0)']
    for i in range(13):
        triple = points[5*i:5*i+3]
        a = literal(float.fromhex(triple[0]['a_hex']))
        xs = [literal(float.fromhex(p['x_hex'])) for p in triple]
        for fn, op in [('Cumulative', '<='), ('Survival', '>=')]:
            for side in range(2):
                lines += [f'    AssertTrue "B6 X seam {i:02d} {fn} {side}", _',
                          f'        (CDbl(K_STATS_Gamma_{fn}({xs[side]}, {a}, 1#)) {op} _',
                          f'         CDbl(K_STATS_Gamma_{fn}({xs[side+1]}, {a}, 1#)))']
    lines += ['End Sub', '', '']
    path = ROOT / 'tests/M_STATS_PROBDIST_TEST.bas'
    source = path.read_text()
    assert 'Private Sub Test_GammaCF_CapAndXSeam()' not in source
    source = source.replace('Private Sub Test_GammaCF_AdditionalCallers()',
                            '\n'.join(lines) + 'Private Sub Test_GammaCF_AdditionalCallers()', 1)
    source = source.replace('    Test_GammaCF_AdditionalCallers\n',
                            '    Test_GammaCF_AdditionalCallers\n    Test_GammaCF_CapAndXSeam\n', 1)
    path.write_bytes(source.replace('\n', '\r\n').encode('ascii'))
    policy = ROOT / '.github/excel-evidence-policy.json'
    text = policy.read_text()
    assert '"expected_assertions": 1881' in text
    policy.write_text(text.replace('"expected_assertions": 1881', '"expected_assertions": 1945'))
    print('Added 12 natural-cap/recovery and 52 X-seam assertions: total 1945')


if __name__ == '__main__':
    main()
