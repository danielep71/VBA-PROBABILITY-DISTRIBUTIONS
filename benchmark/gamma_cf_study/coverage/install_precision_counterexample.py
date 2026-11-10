"""Install the revealed upper-CF monotonicity counterexample, not a new holdout."""
from pathlib import Path
from install_cap_seam_tests import literal

ROOT = Path(__file__).resolve().parents[3]


def main():
    a = literal(.39)
    left = literal(float.fromhex('0x1.63d70a3d70a41p+0'))
    right = literal(float.fromhex('0x1.63d70a3d70a42p+0'))
    lines = ['Private Sub Test_GammaCF_PrecisionCounterexample()',
             "'==============================================================================",
             "' PURPOSE",
             "'   Check the revealed A=.39 adjacent-X counterexample within upper CF.",
             "' NOTES",
             "'   Coordinates were revealed by the frozen 220-point precision screen.",
             "'   Both points are above A+1; this is not independent holdout evidence.",
             "' ERROR POLICY",
             "'   A wrong-direction step fails and retains full-precision diagnostics.",
             "' UPDATED", "'   2026-11-08",
             "'=============================================================================="]
    for fn, op in [('Cumulative', '<='), ('Survival', '>=')]:
        call1 = f'CDbl(K_STATS_Gamma_{fn}({left}, {a}, 1#))'
        call2 = f'CDbl(K_STATS_Gamma_{fn}({right}, {a}, 1#))'
        lines += [f'    AssertTrue "Precision upper CF {fn}" & _',
                  f'        " left=" & GammaCF_FullPrecision({call1}) & _',
                  f'        " right=" & GammaCF_FullPrecision({call2}), _',
                  f'        ({call1} {op} _', f'         {call2})']
    lines += ['End Sub', '', '']
    path = ROOT / 'tests/M_STATS_PROBDIST_TEST.bas'
    source = path.read_text()
    assert 'Private Sub Test_GammaCF_PrecisionCounterexample()' not in source
    source = source.replace('Private Sub Test_GammaCF_CapAndXSeam()',
                            '\n'.join(lines) + 'Private Sub Test_GammaCF_CapAndXSeam()', 1)
    source = source.replace('    Test_GammaCF_CapAndXSeam\n',
                            '    Test_GammaCF_CapAndXSeam\n    Test_GammaCF_PrecisionCounterexample\n', 1)
    path.write_bytes(source.replace('\n', '\r\n').encode('ascii'))
    policy = ROOT / '.github/excel-evidence-policy.json'
    text = policy.read_text()
    assert '"expected_assertions": 1945' in text
    policy.write_text(text.replace('"expected_assertions": 1945', '"expected_assertions": 1947'))


if __name__ == '__main__':
    main()
