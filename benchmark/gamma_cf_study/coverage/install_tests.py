"""Install the frozen coverage assertions once from the 965-test baseline."""
import json
from pathlib import Path
assert b'Private Sub Test_GammaCF_Coverage()' not in Path('tests/M_STATS_PROBDIST_TEST.bas').read_bytes(), 'coverage already installed'
rows=json.loads(Path('benchmark/gamma_cf_study/coverage/references.json').read_text())
def lit(v):
 v=float(v);h=format(v,'.14e');lo=v-float(h);return '('+h.upper()+' + ('+format(lo,'.14e').upper()+'))'
lines=['Private Sub Test_GammaCF_Coverage()',"'==============================================================================","' PURPOSE","'   Exercises frozen #26 boundary and public-function points at three scales.","' NOTES","'   References use exact binary64 inputs, stable at 90/120 decimal digits.","'   See benchmark/gamma_cf_study/coverage/protocol.json; no threshold relaxation.","' UPDATED","'   2026-11-08","'=============================================================================="]
for i,r in enumerate(rows):
 lines += [f'    GammaCF_CoverageCase "B6 point {i:02d}", _',f'        {lit(float.fromhex(r["a_hex"]))}, {lit(float.fromhex(r["x_hex"]))}, _',f'        {lit(r["p"])}, {lit(r["q"])}']
lines+=['End Sub','','Private Sub GammaCF_CoverageCase( _','    ByVal Label As String, ByVal A As Double, ByVal X As Double, _','    ByVal PRef As Double, ByVal QRef As Double)',"'==============================================================================","' PURPOSE","'   Checks Gamma scale invariance, equivalent Chi-square calls and P/Q sum.","' ERROR POLICY","'   Numeric mismatches are recorded by the existing assertion harness.","' NOTES","'   All references are positive normal values; normalize before comparing.","' UPDATED","'   2026-11-08","'==============================================================================",'    Dim P As Double','    Dim Q As Double','    P = CDbl(K_STATS_Gamma_Cumulative(X, A, 1#))','    Q = CDbl(K_STATS_Gamma_Survival(X, A, 1#))','    AssertClose Label & " Gamma P", P / PRef, 1#, 3E-15','    AssertClose Label & " Gamma Q", Q / QRef, 1#, 5E-15']
for scale in ['0.5','2#']:
 for fn,ref,tol in [('Cumulative','PRef','3E-15'),('Survival','QRef','5E-15')]:
  lines += [f'    AssertClose Label & " Gamma {fn} scale {scale}", _',f'        CDbl(K_STATS_Gamma_{fn}(X * {scale}, A, {scale})) / {ref}, 1#, {tol}']
for fn,ref in [('Cumulative','PRef'),('Survival','QRef')]:
 lines += [f'    AssertClose Label & " Chi-square {fn}", _',f'        CDbl(K_STATS_ChiSquare_{fn}(2# * X, 2# * A)) / {ref}, 1#, 2.6E-10']
lines += ['    AssertClose Label & " P+Q", P + Q, 1#, 4.44089209850063E-16','End Sub','','']
p=Path('tests/M_STATS_PROBDIST_TEST.bas');s=p.read_bytes().decode('ascii');marker='Private Sub Test_CN_GammaSurvival()';assert marker in s;s=s.replace(marker,'\r\n'.join(lines)+marker,1);needle='    Debug.Print "-- Gamma survival"';assert needle in s;s=s.replace(needle,needle+'\r\n    Test_GammaCF_Coverage',1);p.write_bytes(s.encode('ascii'))
p=Path('.github/excel-evidence-policy.json');s=p.read_text();assert '"expected_assertions": 965' in s;p.write_text(s.replace('"expected_assertions": 965',f'"expected_assertions": {965+89*9}'))
print('Added 801 assertions; expected total 1766, not Excel verified')
