import json,math
from pathlib import Path
import mpmath as mp
base=Path('benchmark/gamma_cf_study/coverage')
def lit(v):
 v=float(v);h=format(v,'.14e');return '('+h.upper()+' + ('+format(v-float(h),'.14e').upper()+'))'
def quantile(a,p,dps):
 with mp.workdps(dps):
  a,p=mp.mpf(a),mp.mpf(p);lo=mp.mpf(0);hi=mp.mpf(1)
  while mp.gammainc(a,0,hi,regularized=True)<p:hi*=2
  for _ in range(2400):
   mid=(lo+hi)/2
   if mp.gammainc(a,0,mid,regularized=True)<p:lo=mid
   else:hi=mid
   if lo>0 and (hi-lo)/lo<mp.mpf(10)**(-dps+5):return (lo+hi)/2
  raise RuntimeError('inverse reference did not converge')
lines=['Private Sub Test_GammaCF_AdditionalCallers()',"'==============================================================================","' PURPOSE","'   Checks inverse callers, Poisson, arithmetic underflow and routing seams.","' NOTES","'   References and design: benchmark/gamma_cf_study/coverage/.","'   Frozen thresholds apply; subnormal outputs use a one-unit absolute bound.","' ERROR POLICY","'   Failures remain recorded; do not retune the kernel from these cases.","' UPDATED","'   2026-11-08","'==============================================================================",'    Dim Small As Double','    Dim A As Double','    Dim V As Double','    Dim L As Double','    Dim R As Double','    Dim N As Double','    Dim Diag As String','    Small = (2# ^ (-1022)) * (2# ^ (-52))']
records=[];count=0
def assertion(s):
 global count
 lines.extend(s);count+=1
for a in [.01,.5,1.,2.]:
 for p in [.1,.5,.9]:
  lo=quantile(a,p,90);hi=quantile(a,p,120)
  with mp.workdps(120):
   assert abs(lo/hi-1)<mp.mpf('1e-80')
   records.append({'kind':'inverse','a_hex':a.hex(),'p_hex':p.hex(),'reference':mp.nstr(hi,100)})
  for fn,args,ref,tol in [('Gamma',f'{lit(p)}, {lit(a)}, 1#',float(hi),'2E-14'),('ChiSquare',f'{lit(p)}, {lit(2*a)}',float(2*hi),'4.7E-12')]:
   assertion([f'    AssertClose "B6 {fn} inverse {a} {p}", _',f'        CDbl(K_STATS_{fn}_InverseCumulative({args})) / {lit(ref)}, 1#, {tol}'])
for k in [0,1,10]:
 for mean in [math.nextafter(float(k+2),0),float(k+2),math.nextafter(float(k+2),math.inf)]:
  with mp.workdps(120):
   cdf=mp.exp(-mp.mpf(mean))*sum(mp.mpf(mean)**j/mp.factorial(j) for j in range(k+1));sf=1-cdf
  for fn,ref in [('Cumulative',cdf),('Survival',sf)]:
   assertion([f'    AssertClose "B6 Poisson {fn} {k} {mean.hex()}", _',f'        CDbl(K_STATS_Poisson_{fn}({k}#, {lit(mean)})) / {lit(ref)}, 1#, 2E-13'])
for mean in [.5,2.,12.]:
 for p in [.1,.5,.9]:
  lines += [f'    N = CDbl(K_STATS_Poisson_InverseCumulative({lit(p)}, {lit(mean)}))']
  assertion([f'    AssertTrue "B6 Poisson inverse upper {mean} {p}", _',f'        (CDbl(K_STATS_Poisson_Cumulative(N, {lit(mean)})) >= {lit(p)})'])
  assertion(['    If N = 0# Then', '        AssertTrue "B6 Poisson lower support bound", True', '    Else', f'        AssertTrue "B6 Poisson inverse lower {mean} {p}", _',f'            (CDbl(K_STATS_Poisson_Cumulative(N - 1#, {lit(mean)})) < {lit(p)})', '    End If'])
for a,expr in [(2.**-1074,'Small'),(2.**-1022,'2# ^ (-1022)'),(1e-300,'1E-300'),(.5,'0.5'),(1.,'1#')]:
 lines += ['    A = '+expr]
 for x in [2.,700.,740.,745.]:
  with mp.workdps(140):q=mp.gammainc(mp.mpf(a),mp.mpf(x),mp.inf,regularized=True)
  for fn,args in [('Gamma',f'{x}#, A, 1#'),('ChiSquare',f'{2*x}#, 2# * A')]:
   call=f'CDbl(K_STATS_{fn}_Survival({args}))';name=f'B6 underflow {fn} {a.hex()} {x}'
   if q<mp.mpf(2)**-1022:
    units=float(q)/2.**-1074
    assertion([f'    AssertClose "{name}", _',f'        {call} / Small, {lit(units)}, 1#'])
   else:assertion([f'    AssertClose "{name}", _',f'        {call} / {lit(q)}, 1#, 5E-15'])
for x in [2.1,4.,8.,16.]:
 lines += [f'    L = CDbl(K_STATS_Gamma_Survival({lit(x)}, 1# - 2# ^ (-53), 1#))',f'    V = CDbl(K_STATS_Gamma_Survival({lit(x)}, 1#, 1#))',f'    R = CDbl(K_STATS_Gamma_Survival({lit(x)}, 1# + 2# ^ (-52), 1#))']
 assertion([f'    AssertTrue "B6 shape seam lower {x}", (L <= V)']);assertion([f'    AssertTrue "B6 shape seam upper {x}", (V <= R)'])
for fn,args in [('Gamma_Survival','2#, 0#, 1#'),('Gamma_Cumulative','2#, 1#, 0#'),('ChiSquare_Survival','4#, 0#'),('Poisson_Survival','0#, -1#')]:
 assertion([f'    AssertErrorCode "B6 invalid {fn}", K_STATS_{fn}({args}), xlErrNum'])
lines += ['    Diag = "stale"']
assertion(['    AssertErrorCode "B6 status failure", K_STATS_Gamma_Survival(2#, 0#, 1#, Diag), xlErrNum'])
assertion(['    AssertTrue "B6 status populated", (Len(Diag) > 0 And Diag <> "stale")'])
lines+=['    V = CDbl(K_STATS_Gamma_Survival(2#, 1#, 1#, Diag))']
assertion(['    AssertTrue "B6 status cleared", (Len(Diag) = 0)'])
lines+=['End Sub','','']
p=Path('tests/M_STATS_PROBDIST_TEST.bas');s=p.read_bytes().decode('ascii');assert 'Private Sub Test_GammaCF_AdditionalCallers()' not in s;s=s.replace('Private Sub Test_GammaCF_Coverage()', '\r\n'.join(lines)+'Private Sub Test_GammaCF_Coverage()',1);s=s.replace('    Test_GammaCF_Coverage\r\n','    Test_GammaCF_Coverage\r\n    Test_GammaCF_AdditionalCallers\r\n',1);p.write_bytes(s.encode('ascii'))
p=Path('.github/excel-evidence-policy.json');s=p.read_text();p.write_text(s.replace('"expected_assertions": 1766',f'"expected_assertions": {1766+count}'))
(base/'inverse_references.json').write_text(json.dumps(records,indent=2)+'\n')
print('Added',count,'assertions; total',1766+count)
