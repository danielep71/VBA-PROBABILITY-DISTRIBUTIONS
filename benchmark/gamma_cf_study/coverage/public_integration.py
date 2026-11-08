"""Generate and exactly reverse the public adapter around the frozen Excel port.

Reversal is solely for historical source-isolation fixtures. It verifies every
new/replaced byte before restoring the old public bodies and negative controls.
It does not relabel any old Excel artifact as certifying current public code.
"""
import argparse
from functools import lru_cache
import json
import re
import subprocess
from install_scaled_fidelity import ROOT,literal,documented

PORT='b8efaf3c7810c48fadbfece38950ded190f79a56'
SRC='src/M_STATS_PROBDIST_SPECIALFUNCS.bas'
TEST='tests/M_STATS_PROBDIST_TEST.bas'
BEGIN="' BEGIN generated public Gamma integration"
END="' END generated public Gamma integration"
TBEGIN="' BEGIN public Gamma bit types"
TEND="' END public Gamma bit types"
EXTRA=2334


@lru_cache(maxsize=None)
def original(path):
    return subprocess.check_output(['git','show',PORT+':'+path],cwd=ROOT).decode().replace('\r\n','\n')


def procedures(text):
    return {
        m[2]:m[0] for m in re.finditer(r'^(?:Private|Public) (Function|Sub) (\w+)\b[\s\S]*?^End \1\b[^\n]*\n',text,re.M)}


@lru_cache(maxsize=1)
def arithmetic():
    procs=procedures(original(TEST));needed={'WG_Pair'};pending=['WG_Pair']
    while pending:
        name=pending.pop()
        for token in re.findall(r'\b\w+\b',procs[name]):
            if token in procs and token not in needed:
                needed.add(token);pending.append(token)
    assert not any('Case' in n or n.startswith('Test_') for n in needed)
    mapping={n:'PROB_DD_'+n for n in needed}
    mapping.update(SP_Number='PROB_DD_Number',SP_Words='PROB_DD_Words')
    text='\n'.join(procs[n] for n in procs if n in needed)
    text=re.sub(r'\b\w+\b',lambda m:mapping.get(m[0],m[0]),text)
    text=text.replace('private exponent-separated arithmetic fixture.','Gamma exponent-separated arithmetic.')
    text=text.replace('Study-only: no production routing or probability acceptance change.',
                      'Operation graph retained from frozen Excel port b8efaf3.')
    types="""Private Type PROB_DD_Number
    Value As Double
End Type
Private Type PROB_DD_Words
    Low As Long
    High As Long
End Type
"""
    return BEGIN+'\n'+text+'\n'+END+'\n\n',TBEGIN+'\n'+types+TEND+'\n\n',sorted(needed)


def adapter(tail):
    name='PROB_TryGammaRegularized'+tail
    return f'''Public Function {name}( _
    ByVal A As Double, _
    ByVal X As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' PURPOSE
'   Computes the regularized Gamma {tail} tail through the scaled two-Double pair.
' PRECONDITION
'   Finite A > 0 and X >= 0. X = 0 is returned exactly.
' METHOD
'   Lower Q is reconstructed directly from log P, before binary64 rounding.
'   The frozen wide kernel handles the region outside the extended interior.
' ERROR POLICY
'   False clears Result and reports invalid inputs, arithmetic failure or cap.
'   Retains the established 100000-iteration cap diagnostic wording.
' UPDATED
'   2026-11-08 - #24/#26 public integration candidate; acceptance pending.
'==============================================================================
    Dim P As Double, Q As Double
    Result = 0#: FailMsg = ""
    If Not PROB_DD_WG_Pair(A, X, P, Q, FailMsg) Then
        If InStr(1, FailMsg, "continued fraction cap 100000", vbBinaryCompare) > 0 Then
            FailMsg = "Incomplete gamma continued fraction failed to converge in 100000 iterations for A = " & A
        ElseIf InStr(1, FailMsg, "lower series cap 100000", vbBinaryCompare) > 0 Then
            FailMsg = "Incomplete gamma series failed to converge in 100000 iterations for A = " & A
        ElseIf Len(FailMsg) = 0 Then
            FailMsg = "Gamma kernel requires finite A > 0 and X >= 0"
        End If
        Exit Function
    End If
    If Not PROB_IsFinite({tail}) Or {tail} < 0# Or {tail} > 1# Then
        FailMsg = "Gamma kernel produced an invalid probability"
        Exit Function
    End If
    Result = {tail}
    {name} = True
End Function
'''


@lru_cache(maxsize=1)
def test_block():
    old=procedures(original(SRC));lines=[]
    for tail in ('P','Q'):
        name='PROB_TryGammaRegularized'+tail
        lines.append(old[name].replace('Public Function','Private Function',1).replace(name,'PG_B6_'+tail))
    lines.append('''Private Sub PG_Case(ByVal Label As String, ByVal A As Double, ByVal X As Double, _
    ByVal ExpectedP As Double, ByVal ExpectedQ As Double)
    Dim P As Double, Q As Double, OK As Boolean, Diag As String, Value As Variant
    OK = PROB_TryGammaRegularizedP(A, X, P, Diag)
    AssertTrue Label & " P returned " & Diag, OK
    AssertTrue Label & " P bits " & SG_Bits(P), (OK And Diag = "" And SP_SameBits(P, ExpectedP))
    OK = PROB_TryGammaRegularizedQ(A, X, Q, Diag)
    AssertTrue Label & " Q returned " & Diag, OK
    AssertTrue Label & " Q bits " & SG_Bits(Q), (OK And Diag = "" And SP_SameBits(Q, ExpectedQ))
    Value = K_STATS_Gamma_Cumulative(2# * X, A, 2#, Diag)
    PG_Value Label & " Gamma CDF", Value, ExpectedP, Diag
    Value = K_STATS_Gamma_Survival(2# * X, A, 2#, Diag)
    PG_Value Label & " Gamma SF", Value, ExpectedQ, Diag
    Value = K_STATS_ChiSquare_Cumulative(2# * X, 2# * A, Diag)
    PG_Value Label & " Chi-square CDF", Value, ExpectedP, Diag
    Value = K_STATS_ChiSquare_Survival(2# * X, 2# * A, Diag)
    PG_Value Label & " Chi-square SF", Value, ExpectedQ, Diag
End Sub

Private Sub PG_Value(ByVal Label As String, ByVal Value As Variant, ByVal Expected As Double, ByVal Diag As String)
    If IsError(Value) Then
        AssertTrue Label & " returned error " & Diag, False
    Else
        AssertTrue Label & " bits " & SG_Bits(CDbl(Value)) & " " & Diag, (Diag = "" And SP_SameBits(CDbl(Value), Expected))
    End If
End Sub

Private Sub PG_Invalid(ByVal A As Double, ByVal X As Double)
    Dim P As Double, Q As Double, OK As Boolean, Diag As String
    P = 9#: Diag = "stale"
    OK = PROB_TryGammaRegularizedP(A, X, P, Diag)
    AssertTrue "PG invalid P cleared", (Not OK And P = 0# And Len(Diag) > 0 And Diag <> "stale")
    Q = 9#: Diag = "stale"
    OK = PROB_TryGammaRegularizedQ(A, X, Q, Diag)
    AssertTrue "PG invalid Q cleared", (Not OK And Q = 0# And Len(Diag) > 0 And Diag <> "stale")
End Sub
''')
    data=json.loads(original('benchmark/gamma_cf_study/coverage/wide_gamma_fidelity.json'));names=[]
    for start in range(0,len(data['rows']),20):
        name=f'Test_PublicGamma_Chunk{start//20}';names.append(name);lines.append(f'Private Sub {name}()')
        for i,r in enumerate(data['rows'][start:start+20],start):
            lines += [f'    PG_Case "PG {i:03d}", {literal(r["a"])}, {literal(r["x"])}, _',
                      f'        {literal(r["values_hex"][0])}, {literal(r["values_hex"][1])}']
        lines+=['End Sub','']
    lines+=['Private Sub Test_PublicGamma_Integration()']+['    '+n for n in names]
    for a,x in data['invalid']:lines.append(f'    PG_Invalid {literal(a)}, {literal(x)}')
    lines+=['End Sub','']
    return BEGIN+'\n'+documented('\n'.join(lines))+'\n'+END+'\n\n'


def change_controls(text,reverse=False):
    procs=procedures(text)
    for name in ('EG_RoutingControl','WG_RoutingControl'):
        if name not in procs:continue
        old=procs[name];new=old
        for tail in ('P','Q'):
            a,b='PROB_TryGammaRegularized'+tail,'PG_B6_'+tail
            if reverse:a,b=b,a
            new=new.replace(a,b)
        text=text.replace(old,new,1)
    return text


def restore(text):
    """Exact inverse of this generator; only for historical replay fixtures."""
    if BEGIN not in text:return text
    if 'Attribute VB_Name = "M_STATS_PROBDIST_SPECIALFUNCS"' in text:
        block,types,_=arithmetic()
        if text.count(block)!=1 or text.count(types)!=1:raise ValueError('public Gamma block drift')
        text=text.replace(block,'',1).replace(types,'',1)
        old=procedures(original(SRC))
        for tail in ('P','Q'):
            current=adapter(tail)
            if text.count(current)!=1:raise ValueError('public Gamma adapter drift')
            text=text.replace(current,old['PROB_TryGammaRegularized'+tail],1)
        return text
    block=test_block()
    if text.count(block)!=1:raise ValueError('public Gamma test block drift')
    text=text.replace(block,'',1).replace('    Test_PublicGamma_Integration\n','')
    return change_controls(text,reverse=True)


def install():
    text=restore((ROOT/SRC).read_text());old=procedures(text);block,types,names=arithmetic()
    for tail in ('P','Q'):text=text.replace(old['PROB_TryGammaRegularized'+tail],adapter(tail),1)
    text=text.replace('Public Function PROB_LogGamma(',types+'Public Function PROB_LogGamma(',1)
    text+=block
    (ROOT/SRC).write_bytes(text.replace('\n','\r\n').encode('ascii'))
    text=restore((ROOT/TEST).read_text());text=change_controls(text)
    text=text.replace("' BEGIN generated wide-Gamma study",test_block()+"' BEGIN generated wide-Gamma study",1)
    text=text.replace('    Test_WideGamma_Fidelity\n','    Test_WideGamma_Fidelity\n    Test_PublicGamma_Integration\n',1)
    (ROOT/TEST).write_bytes(text.replace('\n','\r\n').encode('ascii'))
    p=ROOT/'.github/excel-evidence-policy.json';policy=json.loads(p.read_text())
    assert policy['expected_assertions'] in (10848,10848+EXTRA)
    policy['expected_assertions']=10848+EXTRA;p.write_text(json.dumps(policy,indent=2)+'\n')
    print(len(names),'arithmetic procedures;',EXTRA,'new assertions;',policy['expected_assertions'],'total')


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--write',action='store_true',required=True);parser.parse_args();install()
