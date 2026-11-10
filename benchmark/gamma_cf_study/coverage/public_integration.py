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
import sys
from install_scaled_fidelity import HERE,ROOT,literal,documented
sys.path.insert(0,str(ROOT/'benchmark'))
import _harness_timing as harness  # noqa: E402
import _harness_studies as studies  # noqa: E402

PORT='b8efaf3c7810c48fadbfece38950ded190f79a56'
SRC='src/M_STATS_PROBDIST_SPECIALFUNCS.bas'
TEST='tests/M_STATS_PROBDIST_TEST.bas'
BEGIN="' BEGIN generated public Gamma integration"
END="' END generated public Gamma integration"
TBEGIN="' BEGIN public Gamma bit types"
TEND="' END public Gamma bit types"
# 4030 integration assertions, plus 8 reach assertions (#242 review).
EXTRA=4040
# The scaled pair's tighter stopping rules (2^-68 series, 2^-101 continued
# fraction) cap at arguments the binary64 route answered: X = A = 1.2E8 and
# A = 5E11, X = A + 2 (wide_gamma.public raises its cap at both). References
# by quadrature around the peak at 40 and 60 digits; they agree to 7E-29 and
# match wide_gamma.public at (1E8, 1E8) to 2.9E-17.
REACH=(('1.2E+08','1.2E+08','0x1.00019754e24c4p-1','0x1.fffcd1563b679p-2'),
       ('500000000000#','500000000002#','0x1.00002c2c275e2p-1','0x1.ffffa7a7b143bp-2'))
# The extreme_shape contracts: Gamma_Cumulative 2E-12, Gamma_Survival 5E-13.
REACH_LIMITS=('0.000000000002','0.0000000000005')
# Kernels whose handlers keep only error 5, the kernel's own domain and cap
# signal; any other runtime error propagates to the public function (#VALUE!).
RERAISE=('WG_Pair','EG1_Pair','EG_Pair')
RERAISE_LINES=("    'Error 5 is this kernel's own domain or cap signal; any other runtime error\n"
               "    'is unexpected and propagates, so the public function reports #VALUE!.\n"
               "    If Err.Number <> 5 Then Err.Raise Err.Number, Err.Source, Err.Description\n")


# Comment-only amendments outside the generated block, so the modules describe
# the integrated kernel (#242 review). install() applies them and restore()
# reverses them exactly; test_public_integration.py checks that every old and
# new text is comment lines only.
AMENDMENTS={
'src/M_STATS_PROBDIST_SPECIALFUNCS.bas':[
("""'   - PROB_TryGammaSeriesP / PROB_TryGammaContinuedFractionQ:
'       Series expansion for x < a + 1, continued fraction for the upper tail,
'       in the arrangement of Numerical Recipes (gser / gcf).
""","""'   - PROB_TryGammaRegularizedP / PROB_TryGammaRegularizedQ:
'       A scaled double-double pair (#26; frozen Excel port b8efaf3,
'       benchmark/gamma_cf_study): lower series below x = a + 1, with Q taken
'       from log P before rounding, and the continued fraction above. Where the
'       pair reaches its 100000-iteration cap, PROB_TryGammaSeriesP or
'       PROB_TryGammaContinuedFractionQ answers instead.
'   - PROB_TryGammaSeriesP / PROB_TryGammaContinuedFractionQ:
'       Series expansion for x < a + 1, continued fraction for the upper tail,
'       in the arrangement of Numerical Recipes (gser / gcf), in binary64.
"""),
("""'   2026-10-13 - Standard normal upper tail added (#72).
""","""'   2026-10-13 - Standard normal upper tail added (#72).
'   2026-11-09 - Regularized gamma P/Q through the scaled double-double pair,
'                with the binary64 kernels as the fallback at its cap (#26).
"""),
("""'terms at 3E-14 and 45946 here, both far inside PROB_GAMMA_MAX_ITER. The two
'continued fractions gain nothing from the same change (the beta CF is limited
'elsewhere and is unchanged above A = 1E6), so they keep PROB_NUM_EPS.
""","""'terms at 3E-14 and 45946 here, both far inside PROB_GAMMA_MAX_ITER. The two
'continued fractions gain nothing from the same change (the beta CF is limited
'elsewhere and is unchanged above A = 1E6), so they keep PROB_NUM_EPS.
'Since #26 the public P and Q try the scaled double-double pair first. Its
'series stops at 2^-68 and needs about 1.32 times these terms (67921 at df 1E8),
'still inside its own 100000 cap; this binary64 series answers where it is not.
"""),
("""'     The gamma series is therefore the binding kernel: it exhausts a 100000
'     budget while the beta CF is still in the low thousands.
""","""'     The gamma series is therefore the binding kernel: it exhausts a 100000
'     budget while the beta CF is still in the low thousands. The scaled
'     double-double pair behind the public P and Q (#26) stops more tightly and
'     needs about 1.32 times the series' terms, so it reaches its own 100000 cap
'     first (near x = a from a of about 1.07E8, where the series still converges
'     to about 1.9E8); the public P and Q then fall back to the series.
"""),
("""' PURPOSE
'   Evaluates P(A, X) by the lower incomplete gamma series expansion.
""","""' PURPOSE
'   Evaluates P(A, X) by the lower incomplete gamma series expansion. Since #26
'   it is the binary64 route behind PROB_TryGammaRegularizedP and
'   PROB_TryGammaRegularizedQ where the scaled pair reaches its iteration cap.
"""),
("""' PURPOSE
'   Evaluates Q(A, X), the regularized upper incomplete gamma function, by the
'   modified Lentz continued fraction.
""","""' PURPOSE
'   Evaluates Q(A, X), the regularized upper incomplete gamma function, by the
'   modified Lentz continued fraction. Since #26 it is the binary64 route behind
'   PROB_TryGammaRegularizedP and PROB_TryGammaRegularizedQ where the scaled
'   pair reaches its iteration cap.
"""),
],
'src/M_STATS_PROBDIST_TFAMILY.bas':[
("""Private Const PROB_CHI_MAX_DF     As Double = 100000000# 'Chi-square envelope; validated to 1E8 (measured, benchmark/envelope_probe, 2026-07-29); above that the lower-tail series reaches PROB_GAMMA_MAX_ITER
""","""Private Const PROB_CHI_MAX_DF     As Double = 100000000# 'Chi-square envelope; validated to 1E8 (measured, benchmark/envelope_probe, 2026-07-29). The incomplete-gamma iteration caps lie above it: near X = df the scaled pair (#26) caps from df of about 2.1E8 and its binary64 fallback from about 3.8E8
"""),
],
'src/M_STATS_PROBDIST_CONTINUOUS.bas':[
("""'     benchmark/cdf_large_shape rather than restated here (the gamma series is
'     bounded by PROB_GAMMA_MAX_ITER, and the beta range moved substantially
""","""'     benchmark/cdf_large_shape rather than restated here (the incomplete-gamma
'     kernels stop at 100000-iteration caps, and the beta range moved substantially
"""),
],
}


def module_path(text):
    """The src path of a VBA module's text, from its VB_Name attribute."""
    match=re.match(r'Attribute VB_Name = "(\w+)"',text)
    return f'src/{match[1]}.bas' if match else None


def amended(path,text):
    """True when every amendment of the module is present, False when none is."""
    present=[text.count(new)==1 for _,new in AMENDMENTS.get(path,())]
    if any(present) and not all(present):raise ValueError(f'comment amendment drift in {path}')
    return bool(present) and all(present)


def amend(path,text,reverse=False):
    """Apply (or exactly reverse) the comment amendments of one module."""
    for old,new in AMENDMENTS.get(path,()):
        a,b=(new,old) if reverse else (old,new)
        if text.count(a)!=1:raise ValueError(f'comment amendment drift in {path}')
        text=text.replace(a,b,1)
    return text


def code_only(text):
    """VBA source without comments: a name in a comment is not a dependency."""
    out=[]
    for line in text.split('\n'):
        quoted=False
        for i,ch in enumerate(line):
            if ch=='"':quoted=not quoted
            elif ch=="'" and not quoted:
                line=line[:i];break
        out.append(line)
    return '\n'.join(out)


@lru_cache(maxsize=None)
def live_total():
    """Regression plus study assertions of the committed policies (#300)."""
    return studies.live_total(ROOT)


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
        for token in re.findall(r'\b\w+\b',code_only(procs[name])):
            if token in procs and token not in needed:
                needed.add(token);pending.append(token)
    assert not any('Case' in n or n.startswith('Test_') for n in needed)
    mapping={n:'PROB_DD_'+n for n in needed}
    mapping.update(SP_Number='PROB_DD_Number',SP_Words='PROB_DD_Words')
    chosen=[]
    for n in procs:
        if n not in needed:continue
        proc=procs[n]
        if n in RERAISE:
            assert proc.count('Err_Handler:\n')==1,n
            proc=proc.replace('Err_Handler:\n','Err_Handler:\n'+RERAISE_LINES,1)
        chosen.append(proc)
    text='\n'.join(chosen)
    text=re.sub(r'\b\w+\b',lambda m:mapping.get(m[0],m[0]),text)
    text=text.replace('private exponent-separated arithmetic fixture.','Gamma exponent-separated arithmetic.')
    text=text.replace('Study-only: no production routing or probability acceptance change.',
                      'Operation graph retained from frozen Excel port b8efaf3.')
    # The Precision v2 arithmetic is the production core under SGD_Try.
    text=text.replace('test-only arithmetic port of the frozen Precision v2 model.',
                      'double-double arithmetic of the frozen Precision v2 model.')
    text=text.replace('This restricted diagnostic implementation is not a public kernel.',
                      'Production arithmetic core under PROB_DD_SGD_Try; not a public entry point.')
    text=text.replace('Errors propagate to PV2_Pair, which returns False without acceptance.',
                      'Errors propagate to the calling Gamma kernel.')
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
    series='\n                Value = 1# - Value' if tail=='Q' else ''
    fraction='\n                Value = 1# - Value' if tail=='P' else ''
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
'   Where the pair reaches its 100000-iteration cap, the binary64 series or
'   continued fraction answers instead: the pair stops at 2^-68 and 2^-101, so
'   near X = A it caps from A of about 1.07E8, and just above X = A + 1 from
'   about 3.3E11, where the binary64 route still converges. That fallback
'   forms the complement, Q = 1 - P below X = A + 1 and P = 1 - Q above it.
'   It is the one place the lower branch forms 1 - P (#24), and it adds no
'   amplification: the pair caps only for A above about 1.07E8, where
'   P(A, A + 1) is about 0.5, so the complemented tail is at least about 0.5.
' ERROR POLICY
'   False clears Result and reports invalid inputs, arithmetic failure or cap.
'   The cap diagnostic keeps the established 100000-iteration wording.
'   An unexpected runtime error propagates to the caller.
' UPDATED
'   2026-11-09 - #24: the fallback's complement is bounded and documented.
'==============================================================================
    Dim P As Double, Q As Double, Value As Double
    Result = 0#: FailMsg = ""
    If Not PROB_DD_WG_Pair(A, X, P, Q, FailMsg) Then
        If InStr(1, FailMsg, "continued fraction cap 100000", vbBinaryCompare) > 0 _
            Or InStr(1, FailMsg, "lower series cap 100000", vbBinaryCompare) > 0 Then
            FailMsg = ""
            If X < A + 1# Then
                If Not PROB_TryGammaSeriesP(A, X, Value, FailMsg) Then Exit Function{series}
            Else
                If Not PROB_TryGammaContinuedFractionQ(A, X, Value, FailMsg) Then Exit Function{fraction}
            End If
            If Value < 0# Then Value = 0#
            If Value > 1# Then Value = 1#
            Result = Value
            {name} = True
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
    fresh=json.loads((HERE/'public_fresh_fidelity.json').read_text())
    for start in range(0,len(fresh['rows']),20):
        name=f'Test_PublicFresh_Chunk{start//20}';names.append(name);lines.append(f'Private Sub {name}()')
        for i,r in enumerate(fresh['rows'][start:start+20],start):
            lines += [f'    PG_Case "PF {i:03d}", {literal(r["a"])}, {literal(r["x"])}, _',
                      f'        {literal(r["values_hex"][0])}, {literal(r["values_hex"][1])}']
        lines+=['End Sub','']
    lines.append('Private Sub Test_PublicFresh_Inverses()');names.append('Test_PublicFresh_Inverses')
    for i,r in enumerate(fresh['inverses']):
        lines += [f'    PF_Inverse "PF inverse {i:02d}", {literal(r["a"])}, {literal(r["p"])}, {literal(r["reference_hex"])}']
    lines+=['End Sub','']
    lines.append("""Private Sub PF_Inverse(ByVal Label As String, ByVal A As Double, ByVal Probability As Double, ByVal ExpectedUnit As Double)
    Dim Value As Variant, Diag As String
    Value = K_STATS_Gamma_InverseCumulative(Probability, A, 2#, Diag)
    PF_InverseValue Label & " Gamma", Value, 2# * ExpectedUnit, Diag
    Value = K_STATS_ChiSquare_InverseCumulative(Probability, 2# * A, Diag)
    PF_InverseValue Label & " Chi-square", Value, 2# * ExpectedUnit, Diag
End Sub

Private Sub PF_InverseValue(ByVal Label As String, ByVal Value As Variant, ByVal Expected As Double, ByVal Diag As String)
    If IsError(Value) Then
        AssertTrue Label & " returned error " & Diag, False
    Else
        AssertTrue Label & " bits " & SG_Bits(CDbl(Value)) & " normalized " & CStr(CDbl(Value) / Expected) & " " & Diag, _
            (Diag = "" And Abs(CDbl(Value) / Expected - 1#) <= 0.00000000000002)
    End If
End Sub
""")
    lines.append('Private Sub Test_PublicGamma_Reach()');names.append('Test_PublicGamma_Reach')
    for i,(a,x,p,q) in enumerate(REACH):
        lines += [f'    PG_Reach "PG reach {i}", {a}, {x}, {literal(p)}, {literal(q)}']
    lines+=['End Sub','']
    lines.append(f"""Private Sub PG_Reach(ByVal Label As String, ByVal A As Double, ByVal X As Double, _
    ByVal ExpectedP As Double, ByVal ExpectedQ As Double)
    Dim P As Double, Q As Double, OK As Boolean, Diag As String
    OK = PROB_TryGammaRegularizedP(A, X, P, Diag)
    AssertTrue Label & " P returned " & Diag, (OK And Diag = "")
    AssertTrue Label & " P relative " & CStr(P / ExpectedP - 1#), _
        (OK And Abs(P / ExpectedP - 1#) <= {REACH_LIMITS[0]})
    OK = PROB_TryGammaRegularizedQ(A, X, Q, Diag)
    AssertTrue Label & " Q returned " & Diag, (OK And Diag = "")
    AssertTrue Label & " Q relative " & CStr(Q / ExpectedQ - 1#), _
        (OK And Abs(Q / ExpectedQ - 1#) <= {REACH_LIMITS[1]})
    'The fallback forms the complemented tail as 1 - other (#24): its relative
    'error may not exceed the other tail's by more than rounding.
    If X < A + 1# Then
        AssertTrue Label & " complement Q amplification", _
            (Abs(Q / ExpectedQ - 1#) <= 1.001 * Abs(P / ExpectedP - 1#) + 0.000000000000001)
    Else
        AssertTrue Label & " complement P amplification", _
            (Abs(P / ExpectedP - 1#) <= 1.001 * Abs(Q / ExpectedQ - 1#) + 0.000000000000001)
    End If
End Sub
""")
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
    """Exact inverse of this generator; only for historical replay fixtures.
    The harness timing of #298 and the study split of #300 are not part of any
    frozen study: they go first."""
    text=studies.strip(harness.strip(text))
    path=module_path(text)
    if path in AMENDMENTS and path!=SRC:
        return amend(path,text,reverse=True) if amended(path,text) else text
    if BEGIN not in text:return text
    if 'Attribute VB_Name = "M_STATS_PROBDIST_SPECIALFUNCS"' in text:
        text=amend(SRC,text,reverse=True)
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
    text=amend(SRC,text)
    (ROOT/SRC).write_bytes(text.replace('\n','\r\n').encode('ascii'))
    for path in AMENDMENTS:
        if path==SRC:continue
        other=(ROOT/path).read_text()
        if other==original(path):other=amend(path,other)
        (ROOT/path).write_bytes(other.replace('\n','\r\n').encode('ascii'))
    text=restore((ROOT/TEST).read_text());text=change_controls(text)
    text=text.replace("' BEGIN generated wide-Gamma study",test_block()+"' BEGIN generated wide-Gamma study",1)
    text=text.replace('    Test_WideGamma_Fidelity\n','    Test_WideGamma_Fidelity\n    Test_PublicGamma_Integration\n',1)
    text=harness.apply(studies.apply(text))
    (ROOT/TEST).write_bytes(text.replace('\n','\r\n').encode('ascii'))
    p=ROOT/'.github/excel-evidence-policy.json';policy=json.loads(p.read_text())
    assert policy['expected_assertions'] in (10848,13182,10848+4030,10848+4038,10848+EXTRA,10848+EXTRA-studies.STUDY_ASSERTIONS)
    policy['expected_assertions']=10848+EXTRA-studies.STUDY_ASSERTIONS;p.write_text(json.dumps(policy,indent=2)+'\n')
    print(len(names),'arithmetic procedures;',EXTRA,'new assertions;',policy['expected_assertions'],'total')


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--write',action='store_true',required=True);parser.parse_args();install()
