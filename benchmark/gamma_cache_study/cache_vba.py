"""Exact one-entry DD LogGamma cache; reversible for historical source proofs."""
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SRC = ROOT / 'src/M_STATS_PROBDIST_SPECIALFUNCS.bas'


def frozen_source():
    """SPECIALFUNCS as committed at the Gamma study pin (#338), never the
    working tree: later production changes must not move a closed study."""
    import sys
    sys.path.insert(0, str(ROOT / 'benchmark'))
    from _frozen_sources import text
    return text('src/M_STATS_PROBDIST_SPECIALFUNCS.bas')
NAME = 'PROB_DD_EG_LogGamma'
SLOW = NAME + 'Uncached'
WRAPPER = '''Private Function PROB_DD_EG_LogGamma(ByVal A As Variant) As Variant
'==============================================================================
' PURPOSE
'   Reuse one successfully computed DD LogGamma tuple for an identical input.
' NOTES
'   Key uses both IEEE bit patterns and the integer exponent, not rounded A.
'   Scalar cached components are copied into a fresh array on every hit.
'   A failed miss leaves the previous successful entry intact.
' GENERATED
'   benchmark/gamma_cache_study/cache_vba.py (#307).
'==============================================================================
    Static Ready As Boolean, PreviousKey As String
    Static CachedH As Double, CachedL As Double, CachedE As Long
    Dim Key As String, Value As Variant
    Key = PROB_DD_SG_Bits(CDbl(A(0))) & PROB_DD_SG_Bits(CDbl(A(1))) & ":" & CStr(CLng(A(2)))
    If Ready Then
        If Key = PreviousKey Then
            PROB_DD_EG_LogGamma = Array(CachedH, CachedL, CachedE)
            Exit Function
        End If
    End If
    Value = PROB_DD_EG_LogGammaUncached(A)
    Ready = False
    CachedH = CDbl(Value(0)): CachedL = CDbl(Value(1)): CachedE = CLng(Value(2))
    PreviousKey = Key
    Ready = True
    PROB_DD_EG_LogGamma = Value
End Function

'''


def procedure(text, name):
    match = re.search(r'^Private Function ' + name + r'\(.*?^End Function\n', text, re.M | re.S)
    if match is None:
        raise ValueError('missing function: ' + name)
    return match.group(0)


def strip(text):
    if SLOW not in text:
        return text
    if text.count(WRAPPER) != 1:
        raise ValueError('DD cache wrapper drift')
    text = text.replace(WRAPPER, '', 1)
    slow = procedure(text, SLOW)
    original = re.sub(r'\b' + SLOW + r'\b', NAME, slow)
    return text.replace(slow, original, 1)


def apply(text):
    if SLOW in text:
        raise ValueError('DD cache already installed')
    original = procedure(text, NAME)
    slow = re.sub(r'\b' + NAME + r'\b', SLOW, original)
    result = text.replace(original, WRAPPER + slow, 1)
    if strip(result) != text:
        raise ValueError('cache reversal changed source')
    return result


if __name__ == '__main__':
    text = SRC.read_text()
    SRC.write_bytes(apply(text).replace('\n', '\r\n').encode('ascii'))
