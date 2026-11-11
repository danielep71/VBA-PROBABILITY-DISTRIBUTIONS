"""
The #307 binary64 Gamma route as it was generated into
src/M_STATS_PROBDIST_SPECIALFUNCS.bas, for the fit wave of FIT_DESIGN.md.

The VBA is written here, one arithmetic statement per line where the rounding
matters, so that route_model.py can mirror it statement by statement. The
exact Doubles it needs (log(k / 256) for k = 128 .. 256 as pairs, log 2 as a
pair, 2 pi) are written as their IEEE words, through PROB_DD_SP_FromWords of
the generated public Gamma block: a VBA decimal literal keeps only 15
significant digits.

History. #312 appended the block to the module at candidate 88ff6ba, where
Excel run 42798193976 measured it (FIT_RESULT.md). No public function ever
called it, and #307 closed without a routing rule, so the block was removed
from production again. It stays reproducible: --check proves that the block
generated here is byte for byte the one the fit candidate carried, and that
the current module no longer carries it, and that b64_route_block.txt, the
frozen copy that historical_block.py serves to the studies without mpmath, is
the same block. apply() and strip() remain for
historical replays; public_integration.py strips it before restoring the
frozen port.

Run:
    python3 b64_route_vba.py --check    (the historical block, absent from the module)
"""
import argparse
import os
import struct
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, HERE)
import route_model as RM  # noqa: E402

SRC = "src/M_STATS_PROBDIST_SPECIALFUNCS.bas"
# The fit candidate of #312 (FIT_DESIGN.md), the only source that carried the block.
FIT_CANDIDATE = "88ff6bae3cbcfb393eb9b6153c7c7cfbf3bf2fbc"
BEGIN = "' BEGIN generated binary64 Gamma route (#307)"
END = "' END generated binary64 Gamma route (#307)"
ANCHOR = "' END generated public Gamma integration\n\n"
PUBLIC = ("PROB_TryGammaB64SeriesP", "PROB_TryGammaB64ContinuedFractionQ")


class RouteBlockError(ValueError):
    """The module does not carry exactly the generated block."""


def words(value):
    high, low = struct.unpack(">II", struct.pack(">d", value))
    return f"PROB_DD_SP_FromWords(&H{high:08X}&, &H{low:08X}&)"


def header(name, purpose, notes=()):
    lines = ["'==============================================================================",
             f"' {name}", "'------------------------------------------------------------------------------",
             "' PURPOSE"]
    lines += [f"'   {p}" for p in purpose]
    if notes:
        lines.append("' NOTES")
        lines += [f"'   {n}" for n in notes]
    lines += ["' GENERATED", "'   benchmark/gamma_route_study/b64_route_vba.py (#307). Do not edit.",
              "'=============================================================================="]
    return "\n".join(lines)


def table_lines():
    out = []
    for k in range(128, 257):
        hi, lo = RM.LOG_TABLE[k - 128]
        out.append(f"        TableHi({k}) = {words(hi)}")
        out.append(f"        TableLo({k}) = {words(lo)}")
    return "\n".join(out)


def block():
    eft = f"""Private Function PROB_B64_Pow2(ByVal Biased As Long) As Double
{header("PROB_B64_Pow2", ["2 ^ (Biased - 1023), exactly, for a biased exponent in 1 .. 2046."])}
    PROB_B64_Pow2 = PROB_DD_SP_FromWords(Biased * &H100000, 0)
End Function

Private Function PROB_B64_Field(ByVal Y As Double) As Long
{header("PROB_B64_Field", ["The biased exponent field of a Double: 0 for zero or a subnormal."])}
    Dim N As PROB_DD_Number, W As PROB_DD_Words
    N.Value = Y
    LSet W = N
    PROB_B64_Field = (W.High And &H7FF00000) \\ &H100000
End Function

Private Sub PROB_B64_TwoSum(ByVal A As Double, ByVal B As Double, ByRef S As Double, ByRef E As Double)
{header("PROB_B64_TwoSum", ["S + E = A + B exactly, S the rounded sum (Knuth)."])}
    Dim BB As Double, T As Double, U As Double
    S = A + B
    BB = S - A
    T = S - BB
    T = A - T
    U = B - BB
    E = T + U
End Sub

Private Sub PROB_B64_FastTwoSum(ByVal A As Double, ByVal B As Double, ByRef S As Double, ByRef E As Double)
{header("PROB_B64_FastTwoSum", ["S + E = A + B exactly for |A| >= |B| (Dekker)."])}
    Dim T As Double
    S = A + B
    T = S - A
    E = B - T
End Sub

Private Sub PROB_B64_Split(ByVal A As Double, ByRef Hi As Double, ByRef Lo As Double)
{header("PROB_B64_Split", ["Hi + Lo = A, each with at most 26 significant bits (Veltkamp)."])}
    Const SPLITTER As Double = 134217729#
    Dim C As Double, T As Double
    C = SPLITTER * A
    T = C - A
    Hi = C - T
    Lo = A - Hi
End Sub

Private Sub PROB_B64_TwoProd(ByVal A As Double, ByVal B As Double, ByRef P As Double, ByRef E As Double)
{header("PROB_B64_TwoProd", ["P + E = A * B exactly, P the rounded product (Dekker)."])}
    Dim AH As Double, AL As Double, BH As Double, BL As Double, T As Double
    P = A * B
    PROB_B64_Split A, AH, AL
    PROB_B64_Split B, BH, BL
    E = AH * BH
    E = E - P
    T = AH * BL
    E = E + T
    T = AL * BH
    E = E + T
    T = AL * BL
    E = E + T
End Sub

Private Sub PROB_B64_PairAdd(ByVal AH As Double, ByVal AL As Double, ByVal BH As Double, ByVal BL As Double, _
    ByRef SH As Double, ByRef SL As Double)
{header("PROB_B64_PairAdd", ["The sum of two unevaluated pairs, as a pair."])}
    Dim S As Double, E As Double, T As Double
    PROB_B64_TwoSum AH, BH, S, E
    T = AL + BL
    E = E + T
    PROB_B64_FastTwoSum S, E, SH, SL
End Sub

Private Sub PROB_B64_PairMul(ByVal AH As Double, ByVal AL As Double, ByVal B As Double, _
    ByRef PH As Double, ByRef PL As Double)
{header("PROB_B64_PairMul", ["The product of a pair and a Double, as a pair."])}
    Dim P As Double, E As Double, T As Double
    PROB_B64_TwoProd AH, B, P, E
    T = AL * B
    E = E + T
    PROB_B64_FastTwoSum P, E, PH, PL
End Sub

Private Sub PROB_B64_PairDiv(ByVal AH As Double, ByVal AL As Double, ByVal B As Double, _
    ByRef QH As Double, ByRef QL As Double)
{header("PROB_B64_PairDiv", ["The quotient of a pair by a Double, as a pair."])}
    Dim Q As Double, P As Double, E As Double, R As Double
    Q = AH / B
    PROB_B64_TwoProd Q, B, P, E
    R = AH - P
    R = R - E
    R = R + AL
    R = R / B
    PROB_B64_FastTwoSum Q, R, QH, QL
End Sub
"""
    ln2h, ln2l = RM.LN2
    log = f"""
Private Sub PROB_B64_PairLog(ByVal YH As Double, ByVal YL As Double, ByRef LH As Double, ByRef LL As Double)
{header("PROB_B64_PairLog", ["log(YH + YL) as a pair, for a positive pair whose high part is normal.",
                             "Y = M 2^E with 0.5 <= M < 1; C = K / 256 nearest to M; then",
                             "log Y = E log 2 + log C + log(1 + U), U = (M - C) / C, |U| <= 2^-8,",
                             "log(1 + U) = U - U^2 / 2 + U^3 S(U), S(U) = sum (-U)^J / (J + 3), J = 0 .. 7."],
                            ["log(K / 256) and log 2 are exact pairs; the table loads on the first call."])}
    Static TableHi(128 To 256) As Double, TableLo(128 To 256) As Double, S(0 To 7) As Double
    Static Ln2Hi As Double, Ln2Lo As Double, Ready As Boolean
    Dim Field As Long, E As Long, K As Long, J As Long
    Dim Scale2 As Double, M As Double, C As Double, VH As Double, VL As Double, WH As Double, WL As Double
    Dim UH As Double, UL As Double, SH As Double, SL As Double, T As Double, Poly As Double, Cube As Double
    Dim AH As Double, AL As Double, EH As Double, EL As Double, Sign As Double
    If Not Ready Then
{table_lines()}
        Ln2Hi = {words(ln2h)}
        Ln2Lo = {words(ln2l)}
        Sign = 1#
        For J = 0 To 7
            S(J) = Sign / CDbl(J + 3)
            Sign = -Sign
        Next J
        Ready = True
    End If
    Field = PROB_B64_Field(YH)
    E = Field - 1022
    Scale2 = PROB_B64_Pow2(2045 - Field)
    M = YH * Scale2
    K = CLng(Int(M * 256# + 0.5))
    C = CDbl(K) / 256#
    VH = M - C
    VL = YL * Scale2
    PROB_B64_TwoSum VH, VL, WH, WL
    PROB_B64_PairDiv WH, WL, C, UH, UL
    PROB_B64_TwoProd UH, UH, SH, SL
    T = 2# * UH
    T = T * UL
    SL = SL + T
    SH = SH * 0.5
    SL = SL * 0.5
    Poly = S(7)
    For J = 6 To 0 Step -1
        Poly = Poly * UH
        Poly = Poly + S(J)
    Next J
    Cube = UH * UH
    Cube = Cube * UH
    Cube = Cube * Poly
    PROB_B64_PairAdd UH, UL, -SH, -SL, AH, AL
    PROB_B64_PairAdd AH, AL, Cube, 0#, AH, AL
    PROB_B64_PairAdd AH, AL, TableHi(K), TableLo(K), AH, AL
    PROB_B64_PairMul Ln2Hi, Ln2Lo, CDbl(E), EH, EL
    PROB_B64_PairAdd AH, AL, EH, EL, LH, LL
End Sub
"""
    factor = f"""
Private Sub PROB_B64_Deviance(ByVal A As Double, ByVal X As Double, ByRef DH As Double, ByRef DL As Double)
{header("PROB_B64_Deviance", ["Loader's deviance D = A log(A/X) + X - A = A (T - log(1 + T)),",
                              "T = X/A - 1, as a pair accurate to about 1E-17 absolute."],
                             ["Raises error 5 when X / A has a biased exponent field outside",
                              "1 .. {RM.EXPONENT_FIELD[1]}; the calling route then declines."])}
    Dim Q As Double, YH As Double, YL As Double, TH As Double, TL As Double
    Dim LH As Double, LL As Double, GH As Double, GL As Double, Field As Long
    Q = X / A
    Field = PROB_B64_Field(Q)
    If Field < {RM.EXPONENT_FIELD[0]} Or Field > {RM.EXPONENT_FIELD[1]} Then Err.Raise 5
    PROB_B64_PairDiv X, 0#, A, YH, YL
    PROB_B64_PairAdd YH, YL, -1#, 0#, TH, TL
    PROB_B64_PairLog YH, YL, LH, LL
    PROB_B64_PairAdd TH, TL, -LH, -LL, GH, GL
    PROB_B64_PairMul GH, GL, A, DH, DL
End Sub

Private Function PROB_B64_Factor(ByVal A As Double, ByVal X As Double) As Double
{header("PROB_B64_Factor", ["x^A e^-x / Gamma(A) for A > 0 and X > 0.",
                            "A > 1: sqrt(A / (2 pi)) exp(-(D + stirlerr(A))), with the exponent",
                            "kept as a pair, exp(EH) (1 + EL).",
                            "A <= 1: A e^-X exp(A log X - log Gamma(1 + A)), as the small-shape",
                            "factor of PROB_TryGammaContinuedFractionQ."],
                           ["Returns 0 where the factor is below about 1E-307 (X >= 708 at A <= 1,",
                            "EH < -708 above): every probability it scales is then below 1E-300."])}
    Dim F As Double, LX As Double, EX As Double, Corr As Double
    Dim DH As Double, DL As Double, Stir As Double, EH As Double, EL As Double
    Dim Root As Double, T As Double
    If A <= 1# Then
        If X >= 708# Then
            PROB_B64_Factor = 0#
            Exit Function
        End If
        F = Exp(-X)
        LX = Log(X)
        EX = A * LX - PROB_GammaSmallLogGamma1p(A)
        Corr = Exp(EX)
        F = A * F
        PROB_B64_Factor = F * Corr
        Exit Function
    End If
    PROB_B64_Deviance A, X, DH, DL
    Stir = PROB_StirlingError(A)
    PROB_B64_TwoSum -DH, -Stir, EH, EL
    If EH < -708# Then
        PROB_B64_Factor = 0#
        Exit Function
    End If
    EL = EL - DL
    Root = A / {words(RM.TWO_PI)}
    Root = Sqr(Root)
    F = Exp(EH)
    F = Root * F
    T = F * EL
    PROB_B64_Factor = F + T
End Function
"""
    series = f"""
Public Function PROB_TryGammaB64SeriesP( _
    ByVal A As Double, _
    ByVal X As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
{header("PROB_TryGammaB64SeriesP",
        ["P(A, X) for 0 < X < A + 1 by the lower series in binary64 (#307): the",
         "recurrence of PROB_TryGammaSeriesP with a compensated (Kahan) sum, times",
         "PROB_B64_Factor. It stops once the term is below an ulp of the sum and the",
         "remaining terms, which fall at least geometrically by X / (Ap + 1), sum to",
         "at most half an ulp."],
        ["False (declined) clears Result: invalid arguments, the 100000-iteration",
         "cap, a deviance outside its range or any runtime error. The caller then",
         "answers with the double-double kernel."])}
    Dim Ap As Double, SumValue As Double, Del As Double, Carry As Double
    Dim S As Double, T As Double, Room As Double, Bound As Double, Limit As Double
    Dim Eps As Double, HalfEps As Double, IterIdx As Long
    Result = 0#: FailMsg = ""
    On Error GoTo Declined
    If Not (A > 0# And X > 0# And X < A + 1#) Then GoTo Declined
    Eps = PROB_B64_Pow2(971)
    HalfEps = PROB_B64_Pow2(970)
    Ap = A
    SumValue = 1# / A
    Del = SumValue
    Carry = 0#
    For IterIdx = 1 To PROB_GAMMA_MAX_ITER
        Ap = Ap + 1#
        Del = Del * X / Ap
        S = SumValue + Del
        T = S - SumValue
        T = Del - T
        Carry = Carry + T
        SumValue = S
        If Abs(Del) <= Abs(SumValue) * Eps Then
            Room = Ap + 1#
            Room = Room - X
            If Room > 0# Then
                Bound = Del * X
                Limit = SumValue * HalfEps
                Limit = Limit * Room
                If Bound <= Limit Then
                    SumValue = SumValue + Carry
                    Result = SumValue * PROB_B64_Factor(A, X)
                    PROB_TryGammaB64SeriesP = True
                    Exit Function
                End If
            End If
        End If
    Next IterIdx
Declined:
    Result = 0#
    FailMsg = "binary64 Gamma series declined"
End Function
"""
    cf = f"""
Public Function PROB_TryGammaB64ContinuedFractionQ( _
    ByVal A As Double, _
    ByVal X As Double, _
    ByRef Result As Double, _
    ByRef FailMsg As String) _
    As Boolean
{header("PROB_TryGammaB64ContinuedFractionQ",
        ["Q(A, X) for X >= A + 1 by the continued fraction in binary64 (#307): the",
         "modified Lentz recurrence of PROB_TryGammaContinuedFractionQ stopped at",
         "machine epsilon at every shape, with its backward re-evaluation at shape 1",
         "and below, times PROB_B64_Factor."],
        ["False (declined) clears Result: invalid arguments, the 100000-iteration",
         "cap, a deviance outside its range or any runtime error. The caller then",
         "answers with the double-double kernel."])}
    Dim B As Double, C As Double, D As Double, H As Double, An As Double, Del As Double
    Dim Eps As Double, Base As Double, Tail As Double, IterIdx As Long, BackIdx As Long
    Result = 0#: FailMsg = ""
    On Error GoTo Declined
    If Not (A > 0# And X >= A + 1#) Then GoTo Declined
    Eps = PROB_B64_Pow2(971)
    B = X + 1# - A
    If Abs(B) < PROB_FPMIN Then B = PROB_FPMIN
    C = 1# / PROB_FPMIN
    D = 1# / B
    H = D
    For IterIdx = 1 To PROB_GAMMA_MAX_ITER
        An = -CDbl(IterIdx) * (CDbl(IterIdx) - A)
        B = B + 2#
        D = An * D + B
        If Abs(D) < PROB_FPMIN Then D = PROB_FPMIN
        C = B + An / C
        If Abs(C) < PROB_FPMIN Then C = PROB_FPMIN
        D = 1# / D
        Del = D * C
        H = H * Del
        If Abs(Del - 1#) <= Eps Then
            If A <= 1# Then
                Base = X + 1# - A
                Tail = 0#
                For BackIdx = IterIdx To 1 Step -1
                    Tail = (-CDbl(BackIdx) * (CDbl(BackIdx) - A)) / (Base + 2# * CDbl(BackIdx) + Tail)
                Next BackIdx
                Tail = Base + Tail
                H = 1# / Tail
            End If
            Result = PROB_B64_Factor(A, X) * H
            PROB_TryGammaB64ContinuedFractionQ = True
            Exit Function
        End If
    Next IterIdx
Declined:
    Result = 0#
    FailMsg = "binary64 Gamma continued fraction declined"
End Function
"""
    return BEGIN + "\n" + eft + log + factor + series + cf + END + "\n\n"


def apply(text):
    """Append the block to the module's LF text (after the public Gamma block)."""
    if BEGIN in text or END in text:
        raise RouteBlockError("the binary64 Gamma route is already installed")
    if not text.endswith(ANCHOR):
        raise RouteBlockError("the module does not end with the public Gamma block")
    return text + block()


def strip(text):
    """Remove the block exactly; refuse a hand-edited or partial block."""
    if BEGIN not in text and END not in text:
        return text
    current = block()
    if not text.endswith(current) or text.count(BEGIN) != 1:
        raise RouteBlockError("binary64 Gamma route block drift")
    return text[: -len(current)]


def as_measured(text):
    """See historical_block.as_measured (standard library only)."""
    import historical_block
    return historical_block.as_measured(text)


def read(path, revision=None):
    if revision:
        raw = subprocess.check_output(["git", "show", f"{revision}:{path}"], cwd=ROOT)
    else:
        with open(os.path.join(ROOT, path), "rb") as f:
            raw = f.read()
    return raw.decode("ascii").replace("\r\n", "\n")


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--check", action="store_true", required=True)
    ap.parse_args(argv)
    historical = read(SRC, FIT_CANDIDATE)
    current = read(SRC)
    carried = historical.endswith(block()) and historical.count(BEGIN) == 1
    with open(os.path.join(HERE, "b64_route_block.txt"), encoding="ascii") as f:
        frozen = f.read() == block()
    absent = BEGIN not in current and END not in current
    print("binary64 Gamma route block: "
          + ("generated as at the fit candidate" if carried else "DIFFERS FROM THE FIT CANDIDATE")
          + "; " + ("frozen copy current" if frozen else "FROZEN COPY DIFFERS")
          + "; " + ("absent from the module" if absent else "STILL IN THE MODULE"))
    return 0 if carried and frozen and absent else 1


if __name__ == "__main__":
    sys.exit(main())
