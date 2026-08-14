Attribute VB_Name = "M_STATS_PROBDIST_GQPROBE"
Option Explicit

'Column count of the output header, checked against every row written
Private Const PROBE_FIELDS As Long = 12

'==============================================================================
' M_STATS_PROBDIST_GQPROBE
'------------------------------------------------------------------------------
' PURPOSE
'   Scopes a Gamma survival accuracy defect that is NOT part of #13.
'
' WHY IT IS NOT #13
'   Every point here uses ScaleParam = 1, so the standardization
'   X / ScaleParam is a no-op and the stored standardized argument is
'   bit-identical to the exact one. Transform relative error is exactly zero at
'   every row. Whatever this measures, it cannot be positive-ratio information
'   loss.
'
'   The defect was found in the #13 survival evidence, where landmark points
'   carry no transform error by construction and still showed the current path
'   at 8.25E-13 against a 5E-15 contract. That was provenance, not the regime:
'   the argument was subnormal only because the study happened to be looking
'   there.
'
' THE MECHANISM UNDER TEST
'   PROB_TryGammaRegularizedQ branches on X < A + 1. Below that it computes the
'   lower tail by series and returns the complement:
'
'       PROB_TryGammaSeriesP(A, X, Value)
'       Value = 1# - Value
'
'   Forming Q as 1 - P amplifies P's relative error by P / Q. As the shape
'   falls, P approaches 1 and the factor grows without bound: it is 2.15 at
'   shape 0.5 and 1.79E+06 at shape 1E-6, both at X = 0.5.
'
'   This is the same complement cancellation that #13 already specifies against
'   for its own survival branch - Q = -Expm1(LogP), never 1 - Exp(LogP).
'
' WHAT THE PROBE MUST DISTINGUISH
'   Two scopes with different remediations:
'
'     A  the series is accurate and only the complement is at fault
'        -> CDF within its 3E-15 contract, survival outside 5E-15,
'           and Q_error is about (P / Q) * P_error
'
'     B  the series itself is outside contract at tiny shape
'        -> CDF ALSO outside 3E-15, survival worse by the amplification
'
'   In case B, replacing 1 - P alone would fix the visible survival error and
'   leave a contractual CDF defect behind. That is why the cumulative is
'   measured here rather than assumed sound.
'
' THE ACCOUNTING IDENTITY
'   PROB_TryGammaSeriesP is Public, so its raw output is recorded rather than
'   inferred. With Q = 1 - P exactly,
'
'       relerr(Q) = (P / Q) * relerr(P)
'
'   If the measured survival error matches that prediction from the measured
'   series error, the mechanism is demonstrated rather than argued from source.
'
' CONTROLS
'   X = 2 takes the continued-fraction branch at every shape here, so it does
'   not form the complement at all. If those rows stay clean while X = 0.5
'   fails, the branch localisation is direct.
'
'   X = 4.144523E-317 reproduces the point that exposed the defect.
'
'   Shapes at or above 0.1 with a subnormal argument are omitted: Q saturates
'   to exactly 1 there and the observable carries no information.
'
' OUTPUT
'   point_id, x, shape, series_p_status, series_p, cdf_status, cdf,
'   survival_status, survival, cand_p, cand_q, branch
'
'   Numeric values are hi;lo tokens. branch is series or contfrac, computed the
'   same way the kernel decides, so a disagreement with the recorded values is
'   itself visible.
'
' SCOPE
'   Characterization only. Promotes no grid row, claims no threshold, touches
'   no registry, and changes no source.
'
' DEPENDENCIES
'   - K_STATS_Gamma_Cumulative, K_STATS_Gamma_Survival (CONTINUOUS)
'   - PROB_TryGammaSeriesP                             (SPECIALFUNCS)
'   - PROB_TryLogGamma1p, PROB_TryExp, PROB_Expm1      (SPECIALFUNCS, CORE)
'
' UPDATED
'   2026-08-14
'==============================================================================


Public Sub Probe_GammaQComplement()
'
'==============================================================================
' Probe_GammaQComplement
'------------------------------------------------------------------------------
' PURPOSE
'   Evaluates the shape ladder at each control argument and writes the CSV.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'Chosen output path
    Dim OutNum              As Integer         'Output file handle
    Dim Shapes              As Variant         'Shape ladder
    Dim Args                As Variant         'Control arguments
    Dim Labels              As Variant         'Argument labels
    Dim Row                 As String          'Assembled CSV row
    Dim Rows                As Long            'Rows written
    Dim I                   As Long            'Argument index
    Dim J                   As Long            'Shape index
    Dim X                   As Double          'Current argument
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

        Path = ResolveOutputPath()
        If Len(Path) = 0 Then Exit Sub          'User cancelled the picker

    'Shape-major: the amplification is a function of shape, so this is the axis
    'the defect turns on. X is a control, not the regime
        Shapes = Array(0.000001, 0.000003, 0.00001, 0.00003, _
                       0.0001, 0.0003, 0.001, 0.003, _
                       0.01, 0.03, 0.1, 0.5)

    'X = 0.5 is the primary slice and takes the series branch at every shape.
    'X = 2 takes the continued fraction at every shape and is the branch
    'control. The subnormal value is the discovery point, kept as provenance
        Labels = Array("x05", "x2", "xsub")
        Args = Array(0.5, 2#, 4.144523E-317)
'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        OutNum = FreeFile
        Open Path For Output As #OutNum

        Print #OutNum, "point_id,x,shape,series_p_status,series_p," & _
                       "cdf_status,cdf,survival_status,survival," & _
                       "cand_p,cand_q,branch"

        Rows = 0
        For I = LBound(Args) To UBound(Args)
            X = CDbl(Args(I))
            For J = LBound(Shapes) To UBound(Shapes)

    'Above shape 0.1 the survival saturates to exactly 1 at a subnormal
    'argument, so those cells carry no information and are skipped
                If Not (I = 2 And CDbl(Shapes(J)) >= 0.1) Then

                    Row = EvaluatePoint(CStr(Labels(I)) & "_s" & J, _
                                        X, CDbl(Shapes(J)))

                    If FieldCount(Row) <> PROBE_FIELDS Then
                        Err.Raise 5, , "Row has " & FieldCount(Row) & _
                                       " fields, expected " & PROBE_FIELDS
                    End If

                    Print #OutNum, Row
                    Rows = Rows + 1
                End If
            Next J
        Next I

        Close #OutNum

    MsgBox "Gamma Q complement probe complete: " & Rows & _
           " rows written to" & vbCrLf & Path, _
           vbInformation, "Gamma Q probe"
    Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    On Error Resume Next
    Close #OutNum
    MsgBox "Gamma Q complement probe failed: " & Err.Description, _
           vbExclamation, "Gamma Q probe"
End Sub


Private Function EvaluatePoint( _
    ByVal Id As String, _
    ByVal X As Double, _
    ByVal Shape As Double) _
    As String
'
'==============================================================================
' EvaluatePoint
'------------------------------------------------------------------------------
' PURPOSE
'   Records the raw series output, both public surfaces, and the log-path
'   mirror at one point.
'
' WHY THE SERIES IS RECORDED
'   It closes the accounting identity. Q = 1 - P exactly, so relerr(Q) is
'   (P / Q) * relerr(P); measuring P directly turns the mechanism from an
'   inference about the source into a demonstrated one.
'
' SCALE IS ALWAYS ONE
'   That is the point of the probe: it puts every row provably outside #13,
'   because the standardization cannot lose anything when it is a no-op.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Row                 As String          'Assembled CSV row
    Dim SeriesP             As Double          'Raw PROB_TryGammaSeriesP output
    Dim SeriesOk            As Boolean         'Series returned a value
    Dim FailMsg             As String          'Kernel diagnostic
    Dim Cdf                 As Variant         'K_STATS_Gamma_Cumulative
    Dim Surv                As Variant         'K_STATS_Gamma_Survival
    Dim CandP               As Double          'Log-path lower tail
    Dim CandQ               As Double          'Log-path survival
    Dim CandOk              As Boolean         'Mirror produced values
    Dim ErrNum              As Long            'Captured error number
'------------------------------------------------------------------------------
' HEADER FIELDS
'------------------------------------------------------------------------------
        Row = Id & "," & FormatFullPrecision(X) & "," & _
              FormatFullPrecision(Shape)
'------------------------------------------------------------------------------
' RAW SERIES
'------------------------------------------------------------------------------
    'Only meaningful on the series side of the branch; recorded regardless so
    'the branch column can be checked against what was actually produced
        On Error Resume Next
        Err.Clear
        SeriesOk = PROB_TryGammaSeriesP(Shape, X, SeriesP, FailMsg)
        ErrNum = Err.Number
        Err.Clear
        On Error GoTo 0

        If ErrNum <> 0 Or Not SeriesOk Then
            Row = Row & ",ERROR,"
        Else
            Row = Row & ",OK," & FormatFullPrecision(SeriesP)
        End If
'------------------------------------------------------------------------------
' PUBLIC SURFACES
'------------------------------------------------------------------------------
        On Error Resume Next
        Err.Clear
        Cdf = K_STATS_Gamma_Cumulative(X, Shape, 1#)
        ErrNum = Err.Number
        Err.Clear
        On Error GoTo 0
        Row = Row & SurfaceFields(ErrNum, Cdf)

        On Error Resume Next
        Err.Clear
        Surv = K_STATS_Gamma_Survival(X, Shape, 1#)
        ErrNum = Err.Number
        Err.Clear
        On Error GoTo 0
        Row = Row & SurfaceFields(ErrNum, Surv)
'------------------------------------------------------------------------------
' LOG-PATH MIRROR
'------------------------------------------------------------------------------
        CandOk = TryCandidate(X, Shape, CandP, CandQ)

        If CandOk Then
            Row = Row & "," & FormatFullPrecision(CandP) & _
                        "," & FormatFullPrecision(CandQ)
        Else
            Row = Row & ",,"
        End If
'------------------------------------------------------------------------------
' BRANCH
'------------------------------------------------------------------------------
    'Computed exactly as PROB_TryGammaRegularizedQ decides, so a row whose
    'values disagree with its declared branch is visible rather than assumed
        Row = Row & "," & IIf(X < Shape + 1#, "series", "contfrac")

    EvaluatePoint = Row
End Function


Private Function SurfaceFields( _
    ByVal ErrNum As Long, _
    ByVal Value As Variant) _
    As String
'
'==============================================================================
' SurfaceFields
'------------------------------------------------------------------------------
' PURPOSE
'   Renders one public result as a status and value pair, so every branch
'   contributes the same field count.
'==============================================================================
'
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    If ErrNum <> 0 Then
        SurfaceFields = ",ERROR,"
    ElseIf IsError(Value) Then
        SurfaceFields = ",CVERR,"
    Else
        SurfaceFields = ",OK," & FormatFullPrecision(CDbl(Value))
    End If
End Function


Private Function TryCandidate( _
    ByVal X As Double, _
    ByVal Shape As Double, _
    ByRef P As Double, _
    ByRef Q As Double) _
    As Boolean
'
'==============================================================================
' TryCandidate
'------------------------------------------------------------------------------
' PURPOSE
'   Mirrors the log-domain arrangement #13 specifies, as a measurement
'   instrument only. It touches no production source.
'
'   The survival is taken as -Expm1(Log P) rather than 1 - Exp(Log P). That is
'   precisely the arrangement the current kernel does not use, so the pair of
'   columns isolates the cost of the complement.
'
'   Valid only where the leading series term dominates, so it is informative at
'   tiny shape and small argument and not elsewhere. The analyzer decides which
'   rows it applies to; this function does not guess.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LogGamma1pShape     As Double          'LogGamma(1 + Shape)
    Dim LogP                As Double          'Log of the lower tail
    Dim FailMsg             As String          'Kernel diagnostic
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

        If X <= 0# Then Exit Function
        If Not PROB_TryLogGamma1p(Shape, LogGamma1pShape, FailMsg) Then
            Exit Function
        End If

        LogP = Shape * Log(X) - LogGamma1pShape
        If Not PROB_TryExp(LogP, P) Then Exit Function

    'Never 1 - Exp(LogP): that is the cancellation under measurement
        Q = -PROB_Expm1(LogP)

        TryCandidate = True
    Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    Err.Clear
End Function


Private Function FieldCount(ByVal Text As String) As Long
'
'==============================================================================
' FieldCount
'------------------------------------------------------------------------------
' PURPOSE
'   Counts the comma-separated fields in an assembled row. Safe because
'   FormatFullPrecision separates its parts with a semicolon.
'==============================================================================
'
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    FieldCount = UBound(Split(Text, ",")) + 1
End Function


Private Function ResolveOutputPath() As String
'
'==============================================================================
' ResolveOutputPath
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a usable LOCAL path for the probe CSV, or an empty string if the
'   user cancels. ThisWorkbook.Path returns a URL under OneDrive, which Open
'   cannot write to and Dir$() cannot test.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Picked              As Variant         'File-dialog result
'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    Picked = Application.GetSaveAsFilename( _
        InitialFileName:="gamma_q_complement.csv", _
        FileFilter:="Probe output (*.csv),*.csv", _
        Title:="Save the Gamma Q complement probe")

    If VarType(Picked) = vbBoolean Then
        ResolveOutputPath = vbNullString
    Else
        ResolveOutputPath = CStr(Picked)
    End If
End Function


Private Function FormatFullPrecision(ByVal X As Double) As String
'
'==============================================================================
' FormatFullPrecision
'------------------------------------------------------------------------------
' PURPOSE
'   Renders a Double as "hi;lo", so the two parts summed in Double precision
'   reproduce the original bit pattern. A 15-digit decimal alone cannot
'   round-trip a subnormal, and one of the control arguments is subnormal.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HiStr               As String          'Value to 15 significant digits
    Dim Hi                  As Double          'The Double that HiStr denotes
    Dim Lo                  As Double          'Exact residual X - Hi
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    If X = 0# Then
        FormatFullPrecision = "0E+000;0E+000"
        Exit Function
    End If

    HiStr = Fmt15(X)
    Hi = Val(HiStr)                            'Val is locale-independent
    Lo = X - Hi

    FormatFullPrecision = HiStr & ";" & Fmt15(Lo)
End Function


Private Function Fmt15(ByVal X As Double) As String
'
'==============================================================================
' Fmt15
'------------------------------------------------------------------------------
' PURPOSE
'   Formats X to 15 significant digits in scientific notation with a US decimal
'   point, whatever the local settings are.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim S                   As String          'Formatted value
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    If X = 0# Then
        Fmt15 = "0E+000"
        Exit Function
    End If

    S = Format$(X, "0.00000000000000E+000")    '1 + 14 = 15 significant digits
    Fmt15 = Replace(S, ",", ".")               'Force US decimal, any locale
End Function
