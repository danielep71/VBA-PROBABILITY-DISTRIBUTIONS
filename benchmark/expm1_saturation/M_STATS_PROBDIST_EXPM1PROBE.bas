Attribute VB_Name = "M_STATS_PROBDIST_EXPM1PROBE"
Option Explicit

'Column count of the output header, checked against every row written
Private Const PROBE_FIELDS As Long = 9

'==============================================================================
' M_STATS_PROBDIST_EXPM1PROBE
'------------------------------------------------------------------------------
' PURPOSE
'   Characterization probe for the PROB_Expm1 subnormal saturation defect.
'
'   PROB_Expm1 rescales U - 1 by X / Log(U) with U = Exp(X). That ratio is
'   exactly 1 in exact arithmetic, and stays within an ULP of 1 while U
'   carries a full 53-bit significand. Below X of about -708.4 the
'   exponential is subnormal and carries fewer bits, so Log(U) no longer
'   recovers X and the ratio drifts. Because U - 1 has already rounded to
'   exactly -1 by then, the drift lands undiluted and the kernel returns
'   values BELOW -1, outside the range of Exp(X) - 1.
'
'   Two public CDFs propagate that directly as a probability above one:
'   K_STATS_Exponential_Cumulative and K_STATS_Weibull_Cumulative both compute
'   -PROB_Expm1(-z) and neither clamps. PROB_DS_TryGeometricCDF reaches that
'   kernel but clamps to [0, 1] afterwards, so it is protected by accident.
'
' RUN THIS TWICE
'   Once against the ORIGINAL kernel and once against the REPAIRED one, saving
'   to different files. The pre-fix run is the evidence that the defect was
'   real and reached the public surface; the post-fix run is the evidence that
'   it no longer does and that ordinary arguments did not move. Neither run is
'   worth much without the other.
'
'   The original kernel is the one at commit a1de93f: it branches on U = 0#
'   before the rescale. The repaired kernel branches on U - 1# = -1#.
'
' WHAT TO LOOK FOR
'   Pre-fix, expm1 falls below -1 from about X = -716 downward, reaching
'   -1.000926774504277 at X = -745.13, and both CDF columns exceed 1 over the
'   same span. Post-fix, expm1 is exactly -1 across the whole window and both
'   CDFs are exactly 1. The benign points from -0.025 to -100 must be
'   BIT-IDENTICAL between the two runs: the repair changes the branch condition
'   over a far wider span than the defect, so proving it changes no value
'   outside the window is the point of including them.
'
' OUTPUT (one row per point, header first)
'   point_id, x, exp_x, u_minus_1, expm1,
'   exp_cdf_status, exp_cdf, weibull_cdf_status, weibull_cdf
'
'   Numeric values are full-precision hi;lo tokens; sum the two parts in Double
'   precision to recover the original. The CDF arguments are the NEGATION of x,
'   since both CDFs evaluate -PROB_Expm1(-z).
'
' SCOPE
'   Characterization only. Promotes no grid row, claims no threshold and
'   touches no registry.
'
' DEPENDENCIES
'   - PROB_Expm1                    (M_STATS_PROBDIST_CORE)
'   - K_STATS_Exponential_Cumulative (M_STATS_PROBDIST_CONTINUOUS)
'   - K_STATS_Weibull_Cumulative     (M_STATS_PROBDIST_CONTINUOUS)
'
' UPDATED
'   2026-08-13
'==============================================================================


Public Sub Probe_Expm1Saturation()
'
'==============================================================================
' Probe_Expm1Saturation
'------------------------------------------------------------------------------
' PURPOSE
'   Evaluates every probe point and writes the characterization CSV.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'Chosen output path
    Dim OutNum              As Integer         'Output file handle
    Dim Ids()               As String          'Point identifiers
    Dim Xs()                As Double          'The points themselves
    Dim Count               As Long            'Number of points
    Dim Row                 As String          'Assembled CSV row
    Dim I                   As Long            'Point index
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

        Path = ResolveOutputPath()
        If Len(Path) = 0 Then Exit Sub          'User cancelled the picker

        BuildPoints Ids, Xs, Count
'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        OutNum = FreeFile
        Open Path For Output As #OutNum

        Print #OutNum, "point_id,x,exp_x,u_minus_1,expm1," & _
                       "exp_cdf_status,exp_cdf," & _
                       "weibull_cdf_status,weibull_cdf"

        For I = 0 To Count - 1
            Row = EvaluatePoint(Ids(I), Xs(I))

    'A row that does not match the header is silently unreadable: refuse to
    'write it rather than emit a file whose columns cannot be trusted
            If FieldCount(Row) <> PROBE_FIELDS Then
                Err.Raise 5, , "Row " & Ids(I) & " has " & FieldCount(Row) & _
                               " fields, expected " & PROBE_FIELDS
            End If

            Print #OutNum, Row
        Next I

        Close #OutNum

    MsgBox "Expm1 saturation probe complete: " & Count & _
           " point(s) written to" & vbCrLf & Path, _
           vbInformation, "Expm1 probe"
    Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    On Error Resume Next
    Close #OutNum
    MsgBox "Expm1 saturation probe failed: " & Err.Description, _
           vbExclamation, "Expm1 probe"
End Sub


Private Sub BuildPoints( _
    ByRef Ids() As String, _
    ByRef Xs() As Double, _
    ByRef Count As Long)
'
'==============================================================================
' BuildPoints
'------------------------------------------------------------------------------
' PURPOSE
'   Constructs the probe point set: ordinary negative arguments that must not
'   move, the error-growth landmarks through the subnormal window, and the two
'   edges where the window opens and closes.
'
' WHY THESE POINTS
'   The landmarks are where the pre-fix relative error crosses each order of
'   magnitude, so the growth of the defect is visible rather than only its
'   peak. All are normal-range decimals, so unlike the Stirling probe there is
'   no need to build them by arithmetic.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Cap                 As Long            'Array capacity
'------------------------------------------------------------------------------
' ALLOCATE
'------------------------------------------------------------------------------
    Cap = 32
    ReDim Ids(0 To Cap - 1)
    ReDim Xs(0 To Cap - 1)
    Count = 0
'------------------------------------------------------------------------------
' BENIGN POINTS - MUST BE BIT-IDENTICAL BEFORE AND AFTER
'------------------------------------------------------------------------------
        AddPoint Ids, Xs, Count, "benign_0025", -0.025
        AddPoint Ids, Xs, Count, "benign_1", -1#
        AddPoint Ids, Xs, Count, "benign_10", -10#
        AddPoint Ids, Xs, Count, "benign_40", -40#
        AddPoint Ids, Xs, Count, "benign_100", -100#
        AddPoint Ids, Xs, Count, "benign_400", -400#
'------------------------------------------------------------------------------
' THE WINDOW OPENS
'------------------------------------------------------------------------------
    'Exp(X) becomes subnormal near -708.4; the error is not yet visible at -709
        AddPoint Ids, Xs, Count, "edge_709", -709#
        AddPoint Ids, Xs, Count, "growth_716_1ulp", -716#
        AddPoint Ids, Xs, Count, "growth_724_1e12", -724.1
        AddPoint Ids, Xs, Count, "growth_731_1e9", -731.2
        AddPoint Ids, Xs, Count, "growth_738_1e6", -738.1
        AddPoint Ids, Xs, Count, "growth_742_1e4", -742.6
        AddPoint Ids, Xs, Count, "peak_74513", -745.13
'------------------------------------------------------------------------------
' THE WINDOW CLOSES
'------------------------------------------------------------------------------
    'Below about -745.14 the exponential underflows hard and the old U = 0#
    'branch took over, which is why the pre-existing tests at -746 passed
        AddPoint Ids, Xs, Count, "edge_74514", -745.14
        AddPoint Ids, Xs, Count, "past_746", -746#
        AddPoint Ids, Xs, Count, "past_1e300", -1E+300
End Sub


Private Sub AddPoint( _
    ByRef Ids() As String, _
    ByRef Xs() As Double, _
    ByRef Count As Long, _
    ByVal Id As String, _
    ByVal X As Double)
'
'==============================================================================
' AddPoint
'------------------------------------------------------------------------------
' PURPOSE
'   Appends one probe point to the parallel arrays.
'==============================================================================
'
'------------------------------------------------------------------------------
' APPEND
'------------------------------------------------------------------------------
        Ids(Count) = Id
        Xs(Count) = X
        Count = Count + 1
End Sub


Private Function EvaluatePoint( _
    ByVal Id As String, _
    ByVal X As Double) _
    As String
'
'==============================================================================
' EvaluatePoint
'------------------------------------------------------------------------------
' PURPOSE
'   Records the kernel inputs, the kernel result and both affected public CDFs
'   at a single point.
'
' WHY Exp(X) AND U - 1 ARE RECORDED
'   They are what the branch decision is made on. Exp(X) shows where the result
'   becomes subnormal, and U - 1 shows where it has already saturated to -1 -
'   the point from which the rescale can only do harm.
'
' WHY THE CDF ARGUMENT IS NEGATED
'   Both CDFs compute -PROB_Expm1(-z), so the probe point X corresponds to a
'   CDF evaluated at -X with unit rate, shape and scale.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Row                 As String          'Assembled CSV row
    Dim U                   As Double          'Exp(X), as actually rounded
    Dim V                   As Double          'U - 1, as actually rounded
    Dim Kernel              As Double          'PROB_Expm1(X)
    Dim ExpCdf              As Variant         'Exponential cumulative
    Dim WeibullCdf          As Variant         'Weibull cumulative
    Dim ErrNum              As Long            'Captured error number
'------------------------------------------------------------------------------
' KERNEL
'------------------------------------------------------------------------------
        U = Exp(X)
        V = U - 1#
        Kernel = PROB_Expm1(X)

        Row = Id & "," & FormatFullPrecision(X) & _
                   "," & FormatFullPrecision(U) & _
                   "," & FormatFullPrecision(V) & _
                   "," & FormatFullPrecision(Kernel)
'------------------------------------------------------------------------------
' EXPONENTIAL CUMULATIVE
'------------------------------------------------------------------------------
        On Error Resume Next
        Err.Clear
        ExpCdf = K_STATS_Exponential_Cumulative(-X, 1#)
        ErrNum = Err.Number
        Err.Clear
        On Error GoTo 0

        If ErrNum <> 0 Then
            Row = Row & ",ERROR,"
        ElseIf IsError(ExpCdf) Then
            Row = Row & ",CVERR,"
        Else
            Row = Row & ",OK," & FormatFullPrecision(CDbl(ExpCdf))
        End If
'------------------------------------------------------------------------------
' WEIBULL CUMULATIVE
'------------------------------------------------------------------------------
    'Unit shape and scale, so PowerValue is -X and the kernel argument matches
        On Error Resume Next
        Err.Clear
        WeibullCdf = K_STATS_Weibull_Cumulative(-X, 1#, 1#)
        ErrNum = Err.Number
        Err.Clear
        On Error GoTo 0

        If ErrNum <> 0 Then
            Row = Row & ",ERROR,"
        ElseIf IsError(WeibullCdf) Then
            Row = Row & ",CVERR,"
        Else
            Row = Row & ",OK," & FormatFullPrecision(CDbl(WeibullCdf))
        End If

    EvaluatePoint = Row
End Function


Private Function FieldCount(ByVal Text As String) As Long
'
'==============================================================================
' FieldCount
'------------------------------------------------------------------------------
' PURPOSE
'   Counts the comma-separated fields in an assembled row. Safe because
'   FormatFullPrecision separates its two parts with a semicolon, so no field
'   can contain a comma.
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
'   user cancels.
'
' WHY THIS EXISTS
'   ThisWorkbook.Path returns an http(s) URL when the workbook lives on
'   OneDrive or SharePoint, and Open cannot write to a URL. Dir$() cannot test
'   such a path either, so the dialog is the reliable route.
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
        InitialFileName:="expm1_saturation_probe.csv", _
        FileFilter:="Probe output (*.csv),*.csv", _
        Title:="Save the Expm1 saturation probe")

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
'   Renders a Double as a two-part sum "hi;lo", so hi + lo summed in Double
'   precision on the Python side reproduces the original Double exactly.
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
