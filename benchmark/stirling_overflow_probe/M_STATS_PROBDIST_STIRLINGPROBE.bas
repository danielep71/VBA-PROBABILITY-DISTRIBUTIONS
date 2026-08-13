Attribute VB_Name = "M_STATS_PROBDIST_STIRLINGPROBE"
Option Explicit

'Column count of the output header, checked against every row written
Private Const PROBE_FIELDS As Long = 16

'==============================================================================
' M_STATS_PROBDIST_STIRLINGPROBE
'------------------------------------------------------------------------------
' PURPOSE
'   Characterization probe for the PROB_StirlingError small-N recurrence, which
'   forms the explicit reciprocal
'
'       (N + 1#) / N
'
'   for every 0 < N < 0.5. Once 1 / N exceeds the Double range the quotient is
'   not representable and the expression faults BEFORE Log() is ever called.
'   The boundary is 1 / DoubleMax, about 5.5626846462680035E-309, which is
'   inside the subnormal range and inside the accepted parameter domain: the
'   Gamma shape guard is PROB_IsPositiveWithinSupportedMagnitude, which tests
'   Shape > 0 and an UPPER bound of 1E100 only. No lower cutoff exists, so
'   K_STATS_Gamma_Density(1, 1E-320, 1) reaches the branch.
'
' WHAT THIS PROBE DECIDES
'   Three things that cannot be settled outside real VBA:
'
'     1. WHETHER the current expression fails at the predicted boundary, and
'        WHERE that boundary actually sits when measured rather than derived.
'     2. HOW it fails. Python returns an infinity here. If VBA raises error 6
'        the fault propagates to Err_Handler and the caller sees CVErr, which
'        is a visible failure. If VBA instead yields 1.#INF the value flows on
'        into Exp() and the caller sees a plausible WRONG density. Those are
'        different severities and the remediation urgency differs with them.
'     3. WHETHER the two candidate replacements are finite across the whole
'        branch and neither degrades where the current form already works.
'
' THE TWO CANDIDATES
'   B  the minimal rearrangement, keeping the recurrence and removing only the
'      reciprocal:
'          StirlingError(N + 1) + (N + 0.5) * (Log1p(N) - Log(N)) - 1
'      PROB_Log1p returns X unchanged once 1 + X rounds back to one, so B needs
'      neither the increment to survive nor 1 / N to be representable.
'
'   C  the direct definition, which never recurses:
'          LogGamma1p(N) - (N + 0.5) * Log(N) + N - HALF_LOG_TWO_PI
'      Measured through PROB_TryLogGamma1p rather than PROB_LogGamma1pSeries,
'      which is Private to SPECIALFUNCS and unreachable from this module. For
'      N <= PROB_LG1P_SERIES_MAX the wrapper delegates to that same series, so
'      the two agree over every point this probe evaluates.
'
' WHY THE EXTREME POINTS ARE BUILT BY ARITHMETIC
'   A VBA source literal cannot be trusted to denote a subnormal Double, and
'   Val() cannot be trusted to parse one. Writing 1E-320 and measuring what
'   comes back would confound a parser defect with a kernel defect. The
'   subnormal points are therefore built by halving and by adding exact
'   multiples of the smallest positive Double, both of which are lossless. The
'   decimal points are retained alongside them and echo_n reports what each one
'   actually parsed to.
'
' WHERE THE BOUNDARY ACTUALLY IS
'   1 / DoubleMax is 2 ^ -1024 * (1 + 2 ^ -53), which exceeds 2 ^ -1024 by
'   2 ^ -1077, an eighth of a subnormal ULP. That excess is below the
'   representable grid, so 1 / DoubleMax rounds to exactly 2 ^ -1024 and the
'   decisive pair is:
'
'       first failing N   2 ^ -1024              1 / N is 2 ^ 1024, not finite
'       last safe N       2 ^ -1024 + 2 ^ -1074  1 / N is DoubleMax exactly
'
'   These are adjacent Doubles. 2 ^ -1023 is NOT adjacent to 2 ^ -1024: the two
'   are 2 ^ 50 subnormal ULPs apart, so 2 ^ -1023 is kept only as a safe anchor
'   well inside the working region, not as half of the seam.
'
' USAGE
'   Run Probe_StirlingOverflow and choose an output path in the dialog.
'
' OUTPUT (one row per point, header first)
'   point_id, construction, echo_n, current_status, current_err, current_desc,
'   current_value, b_status, b_err, b_value, c_status, c_err, c_value,
'   density_status, density_err, density_value
'
'   All values are full-precision hi;lo tokens; sum the parts in Double
'   precision to recover the original. Status is OK, ERROR or CVERR.
'
' SCOPE
'   This is a characterization probe, not a benchmark. It promotes no grid row,
'   claims no threshold and touches no registry. It exists to decide one
'   question before any source edit is proposed.
'
' DEPENDENCIES
'   - PROB_StirlingError, PROB_TryLogGamma1p (M_STATS_PROBDIST_SPECIALFUNCS)
'   - PROB_Log1p, PROB_HALF_LOG_TWO_PI       (M_STATS_PROBDIST_CORE)
'   - K_STATS_Gamma_Density                  (M_STATS_PROBDIST_CONTINUOUS)
'
' UPDATED
'   2026-08-13
'==============================================================================


Public Sub Probe_StirlingOverflow()
'
'==============================================================================
' Probe_StirlingOverflow
'------------------------------------------------------------------------------
' PURPOSE
'   Evaluates every probe point and writes the characterization CSV.
'
' ERROR POLICY
'   Per-point faults are captured as data, not raised: each evaluation runs
'   under On Error Resume Next and reports its own Err.Number. Only a failure
'   of the file handling itself aborts the run.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'Chosen output path
    Dim OutNum              As Integer         'Output file handle
    Dim Ids()               As String          'Point identifiers
    Dim Hows()              As String          'How each point was built
    Dim Ns()                As Double          'The points themselves
    Dim Count               As Long            'Number of points
    Dim Row                 As String          'Assembled CSV row
    Dim I                   As Long            'Point index
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

        Path = ResolveOutputPath()
        If Len(Path) = 0 Then Exit Sub          'User cancelled the picker

        BuildPoints Ids, Hows, Ns, Count
'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        OutNum = FreeFile
        Open Path For Output As #OutNum

        Print #OutNum, "point_id,construction,echo_n," & _
                       "current_status,current_err,current_desc," & _
                       "current_value," & _
                       "b_status,b_err,b_value," & _
                       "c_status,c_err,c_value," & _
                       "density_status,density_err,density_value"

        For I = 0 To Count - 1
            Row = EvaluatePoint(Ids(I), Hows(I), Ns(I))

    'A row that does not match the header is silently unreadable: refuse to
    'write it rather than emit a file whose columns cannot be trusted
            If FieldCount(Row) <> PROBE_FIELDS Then
                Err.Raise 5, , "Row " & Ids(I) & " has " & FieldCount(Row) & _
                               " fields, expected " & PROBE_FIELDS
            End If

            Print #OutNum, Row
        Next I

        Close #OutNum

    MsgBox "Stirling overflow probe complete: " & Count & _
           " point(s) written to" & vbCrLf & Path, _
           vbInformation, "Stirling probe"
    Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    On Error Resume Next
    Close #OutNum
    MsgBox "Stirling overflow probe failed: " & Err.Description, _
           vbExclamation, "Stirling probe"
End Sub


Private Sub BuildPoints( _
    ByRef Ids() As String, _
    ByRef Hows() As String, _
    ByRef Ns() As Double, _
    ByRef Count As Long)
'
'==============================================================================
' BuildPoints
'------------------------------------------------------------------------------
' PURPOSE
'   Constructs the probe point set. Powers of two are built by halving so that
'   the subnormal points are exact and independent of literal parsing; decimal
'   points are kept as literals and adjudicated by EchoN.
'
' THE BOUNDARY, EXACTLY
'   Halving from one reaches 2 ^ -1024 and 2 ^ -1074 losslessly. The adjacent
'   Doubles either side of 2 ^ -1024 are then formed by adding and subtracting
'   one smallest-positive-Double, which is exact in the subnormal range because
'   the spacing there is uniform.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Cap                 As Long            'Array capacity
    Dim Halved              As Double          'Running power of two
    Dim K                   As Long            'Halving counter
    Dim P1023               As Double          '2 ^ -1023, safe anchor
    Dim P1024               As Double          '2 ^ -1024, first failing N
    Dim MinSub              As Double          '2 ^ -1074, one subnormal ULP
'------------------------------------------------------------------------------
' ALLOCATE
'------------------------------------------------------------------------------
    Cap = 32
    ReDim Ids(0 To Cap - 1)
    ReDim Hows(0 To Cap - 1)
    ReDim Ns(0 To Cap - 1)
    Count = 0
'------------------------------------------------------------------------------
' NORMAL-RANGE ANCHORS
'------------------------------------------------------------------------------
    'These already work; they are here so a regression in the safe region is
    'visible rather than assumed away
        AddPoint Ids, Hows, Ns, Count, "top_of_branch", "literal", 0.25
        AddPoint Ids, Hows, Ns, Count, "e100", "literal", 1E-100
        AddPoint Ids, Hows, Ns, Count, "e244_worst_known", "literal", 1E-244
        AddPoint Ids, Hows, Ns, Count, "e307", "literal", 1E-307
        AddPoint Ids, Hows, Ns, Count, "e308", "literal", 1E-308
'------------------------------------------------------------------------------
' BUILD THE POWERS OF TWO
'------------------------------------------------------------------------------
    'Exact by construction; nothing here depends on literal parsing
        Halved = 1#
        For K = 1 To 1074
            Halved = Halved / 2#
            If Halved = 0# Then Exit For

            Select Case K
                Case 1023: P1023 = Halved
                Case 1024: P1024 = Halved
                Case 1074: MinSub = Halved
            End Select
        Next K
'------------------------------------------------------------------------------
' THE DECISIVE SEAM
'------------------------------------------------------------------------------
    'Adjacent Doubles either side of the reciprocal threshold. If current reads
    'OK at the last-safe point and ERROR at the first-failing one, the boundary
    'is established at the finest resolution the format allows
        AddPoint Ids, Hows, Ns, Count, "p2_1023_anchor", _
                 "2 ^ -1023 by halving", P1023
        AddPoint Ids, Hows, Ns, Count, "seam_last_safe", _
                 "2 ^ -1024 + 2 ^ -1074", P1024 + MinSub
        AddPoint Ids, Hows, Ns, Count, "seam_first_fail", _
                 "2 ^ -1024", P1024
        AddPoint Ids, Hows, Ns, Count, "seam_below", _
                 "2 ^ -1024 - 2 ^ -1074", P1024 - MinSub
'------------------------------------------------------------------------------
' DEEP SUBNORMAL
'------------------------------------------------------------------------------
    'Constructed rather than parsed, plus the literal twins so that a
    'disagreement between the two identifies the parser rather than the kernel
        AddPoint Ids, Hows, Ns, Count, "c1e310", _
                 "1E-300 / 1E+10", 1E-300 / 1E+10
        AddPoint Ids, Hows, Ns, Count, "c1e320", _
                 "1E-300 / 1E+20", 1E-300 / 1E+20
        AddPoint Ids, Hows, Ns, Count, "d1e310", "literal", 1E-310
        AddPoint Ids, Hows, Ns, Count, "d1e320", "literal", 1E-320
        AddPoint Ids, Hows, Ns, Count, "p2_1074_min_subnormal", _
                 "2 ^ -1074 by halving", MinSub
End Sub


Private Sub AddPoint( _
    ByRef Ids() As String, _
    ByRef Hows() As String, _
    ByRef Ns() As Double, _
    ByRef Count As Long, _
    ByVal Id As String, _
    ByVal How As String, _
    ByVal N As Double)
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
        Hows(Count) = How
        Ns(Count) = N
        Count = Count + 1
End Sub


Private Function EvaluatePoint( _
    ByVal Id As String, _
    ByVal How As String, _
    ByVal N As Double) _
    As String
'
'==============================================================================
' EvaluatePoint
'------------------------------------------------------------------------------
' PURPOSE
'   Measures the current implementation, both candidates and one real public
'   surface at a single point, capturing each fault as data.
'
' WHY THE DENSITY CALL USES X = 1
'   K_STATS_Gamma_Density(1, N, 1) leaves the shape as the only extreme
'   argument, so a failure there cannot be attributed to X or to the scale.
'   At X = 1 and Scale = 1 the density is Exp(-1) / Gamma(N), which for small N
'   is about N / e: SMALL and finite, not large. It is 3.68E-101 at N = 1E-100
'   and 3.68E-321 at N = 1E-320, both representable. So a CVErr here cannot be
'   excused as legitimate overflow; it exposes the recurrence directly. The one
'   exception is N = 2 ^ -1074, where the true density is about 1.8E-324 and
'   correctly underflows to zero.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Row                 As String          'Assembled CSV row
    Dim Current             As Double          'Current PROB_StirlingError
    Dim BValue              As Double          'Candidate B
    Dim CValue              As Double          'Candidate C
    Dim Kernel              As Double          'PROB_TryLogGamma1p result
    Dim FailMsg             As String          'Kernel diagnostic
    Dim Density             As Variant         'K_STATS_Gamma_Density result
    Dim ErrNum              As Long            'Captured error number
    Dim ErrDesc             As String          'Captured error description
    Dim Ok                  As Boolean         'Kernel Boolean return
'------------------------------------------------------------------------------
' ECHO THE ARGUMENT
'------------------------------------------------------------------------------
    'Written first so a mis-parsed literal is visible before anything else
        Row = Id & "," & How & "," & FormatFullPrecision(N)
'------------------------------------------------------------------------------
' CURRENT IMPLEMENTATION
'------------------------------------------------------------------------------
        On Error Resume Next
        Err.Clear
        Current = PROB_StirlingError(N)
        ErrNum = Err.Number
        ErrDesc = Err.Description
        Err.Clear
        On Error GoTo 0

        If ErrNum <> 0 Then
            Row = Row & ",ERROR," & ErrNum & "," & CleanDesc(ErrDesc) & ","
        Else
            'The empty field is current_desc: only this block carries one, so
            'the OK branch must still emit it or every successful row is
            'written one column left of the header
            Row = Row & ",OK,0,," & FormatFullPrecision(Current)
        End If
'------------------------------------------------------------------------------
' CANDIDATE B
'------------------------------------------------------------------------------
        On Error Resume Next
        Err.Clear
        BValue = PROB_StirlingError(N + 1#) + _
                 (N + 0.5) * (PROB_Log1p(N) - Log(N)) - 1#
        ErrNum = Err.Number
        Err.Clear
        On Error GoTo 0

        If ErrNum <> 0 Then
            Row = Row & ",ERROR," & ErrNum & ","
        Else
            Row = Row & ",OK,0," & FormatFullPrecision(BValue)
        End If
'------------------------------------------------------------------------------
' CANDIDATE C
'------------------------------------------------------------------------------
        On Error Resume Next
        Err.Clear
        Ok = PROB_TryLogGamma1p(N, Kernel, FailMsg)
        If Ok Then
            CValue = Kernel - (N + 0.5) * Log(N) + N - PROB_HALF_LOG_TWO_PI
        End If
        ErrNum = Err.Number
        Err.Clear
        On Error GoTo 0

        If ErrNum <> 0 Then
            Row = Row & ",ERROR," & ErrNum & ","
        ElseIf Not Ok Then
            Row = Row & ",ERROR,0,"
        Else
            Row = Row & ",OK,0," & FormatFullPrecision(CValue)
        End If
'------------------------------------------------------------------------------
' PUBLIC SURFACE
'------------------------------------------------------------------------------
        On Error Resume Next
        Err.Clear
        Density = K_STATS_Gamma_Density(1#, N, 1#)
        ErrNum = Err.Number
        Err.Clear
        On Error GoTo 0

        If ErrNum <> 0 Then
            Row = Row & ",ERROR," & ErrNum & ","
        ElseIf IsError(Density) Then
            Row = Row & ",CVERR," & ErrorNumberOf(Density) & ","
        Else
            Row = Row & ",OK,0," & FormatFullPrecision(CDbl(Density))
        End If

    EvaluatePoint = Row
End Function


Private Function ErrorNumberOf(ByVal V As Variant) As String
'
'==============================================================================
' ErrorNumberOf
'------------------------------------------------------------------------------
' PURPOSE
'   Renders the numeric code carried by an Error variant, without letting a
'   coercion failure abort the probe.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Code                As Long            'Extracted error code
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler
        Code = CLng(V)
        ErrorNumberOf = CStr(Code)
    Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    ErrorNumberOf = "unreadable"
End Function


Private Function FieldCount(ByVal Text As String) As Long
'
'==============================================================================
' FieldCount
'------------------------------------------------------------------------------
' PURPOSE
'   Counts the comma-separated fields in an assembled row. Safe because
'   CleanDesc removes commas from descriptions and FormatFullPrecision
'   separates its two parts with a semicolon, so no field can contain one.
'==============================================================================
'
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    FieldCount = UBound(Split(Text, ",")) + 1
End Function


Private Function CleanDesc(ByVal Text As String) As String
'
'==============================================================================
' CleanDesc
'------------------------------------------------------------------------------
' PURPOSE
'   Makes an Err.Description safe to place in a comma-separated field.
'==============================================================================
'
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    CleanDesc = Replace(Replace(Trim$(Text), ",", ";"), vbCrLf, " ")
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
        InitialFileName:="stirling_overflow_probe.csv", _
        FileFilter:="Probe output (*.csv),*.csv", _
        Title:="Save the Stirling overflow probe")

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
'
' LIMITATION
'   At subnormal magnitudes the residual X - hi is itself below the
'   representable grid, so the pair cannot round-trip exactly there. The
'   analysis compares arguments as Doubles rather than as decimals for exactly
'   this reason; see the echo_n column.
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
