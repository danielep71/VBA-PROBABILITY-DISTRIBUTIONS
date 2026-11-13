Attribute VB_Name = "M_STATS_PROBDIST_CHISQHOLD2"
Option Explicit

'Column count of the committed schema, checked against every row written
Private Const PROBE_FIELDS As Long = 14

'Shape slices are surface-specific; shape_id is provenance, echo_df is
'authoritative (see ShapeSlicesFor).
'==============================================================================
' M_STATS_PROBDIST_CHISQHOLD2
'------------------------------------------------------------------------------
' PURPOSE
'   Second independent HOLDOUT exporter for #13, Chi-square cumulative only. It
'   tests Amendment 1 (AMENDMENT_1.md): on the integrated incomplete-gamma
'   kernel, the log-domain candidate wins only where the transform
'   X / 2 rounded.
'
'   The first holdout (M_STATS_PROBDIST_CHISQHOLD) has been seen on the
'   current kernel, so it is fitting evidence for the amendment. This point
'   set shares no bucket and no degrees of freedom with it or with the fitting
'   arm.
'
' EVIDENCE CLASSES
'   exact_scaled  X = 2N * m with N = 2 ^ (k - 1), m = 2 ^ -1074: X / 2 is
'                 exactly N * m.
'   rounded       X = (2N + 1) * m: X / 2 falls on the midpoint between two
'                 subnormals, the only rounding a halving can produce.
'
' NOTHING IS WRITTEN AS A SUBNORMAL LITERAL
'   Every value is built by halving, doubling and integer scaling, all exact.
'
' WHY THE EXACT QUOTIENT IS NOT EXPORTED
'   The analyzer reconstructs it as a rational from the echoed binary64
'   operands.
'
' OUTPUT
'   The committed schema with echo_df in place of echo_shape, and no
'   echo_scale: this family always standardizes by 2, and the analyzer
'   supplies that rather than trusting an echoed constant.
'   Only the cumulative is written.
'
' USAGE
'   Run Probe_ChiSquareHoldout2, then validate with
'   analyze_positive_ratio_subnormal.py before committing the CSV.
'   Export_ChiSquareHoldout2 writes chisquare_holdout2.csv beside the workbook
'   with no dialog; ci/Export-ExcelGrids.ps1 runs it on the self-hosted runner
'   (#225).
'
' DEPENDENCIES
'   - K_STATS_ChiSquare_Density, _Cumulative, _Survival
'                                          (M_STATS_PROBDIST_TFAMILY)
'   - PROB_TryLogGamma1p, PROB_LogGamma    (M_STATS_PROBDIST_SPECIALFUNCS)
'   - PROB_TryExp, PROB_Expm1              (M_STATS_PROBDIST_CORE)
'
' UPDATED
'   2026-11-13
'==============================================================================


Public Sub Probe_ChiSquareHoldout2()
'
'==============================================================================
' Probe_ChiSquareHoldout2
'------------------------------------------------------------------------------
' PURPOSE
'   Asks for an output path, then writes the CSV there.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'Chosen output path
'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        Path = ResolveOutputPath()
        If Len(Path) = 0 Then Exit Sub          'User cancelled the picker

        WriteChiSquareHoldout2 Path
End Sub


Public Sub Export_ChiSquareHoldout2()
'
'==============================================================================
' Export_ChiSquareHoldout2
'------------------------------------------------------------------------------
' PURPOSE
'   Writes chisquare_holdout2.csv beside this workbook, with no dialog: the
'   entry point ci/Export-ExcelGrids.ps1 runs on the self-hosted runner
'   (#225). The workbook there is a local file, so its path is usable by Open.
'==============================================================================
'
'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        WriteChiSquareHoldout2 ThisWorkbook.Path & _
            Application.PathSeparator & "chisquare_holdout2.csv"
End Sub


Private Sub WriteChiSquareHoldout2(ByVal Path As String)
'
'==============================================================================
' WriteChiSquareHoldout2
'------------------------------------------------------------------------------
' PURPOSE
'   Evaluates every constructed point on all three surfaces and writes the CSV.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OutNum              As Integer         'Output file handle
    Dim Ids()               As String          'Point identifiers
    Dim Cons()              As String          'Construction class
    Dim Bits()              As Long            'Declared bucket
    Dim Xs()                As Double          'Public X
    Dim Scales()            As Double          'Always 2 in this family
    Dim Count               As Long            'Number of points
    Dim Row                 As String          'Assembled CSV row
    Dim I                   As Long            'Point index
    Dim S                   As Long            'Surface index
    Dim H                   As Long            'Shape-slice index
    Dim Rows                As Long            'Rows written
    Dim Surfaces            As Variant         'Surface names
    Dim Shapes              As Variant         'Shape slices
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

        BuildPoints Ids, Cons, Bits, Xs, Scales, Count
        Surfaces = Array("cumulative")

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        OutNum = FreeFile
        Open Path For Output As #OutNum

        Print #OutNum, "family,surface,point_id,construction,shape_id," & _
                       "bucket_bits,echo_x,echo_df," & _
                       "stored_standardx,log_standardx," & _
                       "current_status,current_value," & _
                       "candidate_status,candidate_value"

        Rows = 0
        For S = LBound(Surfaces) To UBound(Surfaces)
            Shapes = ShapeSlicesFor(CStr(Surfaces(S)))

            For H = LBound(Shapes) To UBound(Shapes)
                For I = 0 To Count - 1
                    Row = EvaluatePoint(CStr(Surfaces(S)), _
                                        CDbl(Shapes(H)), Ids(I), _
                                        Cons(I), Bits(I), Xs(I), Scales(I))

    'A row that does not match the committed schema is silently
    'unreadable: refuse to write it rather than emit evidence whose
    'columns cannot be trusted
                    If FieldCount(Row) <> PROBE_FIELDS Then
                        Err.Raise 5, , "Row " & Ids(I) & " has " & _
                                       FieldCount(Row) & " fields, " & _
                                       "expected " & PROBE_FIELDS
                    End If

                    Print #OutNum, Row
                    Rows = Rows + 1
                Next I
            Next H
        Next S

        Close #OutNum

    MsgBox "Chi-square positive-ratio HOLDOUT 2 complete: " & Count & _
           " construction(s), " & Rows & " rows written to" & _
           vbCrLf & Path, vbInformation, "Chi-square holdout"
    Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    On Error Resume Next
    Close #OutNum
    MsgBox "Chi-square positive-ratio probe failed: " & Err.Description, _
           vbExclamation, "Chi-square holdout"
End Sub


Private Function ShapeSlicesFor(ByVal Surface As String) As Variant
'
'==============================================================================
' ShapeSlicesFor
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the shape slices of the second holdout. Only the cumulative is
'   exported, so one set serves.
'
' NOTES
''   Degrees of freedom, twice the shapes of the Gamma second holdout:
'   EulerGamma/50, 1/(500 pi), sqrt3/50, ln3/5, e/5, 2/sqrt3, pi/2, 2 ln 2.5.
'==============================================================================
'
    ShapeSlicesFor = Array(0.0115443132980306, _
                           0.00063661977236758, _
                           0.034641016151378, _
                           0.21972245773362, _
                           0.5436563656918, _
                           1.15470053837926, _
                           1.5707963267949, _
                           1.83258146374832)
End Function


Private Sub BuildPoints( _
    ByRef Ids() As String, _
    ByRef Cons() As String, _
    ByRef Bits() As Long, _
    ByRef Xs() As Double, _
    ByRef Scales() As Double, _
    ByRef Count As Long)
'
'==============================================================================
' BuildPoints
'------------------------------------------------------------------------------
' PURPOSE
'   Constructs the point set of the second holdout: thirteen buckets (51 down
'   to 7), each with one exact halving and one rounded halving.
'
' THE CONSTRUCTION
'   With m = 2 ^ -1074 the subnormal ULP and N = 2 ^ (k - 1):
'
'       exact_scaled  X = 2N * m        X / 2 = N*m
'       rounded       X = (2N + 1) * m  X / 2 = N*m + m/2
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Cap                 As Long            'Array capacity
    Dim MinSub              As Double          '2 ^ -1074, built by halving
    Dim N                   As Double          '2 ^ (k - 1) for the bucket
    Dim Big                 As Double          '2 ^ 300, built by doubling
    Dim Buckets             As Variant         'Bucket widths to cover
    Dim B                   As Long            'Bucket index
    Dim K                   As Long            'Bits in the current bucket
    Dim J                   As Long            'Doubling counter
'------------------------------------------------------------------------------
' ALLOCATE
'------------------------------------------------------------------------------
    Cap = 128
    ReDim Ids(0 To Cap - 1)
    ReDim Cons(0 To Cap - 1)
    ReDim Bits(0 To Cap - 1)
    ReDim Xs(0 To Cap - 1)
    ReDim Scales(0 To Cap - 1)
    Count = 0
'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
    MinSub = SmallestPositiveDouble()

    Big = 1#
    For J = 1 To 300
        Big = Big * 2#
    Next J

    Buckets = Array(51, 47, 43, 42, 39, 35, 31, 27, 23, 19, 15, 11, 7)

    For B = LBound(Buckets) To UBound(Buckets)
        K = CLng(Buckets(B))

    'N = 2 ^ (K - 1) by doubling; exact, and never large enough to lose bits
        N = 1#
        For J = 2 To K
            N = N * 2#
        Next J

        AddPoint Ids, Cons, Bits, Xs, Scales, Count, _
                 "c2_even_k" & K, "exact_scaled", K, (2# * N) * MinSub, 2#
        AddPoint Ids, Cons, Bits, Xs, Scales, Count, _
                 "c2_tie_k" & K, "rounded", K, (2# * N + 1#) * MinSub, 2#
    Next B
End Sub


Private Function SmallestPositiveDouble() As Double
'
'==============================================================================
' SmallestPositiveDouble
'------------------------------------------------------------------------------
' PURPOSE
'   Returns 2 ^ -1074, the subnormal ULP, built by halving.
'
' ASSIGN, THEN TEST
'   The stored value decides, not the expression. VBA does not necessarily
'   round an intermediate expression to Double at every operator, so testing
'   MinSub / 2# in a loop condition would let the loop run past the floor and
'   leave zero behind. See the IMPORTANT note under "Start here" in
'   CONTRIBUTING.md (forced Double rounding boundaries).
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Smallest            As Double          'Running power of two
    Dim Halved              As Double          'Candidate next value
    Dim I                   As Long            'Halving counter
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    Smallest = 1#
    For I = 1 To 1074
        Halved = Smallest / 2#
        If Halved = 0# Then Exit For
        Smallest = Halved
    Next I

    SmallestPositiveDouble = Smallest
End Function


Private Sub AddPoint( _
    ByRef Ids() As String, _
    ByRef Cons() As String, _
    ByRef Bits() As Long, _
    ByRef Xs() As Double, _
    ByRef Scales() As Double, _
    ByRef Count As Long, _
    ByVal Id As String, _
    ByVal Construction As String, _
    ByVal BucketBits As Long, _
    ByVal X As Double, _
    ByVal ScaleParam As Double)
'
'==============================================================================
' AddPoint
'------------------------------------------------------------------------------
' PURPOSE
'   Appends one constructed point to the parallel arrays.
'==============================================================================
'
'------------------------------------------------------------------------------
' APPEND
'------------------------------------------------------------------------------
        Ids(Count) = Id
        Cons(Count) = Construction
        Bits(Count) = BucketBits
        Xs(Count) = X
        Scales(Count) = ScaleParam
        Count = Count + 1
End Sub


Private Function EvaluatePoint( _
    ByVal Surface As String, _
    ByVal Shape As Double, _
    ByVal Id As String, _
    ByVal Construction As String, _
    ByVal BucketBits As Long, _
    ByVal X As Double, _
    ByVal ScaleParam As Double) _
    As String
'
'==============================================================================
' EvaluatePoint
'------------------------------------------------------------------------------
' PURPOSE
'   Emits one schema row: the echoed operands, the standardized argument as
'   stored, the log as VBA forms it, and both branch results.
'
' WHY BOTH THE STORED VALUE AND THE LOG ARE RECORDED
'   They isolate two different error sources. The transform path loses
'   information going from the exact ratio to the stored Double. The log path
'   loses information subtracting two logarithms of magnitude near 700. #13
'   anticipates the second limiting the candidate near 1E-12; recording both
'   lets that be measured rather than assumed.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Row                 As String          'Assembled CSV row
    Dim StandardX           As Double          '0.5 * X, as stored
    Dim LogStandardX        As Double          'Log(X) - Log(2)
    Dim HasStandard         As Boolean         'Division succeeded
    Dim Current             As Variant         'Current public result
    Dim Candidate           As Double          'Log-path mirror
    Dim CandidateOk         As Boolean         'Mirror produced a value
    Dim ErrNum              As Long            'Captured error number
'------------------------------------------------------------------------------
' STANDARDIZE
'------------------------------------------------------------------------------
    'Exactly as the production path does it, so the stored value recorded here
    'is the one the kernels actually receive
    'Mirrors production exactly: the multiplication is inline and there is
    'no Try wrapper, so the only failure mode is underflow to zero
        StandardX = 0.5 * X
        HasStandard = True

        On Error Resume Next
        Err.Clear
        LogStandardX = Log(X) - Log(2#)
        If Err.Number <> 0 Then LogStandardX = 0#
        Err.Clear
        On Error GoTo 0

        Row = "chisquare," & Surface & "," & Id & "," & Construction & _
              "," & ShapeId(Shape) & "," & BucketBits & "," & _
              FormatFullPrecision(X) & "," & _
              FormatFullPrecision(Shape) & "," & _
              IIf(HasStandard, FormatFullPrecision(StandardX), "") & "," & _
              FormatFullPrecision(LogStandardX)
'------------------------------------------------------------------------------
' CURRENT PATH
'------------------------------------------------------------------------------
        On Error Resume Next
        Err.Clear
        Select Case Surface
            Case "density"
                Current = K_STATS_ChiSquare_Density(X, Shape)
            Case "cumulative"
                Current = K_STATS_ChiSquare_Cumulative(X, Shape)
            Case Else
                Current = K_STATS_ChiSquare_Survival(X, Shape)
        End Select
        ErrNum = Err.Number
        Err.Clear
        On Error GoTo 0

        If ErrNum <> 0 Then
            Row = Row & ",ERROR,"
        ElseIf IsError(Current) Then
            Row = Row & ",CVERR,"
        Else
            Row = Row & ",OK," & FormatFullPrecision(CDbl(Current))
        End If
'------------------------------------------------------------------------------
' CANDIDATE LOG PATH
'------------------------------------------------------------------------------
        CandidateOk = TryCandidate(Surface, Shape, LogStandardX, _
                                   StandardX, ScaleParam, Candidate)

        If CandidateOk Then
            Row = Row & ",OK," & FormatFullPrecision(Candidate)
        Else
            Row = Row & ",ERROR,"
        End If

    EvaluatePoint = Row
End Function


Private Function TryCandidate( _
    ByVal Surface As String, _
    ByVal Shape As Double, _
    ByVal LogStandardX As Double, _
    ByVal StandardX As Double, _
    ByVal ScaleParam As Double, _
    ByRef Result As Double) _
    As Boolean
'
'==============================================================================
' TryCandidate
'------------------------------------------------------------------------------
' PURPOSE
'   Mirrors the log-domain branch proposed in #13, without touching production
'   source. This is a measurement instrument, not a candidate implementation:
'   it exists so the two branches can be compared at identical inputs.
'
' THE FORMS
'   For a standardized argument far below the shape, the regularized lower
'   incomplete gamma is dominated by its leading series term:
'
'       Log P = Shape * Log(z) - LogGamma(1 + Shape)
'
'   The survival is taken as -Expm1(Log P) rather than 1 - Exp(Log P), which
'   would cancel catastrophically when P is tiny. The density is the ordinary
'   log form, evaluated without ever forming z ^ (Shape - 1) directly.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HalfDF              As Double          'df / 2, the kernel shape
    Dim LogP                As Double          'Log of the lower tail
    Dim LogGamma1pShape     As Double          'LogGamma(1 + Shape)
    Dim FailMsg             As String          'Kernel diagnostic
    Dim P                   As Double          'Lower tail probability
    Dim LogDensity          As Double          'Log of the density
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

    'The kernel receives df / 2, not df. Using the public input here would
    'silently measure a different distribution than the one under test.
        HalfDF = 0.5 * Shape

        If Not PROB_TryLogGamma1p(HalfDF, LogGamma1pShape, FailMsg) Then
            Exit Function
        End If

        Select Case Surface

            Case "density"
    'Log f = (a - 1) * Log(z) - z - LogGamma(a) - Log(2), with a = df / 2
                LogDensity = (HalfDF - 1#) * LogStandardX _
                             - StandardX _
                             - PROB_LogGamma(HalfDF) _
                             - Log(2#)
                If Not PROB_TryExp(LogDensity, Result) Then Exit Function

            Case "cumulative"
                LogP = HalfDF * LogStandardX - LogGamma1pShape
                If Not PROB_TryExp(LogP, Result) Then Exit Function

            Case Else
    'Q = -Expm1(Log P): never 1 - Exp(Log P), which cancels when P is tiny
                LogP = HalfDF * LogStandardX - LogGamma1pShape
                Result = -PROB_Expm1(LogP)

        End Select

        TryCandidate = True
    Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    Err.Clear
End Function


Private Function ShapeId(ByVal Shape As Double) As String
'
'==============================================================================
' ShapeId
'------------------------------------------------------------------------------
' PURPOSE
'   A readable, stable grouping label for a shape slice.
'
' NOT AUTHORITATIVE
'   echo_shape carries the binary64 value and is what the analyzer computes
'   from. This label exists only so reports can be grouped and cited. Keeping
'   the two separate is what stops a 0.001 that actually parsed as
'   0.0010000000000000002 from hiding behind a tidy name.
'==============================================================================
'
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    ShapeId = "s" & Replace(Format$(Shape, "0.0######"), ",", ".")
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
        InitialFileName:="chisquare_holdout2.csv", _
        FileFilter:="Probe output (*.csv),*.csv", _
        Title:="Save the Chi-square positive-ratio HOLDOUT")

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
'   precision on the Python side reproduces the original Double exactly. A
'   15-digit decimal alone cannot round-trip a subnormal, which is most of the
'   values in this study.
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
