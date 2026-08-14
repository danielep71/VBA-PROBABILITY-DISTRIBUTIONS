Attribute VB_Name = "M_STATS_PROBDIST_GAMMAPROBE"
Option Explicit

'Column count of the committed schema, checked against every row written
Private Const PROBE_FIELDS As Long = 15

'Shape is part of the experimental design, not a setting: each surface has a
'different conditioning problem and one shape cannot serve all three. At
'Shape is part of the experimental design, not a setting: each surface has
'a different conditioning problem and one shape cannot serve all three. At
'Shape = 0.5 the survival is already correctly saturated at 1.0 across the
'whole subnormal ladder and cannot constrain dispatch. The slices are
'therefore surface-specific; shape_id is provenance, echo_shape is
'authoritative.
Private Const GAMMA_SURFACES As String = "density|cumulative|survival"

'==============================================================================
' M_STATS_PROBDIST_GAMMAPROBE
'------------------------------------------------------------------------------
' PURPOSE
'   Phase A1 exporter for ICR-P1-01A (#13), Gamma arm.
'
'   Gamma standardizes through PROB_TryDivide(X, ScaleParam, StandardX). The
'   division is correctly rounded; the defect is that a positive mathematical
'   ratio can become a low-precision subnormal or a hard zero, after which
'   downstream code treats the rounded value as an exact support coordinate.
'
'   This exporter measures, per public surface, what the current direct path
'   returns and what a log-domain candidate would return, across constructed
'   points whose exact quotient is known.
'
' PHASE A IS MEASUREMENT ONLY
'   No cutoff is proposed here and no production source is changed. The
'   crossover constants quoted in #13 are hypotheses until this evidence and
'   the Chi-square arm are analysed.
'
' EVIDENCE CLASSES
'   landmark          q = N * m with N = 2 ^ (k - 1) and m = 2 ^ -1074. Exactly
'                     representable, carrying exactly k significand bits.
'                     Isolates downstream behaviour from transform error.
'   transform_stress  exact quotients 1/3, 1/2 and 2/3 of a ULP above the
'                     landmark, reached with integer numerators over scales of
'                     3, 2 and 3. Both sides of the rounding midpoint.
'   underflow         a positive mathematical ratio that stores as zero.
'                     Classified apart, never as bucket 0.
'   decimal_twin      the same magnitude written as a source decimal.
'                     Diagnostic only; the analyzer compares it against its
'                     constructed partner to detect a parser difference.
'
' NOTHING IS WRITTEN AS A SUBNORMAL LITERAL
'   A VBA source literal cannot be relied on to denote a subnormal, and a
'   mis-parsed literal would silently test a different number. Every
'   authoritative
'   value here is built by halving and integer scaling, both exact. The decimal
'   twins exist precisely to prove that choice was necessary or unnecessary.
'
' WHY THE EXACT QUOTIENT IS NOT EXPORTED
'   The analyzer reconstructs it as a rational from the echoed binary64
'   operands. Serialising it here would let this module launder its own
'   rounding into the reference, which is the failure #17 exists to prevent.
'
' OUTPUT
'   The committed 13-column schema, long form: one row per public surface, so
'   each constructed point contributes three rows.
'
' USAGE
'   Run Probe_GammaPositiveRatio and choose an output path. Validate with
'   analyze_positive_ratio_subnormal.py before committing the CSV.
'
' DEPENDENCIES
'   - K_STATS_Gamma_Density, _Cumulative, _Survival
'                                      (M_STATS_PROBDIST_CONTINUOUS)
'   - PROB_TryLogGamma1p, PROB_LogGamma
'                                      (M_STATS_PROBDIST_SPECIALFUNCS)
'   - PROB_Expm1, PROB_TryExp                      (M_STATS_PROBDIST_CORE)
'
' UPDATED
'   2026-08-14
'==============================================================================


Public Sub Probe_GammaPositiveRatio()
'
'==============================================================================
' Probe_GammaPositiveRatio
'------------------------------------------------------------------------------
' PURPOSE
'   Evaluates every constructed point on all three surfaces and writes the CSV.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'Chosen output path
    Dim OutNum              As Integer         'Output file handle
    Dim Ids()               As String          'Point identifiers
    Dim Cons()              As String          'Construction class
    Dim Bits()              As Long            'Declared bucket
    Dim Xs()                As Double          'Public X
    Dim Scales()            As Double          'Public ScaleParam
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

        Path = ResolveOutputPath()
        If Len(Path) = 0 Then Exit Sub          'User cancelled the picker

        BuildPoints Ids, Cons, Bits, Xs, Scales, Count
        Surfaces = Array("density", "cumulative", "survival")

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        OutNum = FreeFile
        Open Path For Output As #OutNum

        Print #OutNum, "family,surface,point_id,construction,shape_id," & _
                       "bucket_bits,echo_x,echo_shape,echo_scale," & _
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

    MsgBox "Gamma positive-ratio probe complete: " & Count & _
           " construction(s), " & Rows & " rows written to" & _
           vbCrLf & Path, vbInformation, "Gamma probe"
    Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    On Error Resume Next
    Close #OutNum
    MsgBox "Gamma positive-ratio probe failed: " & Err.Description, _
           vbExclamation, "Gamma probe"
End Sub


Private Function ShapeSlicesFor(ByVal Surface As String) As Variant
'
'==============================================================================
' ShapeSlicesFor
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the shape slices for one surface. The sets are surface-specific
'   because each surface has a different conditioning problem, and one
'   universal set wastes evidence on slices that cannot discriminate.
'
' WHY SURVIVAL IS DIFFERENT
'   At Shape = 0.5 with a standardized argument near 1E-308 the lower tail is
'   about 1E-154, so Q = 1 - P rounds to exactly 1 across the whole subnormal
'   ladder. Both branches return 1 and both are right, so the slice carries no
'   information about dispatch. Near Shape = 0.001 the survival stays close to
'   0.5 throughout, which keeps the bucket varying while the probability
'   regime stays put. Shape = 0.5 is retained so that the saturation is
'   recorded as characterisation rather than silently dropped.
'
' WHY SHAPE = 1 IS IN THE OTHER TWO
'   The z ^ (Shape - 1) dependence disappears there, so transform sensitivity
'   should behave differently. It is a structural control: if the analyzer
'   reports otherwise, that is immediately informative.
'==============================================================================
'
'------------------------------------------------------------------------------
' SELECT
'------------------------------------------------------------------------------
    Select Case Surface
        Case "survival"
            ShapeSlicesFor = Array(0.0001, 0.0005, 0.001, 0.002, 0.5)
        Case Else
            ShapeSlicesFor = Array(0.001, 0.01, 0.1, 0.25, 0.5, 0.75, 1#)
    End Select
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
'   Constructs the point set: fifteen buckets from 52 bits down to 1, each with
'   a landmark and three transform stresses, plus hard-underflow cases and two
'   decimal twins.
'
' THE CONSTRUCTION
'   With m = 2 ^ -1074 the subnormal ULP and N = 2 ^ (k - 1):
'
'       landmark   X = N * m,         Scale = 1   exact quotient N*m
'       1/3 ULP    X = (3N + 1) * m,  Scale = 3   exact quotient N*m + m/3
'       1/2 ULP    X = (2N + 1) * m,  Scale = 2   exact quotient N*m + m/2
'       2/3 ULP    X = (3N + 2) * m,  Scale = 3   exact quotient N*m + 2m/3
'
'   Every X above is exactly representable, including at k = 52 where the
'   scale-3 forms sit near 1.5 * MIN_NORMAL and the spacing is still one m.
'   The integer coefficients stay below 2 ^ 53, so forming them in Double is
'   exact, and multiplying by m is a pure exponent shift.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Cap                 As Long            'Array capacity
    Dim MinSub              As Double          '2 ^ -1074, built by halving
    Dim N                   As Double          '2 ^ (k - 1) for the bucket
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

    MinSub = SmallestPositiveDouble()
    Buckets = Array(52, 48, 44, 40, 36, 32, 28, 24, 20, 16, 12, 8, 4, 2, 1)
'------------------------------------------------------------------------------
' BUCKETS
'------------------------------------------------------------------------------
    For B = LBound(Buckets) To UBound(Buckets)
        K = CLng(Buckets(B))

    'N = 2 ^ (K - 1) by doubling; exact, and never large enough to lose bits
        N = 1#
        For J = 2 To K
            N = N * 2#
        Next J

        AddPoint Ids, Cons, Bits, Xs, Scales, Count, _
                 "g_landmark_k" & K, "landmark", K, N * MinSub, 1#
        AddPoint Ids, Cons, Bits, Xs, Scales, Count, _
                 "g_s13_k" & K, "transform_stress", K, _
                 (3# * N + 1#) * MinSub, 3#
        AddPoint Ids, Cons, Bits, Xs, Scales, Count, _
                 "g_s12_k" & K, "transform_stress", K, _
                 (2# * N + 1#) * MinSub, 2#
        AddPoint Ids, Cons, Bits, Xs, Scales, Count, _
                 "g_s23_k" & K, "transform_stress", K, _
                 (3# * N + 2#) * MinSub, 3#
    Next B
'------------------------------------------------------------------------------
' HARD UNDERFLOW
'------------------------------------------------------------------------------
    'A positive mathematical ratio that stores as zero. Declared bucket 0, and
    'classified apart by the analyzer: the value is not merely
    'imprecise, it has
    'become indistinguishable from the support boundary
        AddPoint Ids, Cons, Bits, Xs, Scales, Count, _
                 "g_underflow_minsub", "underflow", 0, MinSub, 4#
        AddPoint Ids, Cons, Bits, Xs, Scales, Count, _
                 "g_underflow_wide", "underflow", 0, 1E-300, 1E+300
        AddPoint Ids, Cons, Bits, Xs, Scales, Count, _
                 "g_underflow_deep", "underflow", 0, 1E-200, 1E+200
'------------------------------------------------------------------------------
' DECIMAL TWINS
'------------------------------------------------------------------------------
    'Diagnostic only. If a twin's echoed X differs from its constructed
    'partner's, the VBE parser altered the value and every number in this file
    'is suspect. The analyzer pairs them by stripping the twin_ prefix
        AddPoint Ids, Cons, Bits, Xs, Scales, Count, _
                 "twin_g_landmark_k8", "decimal_twin", 8, _
                 6.32404026676796E-322, 1#
        AddPoint Ids, Cons, Bits, Xs, Scales, Count, _
                 "twin_g_landmark_k32", "decimal_twin", 32, _
                 1.06099789548264E-314, 1#
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
'   leave zero behind. See CONTRIBUTING, "Force the rounding boundary".
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
    Dim StandardX           As Double          'X / ScaleParam, as stored
    Dim LogStandardX        As Double          'Log(X) - Log(ScaleParam)
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
        HasStandard = PROB_TryDivide(X, ScaleParam, StandardX)

        On Error Resume Next
        Err.Clear
        LogStandardX = Log(X) - Log(ScaleParam)
        If Err.Number <> 0 Then LogStandardX = 0#
        Err.Clear
        On Error GoTo 0

        Row = "gamma," & Surface & "," & Id & "," & Construction & "," & _
              ShapeId(Shape) & "," & BucketBits & "," & _
              FormatFullPrecision(X) & "," & _
              FormatFullPrecision(Shape) & "," & _
              FormatFullPrecision(ScaleParam) & "," & _
              IIf(HasStandard, FormatFullPrecision(StandardX), "") & "," & _
              FormatFullPrecision(LogStandardX)
'------------------------------------------------------------------------------
' CURRENT PATH
'------------------------------------------------------------------------------
        On Error Resume Next
        Err.Clear
        Select Case Surface
            Case "density"
                Current = K_STATS_Gamma_Density(X, Shape, ScaleParam)
            Case "cumulative"
                Current = K_STATS_Gamma_Cumulative(X, Shape, ScaleParam)
            Case Else
                Current = K_STATS_Gamma_Survival(X, Shape, ScaleParam)
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
    Dim LogP                As Double          'Log of the lower tail
    Dim LogGamma1pShape     As Double          'LogGamma(1 + Shape)
    Dim FailMsg             As String          'Kernel diagnostic
    Dim P                   As Double          'Lower tail probability
    Dim LogDensity          As Double          'Log of the density
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

        If Not PROB_TryLogGamma1p(Shape, LogGamma1pShape, FailMsg) Then
            Exit Function
        End If

        Select Case Surface

            Case "density"
    'Log f = (Shape - 1) * Log(z) - z - LogGamma(Shape) - Log(ScaleParam)
                LogDensity = (Shape - 1#) * LogStandardX _
                             - StandardX _
                             - PROB_LogGamma(Shape) _
                             - Log(ScaleParam)
                If Not PROB_TryExp(LogDensity, Result) Then Exit Function

            Case "cumulative"
                LogP = Shape * LogStandardX - LogGamma1pShape
                If Not PROB_TryExp(LogP, Result) Then Exit Function

            Case Else
    'Q = -Expm1(Log P): never 1 - Exp(Log P), which cancels when P is tiny
                LogP = Shape * LogStandardX - LogGamma1pShape
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
        InitialFileName:="gamma_probe.csv", _
        FileFilter:="Probe output (*.csv),*.csv", _
        Title:="Save the Gamma positive-ratio probe")

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
