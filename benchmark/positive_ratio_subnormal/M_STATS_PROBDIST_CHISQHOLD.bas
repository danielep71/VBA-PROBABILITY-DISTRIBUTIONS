Attribute VB_Name = "M_STATS_PROBDIST_CHISQHOLD"
Option Explicit

'Column count of the committed schema, checked against every row written
Private Const PROBE_FIELDS As Long = 14

'Shape is part of the experimental design, not a setting: each surface has a
'different conditioning problem and one shape cannot serve all three. At
'Shape is part of the experimental design, not a setting: each surface has
'a different conditioning problem and one shape cannot serve all three. At
'Shape = 0.5 the survival is already correctly saturated at 1.0 across the
'whole subnormal ladder and cannot constrain dispatch. The slices are
'therefore surface-specific; shape_id is provenance, echo_shape is
'authoritative.
'==============================================================================
' M_STATS_PROBDIST_CHISQPROBE
'------------------------------------------------------------------------------
' PURPOSE
'   Independent HOLDOUT exporter for ICR-P1-01A (#13), Chi-square arm.
'
'   Mechanical from the frozen specification in HOLDOUT_DESIGN.md, which was
'   committed before any holdout output existed. Nothing here was chosen after
'   seeing a measured error.
'
'   It tests two preregistered claims from the A2 fitting arm:
'
'       ChiSquare density    candidate dominates at <= 48 bits
'       ChiSquare cumulative candidate dominates at <= 48 bits
'
'   Survival carries no claim and is characterised only: the A1 landmark and
'   stress decomposition showed two superimposed mechanisms, and separating
'   them is unfinished.
'
'   A holdout indicating some other bucket is a REJECTION of the fitted
'   cutoff, not a replacement for it. Adopting its crossover would make it a
'   second fitting set and leave no independent evidence at all.
'
' WHY THIS IS A SEPARATE EXPORTER
'   Chi-square does NOT standardize the way Gamma does. All three public
'   surfaces compute 0.5 * X inline, with no Try wrapper:
'
'       PROB_TryGammaRegularizedP(0.5 * DegreesFreedom, 0.5 * X, ...)
'
'   Gamma divides through PROB_TryDivide. That difference is structural, not
'   cosmetic, and it changes both the reachability and the rounding:
'
'     - Rounding is BINARY, not graded. For a binary64 X, halving is either
'       exact or lands on a midpoint. The 1/3 and 2/3 ULP stresses used in the
'       Gamma arm are physically unconstructible here: aiming at a 1/3 ULP
'       quotient needs an X that is not representable, and rounding it lands on
'       exactly 1/2.
'
'     - Reachability is far narrower. One free parameter instead of two, so the
'       only route to a tiny standardized argument is a tiny X. Hard underflow
'       needs X below 2 ^ -1073, where Gamma reaches it with any quotient that
'       happens to underflow -- including operands as ordinary as
'       1E-300 / 1E+300.
'
'   Nothing measured in the Gamma arm carries over. Its 48-bit and 40-bit
'   figures are family-specific measurements, not priors.
'
' A LOWER GUARD GAMMA DOES NOT HAVE
'   PROB_TF_ValidateDF rejects a degrees-of-freedom whose half underflows:
'   "0.5 * DegreesFreedom <= 0". Gamma's shape guard has no lower cutoff at
'   all. Every df slice below is far above that boundary.
'
' EVIDENCE CLASSES
'   landmark          X = 2N * m with N = 2 ^ (k - 1) and m = 2 ^ -1074.
'                     X / 2 is exactly the k-bit landmark; no transform error.
'   transform_stress  X = (2N + 1) * m. X / 2 falls exactly on the midpoint
'                     between two representable subnormals -- the only
'                     transform error this family can produce.
'   underflow         X = m, so the mathematical X / 2 is 2 ^ -1075 and the
'                     stored value is zero. Classified apart, never bucket 0.
'   decimal_twin      diagnostic only, to detect a parser difference.
'
' DEGREES OF FREEDOM, NOT SHAPE
'   The kernel receives df / 2. The slices below are chosen so the kernel
'   shapes match the Gamma arm exactly, which is what makes the two families
'   comparable at all. shape_id records the DF for provenance; echo_df is
'   authoritative. Non-integer df is legal and deliberate here.
'
' OUTPUT
'   The committed schema with echo_df in place of echo_shape, and no
'   echo_scale: this family always standardizes by 2, and the analyzer
'   supplies that rather than trusting an echoed constant.
'
' USAGE
'   Run Probe_ChiSquareHoldout, then validate with
'   analyze_positive_ratio_subnormal.py before committing the CSV.
'
' DEPENDENCIES
'   - K_STATS_ChiSquare_Density, _Cumulative, _Survival
'                                          (M_STATS_PROBDIST_TFAMILY)
'   - PROB_TryLogGamma1p, PROB_LogGamma    (M_STATS_PROBDIST_SPECIALFUNCS)
'   - PROB_TryExp, PROB_Expm1              (M_STATS_PROBDIST_CORE)
'
' UPDATED
'   2026-08-14
'==============================================================================


Public Sub Probe_ChiSquareHoldout()
'
'==============================================================================
' Probe_ChiSquareHoldout
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

    MsgBox "Chi-square positive-ratio HOLDOUT complete: " & Count & _
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
    'df = 2 * kernel shape, preserving comparability with the Gamma arm.
    'Five values admitted by min(P, Q) > 2 ^ -20 on the exact reference at
    'every holdout bucket - a conditioning criterion evaluated before any
    'implementation runs, so it cannot encode which branch wins.
            ShapeSlicesFor = Array(0.00073575888234288, _
                                   0.0014142135623731, _
                                   0.0031415926535898, _
                                   0.0145947051386, _
                                   0.028284271247462)
        Case Else
    'df = 2 * the eight irrational-derived kernel shapes: alpha, 1/1000e,
    'sqrt2/100, ln2/10, pi/10, phi - 1, sqrt2/2, e/3. The frozen binary64 is
    'authoritative; the expressions record only how each was selected.
            ShapeSlicesFor = Array(0.0145947051386, _
                                   0.00073575888234288, _
                                   0.028284271247462, _
                                   0.13862943611199, _
                                   0.62831853071796, _
                                   1.23606797749978, _
                                   1.4142135623731, _
                                   1.81218788521484)
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
'   Constructs the point set: fifteen buckets from 52 bits down to 1, each
'   with a landmark and ONE transform stress, plus the single hard-underflow
'   case and two decimal twins.
'
' THE CONSTRUCTION
'   With m = 2 ^ -1074 the subnormal ULP and N = 2 ^ (k - 1):
'
'       landmark   X = 2N * m        X / 2 is exactly the k-bit landmark
'       1/2 ULP    X = (2N + 1) * m  X / 2 falls on the midpoint
'
'   Two classes, not four. Halving a binary64 is either exact or lands on a
'   midpoint, so the 1/3 and 2/3 ULP offsets used in the Gamma arm cannot be
'   produced here: aiming at one requires an X that is not representable, and
'   rounding it lands back on exactly 1/2. The analyzer asserts this.
'
'   Every X above is exactly representable, verified at all fifteen buckets.
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
    Buckets = Array(50, 46, 41, 37, 33, 29, 25, 21, 17, 13, 9, 5, 3)
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

    'X = 2N * m, so X / 2 is exactly the K-bit landmark and no transform
    'error is introduced. Isolates downstream behaviour.
        AddPoint Ids, Cons, Bits, Xs, Scales, Count, _
                 "ch_landmark_k" & K, "landmark", K, 2# * N * MinSub, 2#

    'X = (2N + 1) * m, so X / 2 falls exactly on the midpoint between two
    'representable subnormals. This is the ONLY transform error a halving
    'can produce: the 1/3 and 2/3 ULP offsets used in the Gamma arm require
    'an X that is not representable and would round back onto this same
    'midpoint.
        AddPoint Ids, Cons, Bits, Xs, Scales, Count, _
                 "ch_tie_k" & K, "transform_stress", K, _
                 (2# * N + 1#) * MinSub, 2#
    Next B
'------------------------------------------------------------------------------
' HARD UNDERFLOW
'------------------------------------------------------------------------------
    'The only route in this family: X must itself be the smallest positive
    'Double, so the mathematical X / 2 is 2 ^ -1075 and the stored value is
    'zero. Contrast Gamma, which reaches the same state from operands as
    'ordinary as 1E-300 and 1E+300.
        AddPoint Ids, Cons, Bits, Xs, Scales, Count, _
                 "ch_underflow_minsub", "underflow", 0, MinSub, 2#
'------------------------------------------------------------------------------
' DECIMAL TWINS
'------------------------------------------------------------------------------
    'Diagnostic only. If a twin's echoed X differs from its constructed
    'partner's, the VBE parser altered the value and every number in this file
    'is suspect. The analyzer pairs them by stripping the twin_ prefix
        AddPoint Ids, Cons, Bits, Xs, Scales, Count, _
                 "twin_ch_landmark_k9", "decimal_twin", 9, _
                 2.52961610670718E-321, 2#
        AddPoint Ids, Cons, Bits, Xs, Scales, Count, _
                 "twin_ch_landmark_k33", "decimal_twin", 33, _
                 4.24399158193054E-314, 2#
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
        InitialFileName:="chisquare_holdout.csv", _
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
