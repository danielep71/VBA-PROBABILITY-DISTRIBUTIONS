Attribute VB_Name = "M_STATS_PROBDIST_GCP"
Option Explicit

'One Double and its two 32-bit words, to compare two results bit for bit
Private Type GCP_Words
    Low                     As Long
    High                    As Long
End Type

Private Type GCP_Number
    Value                   As Double
End Type

'The parts a row can name
Private Const GCP_EMPTY                 As Long = 1
Private Const GCP_PUBLIC_CDF            As Long = 2
Private Const GCP_PUBLIC_SF             As Long = 3
Private Const GCP_ROUTED_P              As Long = 4
Private Const GCP_ROUTED_Q              As Long = 5
Private Const GCP_B64_KERNEL            As Long = 6
Private Const GCP_LEGACY_LOOP           As Long = 7
Private Const GCP_KERNEL_P              As Long = 8
Private Const GCP_KERNEL_Q              As Long = 9
Private Const GCP_LOG_PDF               As Long = 10
Private Const GCP_STIRLING              As Long = 11
Private Const GCP_SET_STATUS            As Long = 12
Private Const GCP_TRY_DIVIDE            As Long = 13
Private Const GCP_CLASSIFY              As Long = 14
Private Const GCP_IS_FINITE             As Long = 15
Private Const GCP_RUN_POW2              As Long = 16
Private Const GCP_RUN_FACTOR            As Long = 17
Private Const GCP_RUN_SMALL_LG1P        As Long = 18
Private Const GCP_RUN_VALIDATE          As Long = 19
Private Const GCP_PUBLIC_INV            As Long = 20
Private Const GCP_INV_P                 As Long = 21

'Private procedures of the production modules, reached through
'Application.Run. GCP_RUN_POW2 measures what that call itself costs
Private Const GCP_NAME_POW2             As String = "M_STATS_PROBDIST_SPECIALFUNCS.PROB_B64_Pow2"
Private Const GCP_NAME_FACTOR           As String = "M_STATS_PROBDIST_SPECIALFUNCS.PROB_B64_Factor"
Private Const GCP_NAME_SMALL_LG1P       As String = "M_STATS_PROBDIST_SPECIALFUNCS.PROB_GammaSmallLogGamma1p"
Private Const GCP_NAME_VALIDATE         As String = "M_STATS_PROBDIST_CONTINUOUS.PROB_CN_ValidateXShapeScale"

'Each timing repeats its call until this many seconds have elapsed
Private Const GCP_MIN_SECONDS           As Double = 0.5
Private Const GCP_TAIL_SECONDS          As Double = 0.1

'Copies of private constants of M_STATS_PROBDIST_SPECIALFUNCS, which the
'solver replica and the route test need. test_profile.py fails when a copy
'differs from its source
Private Const GCP_INV_MAX_ITER          As Long = 200
Private Const GCP_FAST_MIN_SHAPE        As Double = 0.01
Private Const GCP_FAST_MAX_SHAPE        As Double = 1E+15
Private Const GCP_FAST_FLOOR            As Double = 1E-290
Private Const GCP_FAST_LOWER_MAX        As Double = 0.9
Private Const GCP_FAST_UPPER_MAX        As Double = 0.5

'==============================================================================
' M_STATS_PROBDIST_GCP
'------------------------------------------------------------------------------
' PURPOSE
'   Export macro for the profile of the public Gamma call and of the Gamma
'   inverse solver (#348, benchmark/gamma_call_profile/README.md). It times
'   the layers of one call separately and counts what the solver evaluates.
'   It changes no production source and asserts nothing.
'
' USAGE
'   Export_GammaCallProfile fills gamma_call_profile.csv. It finds the file
'   beside the workbook, or asks for it; ci/Export-ExcelGrids.ps1 runs it on
'   the self-hosted runner with the file staged beside the export workbook.
'
' GRID FORMAT (study grid; 8 columns)
'   point_id, kind, part, a, x, calls, us_per_call, note
'
'   The macro writes the last three. On a row of kind call or fixed, part
'   names one routine, calls is the number of repetitions timed and note is
'   ok=1 when the routine returned a value. On a row of kind inverse, x holds
'   the probability; part solver replays the solver: calls is its number of
'   iterations, us_per_call the summed time of its tail evaluations, and
'   note says whether the replay returned the Double of PROB_TryGammaInvP
'   and which routine served each evaluation.
'
' THE SOLVER REPLICA
'   ReplicaInverse is PROB_TryGammaInvP written again from public routines,
'   recording every iterate. It is trusted only where it returns the same
'   Double as the production routine: same=1 in the note.
'
' ERROR POLICY
'   A part that raises is recorded with ok=0 and a time of zero. Reaching a
'   Private procedure through Application.Run may be refused by the host; the
'   row then says so and the analysis leaves that part out.
'
' DEPENDENCIES
'   - K_STATS_Gamma_Cumulative, K_STATS_Gamma_Survival,
'     K_STATS_Gamma_InverseCumulative       (M_STATS_PROBDIST_CONTINUOUS)
'   - PROB_TryGammaRoutedP/Q, PROB_TryGammaB64SeriesP,
'     PROB_TryGammaB64ContinuedFractionQ, PROB_TryGammaSeriesP,
'     PROB_TryGammaContinuedFractionQ, PROB_TryGammaRegularizedP/Q,
'     PROB_TryGammaLogPdf, PROB_TryGammaInvP, PROB_StirlingError,
'     PROB_ClassifyPositiveRatio, PROB_LogGamma
'                                           (M_STATS_PROBDIST_SPECIALFUNCS)
'   - PROB_SetStatus, PROB_TryDivide, PROB_IsFinite, PROB_TryExp,
'     PROB_NormalInvCDFRaw                  (M_STATS_PROBDIST_CORE)
'
' UPDATED
'   2026-11-16
'==============================================================================


Public Sub Export_GammaCallProfile()
'
'==============================================================================
' Export_GammaCallProfile
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the measurement columns of gamma_call_profile.csv.
'==============================================================================
'
    Dim Path                As String          'File written
    Dim Filled              As Long            'Rows written, or -1

    Filled = ExportFile("gamma_call_profile.csv", "Gamma call profile", Path)
    If Filled < 0 Then Exit Sub

    MsgBox "Gamma call profile complete: " & Filled & _
           " observation(s) written to" & vbCrLf & Path, _
           vbInformation, "Gamma call profile"
End Sub


Private Function ExportFile( _
    ByVal FileName As String, _
    ByVal Title As String, _
    ByRef Path As String) _
    As Long
'
'==============================================================================
' ExportFile
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the measurement columns of the grid file, row by row. Returns the
'   number of rows written, or -1 when the file was not chosen or the export
'   failed; the failure is reported here, the success by the caller.
'
' BEHAVIOR
'   - Reads the whole file, measures every data row, rewrites the file.
'   - A row with fewer than 8 fields is left as it is and not counted.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Lines()             As String          'File lines
    Dim Raw                 As String          'File contents
    Dim FileNum             As Integer         'Input file handle
    Dim OutNum              As Integer         'Output file handle
    Dim Cols                As Variant         'Split fields of one row
    Dim A                   As Double          'Shape
    Dim X                   As Double          'Argument, or probability
    Dim Calls               As Long            'Repetitions, or iterations
    Dim Us                  As Double          'Microseconds
    Dim Note                As String          'Third measurement column
    Dim OK                  As Boolean         'TRUE when the part returned a value
    Dim Filled              As Long            'Rows written
    Dim I                   As Long            'Row index
'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

        Path = ResolveGridPath(FileName, Title)
        ExportFile = -1
        If Len(Path) = 0 Then Exit Function     'User cancelled the picker

        Filled = 0

        FileNum = FreeFile
        Open Path For Input As #FileNum
        Raw = Input$(LOF(FileNum), FileNum)
        Close #FileNum

        Raw = Replace(Raw, vbCrLf, vbLf)
        Raw = Replace(Raw, vbCr, vbLf)
        Lines = Split(Raw, vbLf)
'------------------------------------------------------------------------------
' MEASURE
'------------------------------------------------------------------------------
    'Row 0 is the header; data starts at row 1
        For I = 1 To UBound(Lines)
            If Len(Trim$(Lines(I))) = 0 Then GoTo ContinueRow

            Cols = Split(Lines(I), ",")
            If UBound(Cols) < 7 Then GoTo ContinueRow

            A = Val(CStr(Cols(3)))
            X = Val(CStr(Cols(4)))

            If CStr(Cols(2)) = "solver" Then
                SolverRow A, X, Calls, Us, Note
            Else
                Us = TimeCode(PartCode(CStr(Cols(2))), A, X, GCP_MIN_SECONDS, Calls, OK)
                Note = "ok=0"
                If OK Then Note = "ok=1"
            End If

            Cols(5) = CStr(Calls)
            Cols(6) = Micro(Us)
            Cols(7) = Note

            Lines(I) = Join(Cols, ",")
            Filled = Filled + 1
ContinueRow:
        Next I
'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        OutNum = FreeFile
        Open Path For Output As #OutNum
        For I = 0 To UBound(Lines)
            If I < UBound(Lines) Or Len(Lines(I)) > 0 Then Print #OutNum, Lines(I)
        Next I
        Close #OutNum

    ExportFile = Filled
    Exit Function

Err_Handler:
    On Error Resume Next
    Close #FileNum
    Close #OutNum
    MsgBox Title & " failed: " & Err.Description, _
           vbExclamation, Title
    ExportFile = -1
End Function


Private Function PartCode(ByVal Part As String) As Long
'
'==============================================================================
' PartCode
'------------------------------------------------------------------------------
' PURPOSE
'   The code of a part name, or zero for a name this module does not know.
'==============================================================================
'
    Select Case Part
        Case "empty": PartCode = GCP_EMPTY
        Case "public_cdf": PartCode = GCP_PUBLIC_CDF
        Case "public_sf": PartCode = GCP_PUBLIC_SF
        Case "routed_p": PartCode = GCP_ROUTED_P
        Case "routed_q": PartCode = GCP_ROUTED_Q
        Case "b64_kernel": PartCode = GCP_B64_KERNEL
        Case "legacy_loop": PartCode = GCP_LEGACY_LOOP
        Case "kernel_p": PartCode = GCP_KERNEL_P
        Case "kernel_q": PartCode = GCP_KERNEL_Q
        Case "log_pdf": PartCode = GCP_LOG_PDF
        Case "stirling": PartCode = GCP_STIRLING
        Case "set_status": PartCode = GCP_SET_STATUS
        Case "try_divide": PartCode = GCP_TRY_DIVIDE
        Case "classify": PartCode = GCP_CLASSIFY
        Case "is_finite": PartCode = GCP_IS_FINITE
        Case "run_pow2": PartCode = GCP_RUN_POW2
        Case "run_factor": PartCode = GCP_RUN_FACTOR
        Case "run_small_lg1p": PartCode = GCP_RUN_SMALL_LG1P
        Case "run_validate": PartCode = GCP_RUN_VALIDATE
        Case "public_inv": PartCode = GCP_PUBLIC_INV
        Case "inv_p": PartCode = GCP_INV_P
    End Select
End Function


Private Function TimeCode( _
    ByVal Code As Long, _
    ByVal A As Double, _
    ByVal X As Double, _
    ByVal MinSeconds As Double, _
    ByRef Calls As Long, _
    ByRef OK As Boolean) _
    As Double
'
'==============================================================================
' TimeCode
'------------------------------------------------------------------------------
' PURPOSE
'   Microseconds per call of one part at (A, X). The call is repeated, the
'   count doubling, until MinSeconds have elapsed; one call comes first, so
'   a part that takes a second is not repeated.
'
' NOTES
'   Every part is timed through RunPart, so each figure includes the cost of
'   that call and of its Select Case. The part empty is that cost alone; the
'   analysis subtracts it.
'==============================================================================
'
    Dim N                   As Long            'Repetitions in this round
    Dim K                   As Long            'Repetition index
    Dim T0                  As Double          'Start time
    Dim Elapsed             As Double          'Seconds of this round

    Calls = 0
    OK = False
    If Code = 0 Then Exit Function

    N = 1
    Do
        T0 = Timer
        For K = 1 To N
            OK = RunPart(Code, A, X)
        Next K
        Elapsed = Timer - T0
        If Elapsed < 0# Then Elapsed = Elapsed + 86400#
        If Not OK Then Exit Function
        If Elapsed >= MinSeconds Or N >= 536870912 Then Exit Do
        N = N * 2
    Loop

    Calls = N
    TimeCode = Elapsed / N * 1000000#
End Function


Private Function RunPart( _
    ByVal Code As Long, _
    ByVal A As Double, _
    ByVal X As Double) _
    As Boolean
'
'==============================================================================
' RunPart
'------------------------------------------------------------------------------
' PURPOSE
'   Calls one part once. TRUE when it returned a value; FALSE when it
'   declined or raised.
'==============================================================================
'
    Dim V                   As Double          'Numeric result
    Dim M                   As String          'Diagnostic text
    Dim R                   As Variant         'Result of a public or Run call

    On Error GoTo Failed

    Select Case Code
        Case GCP_EMPTY
            RunPart = True
        Case GCP_PUBLIC_CDF
            R = K_STATS_Gamma_Cumulative(X, A, 1#)
            RunPart = Not IsError(R)
        Case GCP_PUBLIC_SF
            R = K_STATS_Gamma_Survival(X, A, 1#)
            RunPart = Not IsError(R)
        Case GCP_ROUTED_P
            RunPart = PROB_TryGammaRoutedP(A, X, V, M)
        Case GCP_ROUTED_Q
            RunPart = PROB_TryGammaRoutedQ(A, X, V, M)
        Case GCP_B64_KERNEL
            If X < A + 1# Then
                RunPart = PROB_TryGammaB64SeriesP(A, X, V, M)
            Else
                RunPart = PROB_TryGammaB64ContinuedFractionQ(A, X, V, M)
            End If
        Case GCP_LEGACY_LOOP
            If X < A + 1# Then
                RunPart = PROB_TryGammaSeriesP(A, X, V, M)
            Else
                RunPart = PROB_TryGammaContinuedFractionQ(A, X, V, M)
            End If
        Case GCP_KERNEL_P
            RunPart = PROB_TryGammaRegularizedP(A, X, V, M)
        Case GCP_KERNEL_Q
            RunPart = PROB_TryGammaRegularizedQ(A, X, V, M)
        Case GCP_LOG_PDF
            RunPart = PROB_TryGammaLogPdf(X, A, Log(X), 0#, V, M)
        Case GCP_STIRLING
            V = PROB_StirlingError(A)
            RunPart = True
        Case GCP_SET_STATUS
            PROB_SetStatus M, vbNullString
            RunPart = True
        Case GCP_TRY_DIVIDE
            RunPart = PROB_TryDivide(X, 1#, V)
        Case GCP_CLASSIFY
            V = PROB_ClassifyPositiveRatio(True, X, PROB_PRS_CUTOFF_40BIT)
            RunPart = True
        Case GCP_IS_FINITE
            RunPart = PROB_IsFinite(X)
        Case GCP_RUN_POW2
            R = Application.Run(GCP_NAME_POW2, 971)
            RunPart = Not IsError(R)
        Case GCP_RUN_FACTOR
            R = Application.Run(GCP_NAME_FACTOR, A, X)
            RunPart = Not IsError(R)
        Case GCP_RUN_SMALL_LG1P
            R = Application.Run(GCP_NAME_SMALL_LG1P, A)
            RunPart = Not IsError(R)
        Case GCP_RUN_VALIDATE
            R = Application.Run(GCP_NAME_VALIDATE, X, A, 1#, M, "Shape", "ScaleParam")
            RunPart = CBool(R)
        Case GCP_PUBLIC_INV
            R = K_STATS_Gamma_InverseCumulative(X, A, 1#)
            RunPart = Not IsError(R)
        Case GCP_INV_P
            RunPart = PROB_TryGammaInvP(X, 1# - X, A, V, M)
    End Select
    Exit Function

Failed:
    RunPart = False
End Function


Private Sub SolverRow( _
    ByVal A As Double, _
    ByVal P As Double, _
    ByRef Calls As Long, _
    ByRef Us As Double, _
    ByRef Note As String)
'
'==============================================================================
' SolverRow
'------------------------------------------------------------------------------
' PURPOSE
'   Replays the Gamma inverse solver at (P, A) and reports its iterations,
'   the routine that served each tail evaluation and the time those
'   evaluations and the density evaluations take.
'==============================================================================
'
    Dim Xs(1 To GCP_INV_MAX_ITER) As Double    'Iterates, in order
    Dim Count               As Long            'Number of iterates
    Dim Lower               As Boolean         'TRUE when the lower tail is matched
    Dim Replica             As Double          'Quantile of the replica
    Dim Production          As Double          'Quantile of PROB_TryGammaInvP
    Dim M                   As String          'Diagnostic text
    Dim Same                As Boolean         'TRUE when the two are one Double
    Dim Code                As Long            'Tail routine the solver calls
    Dim SeriesCount         As Long            'Evaluations served by the binary64 series
    Dim FractionCount       As Long            'Evaluations served by the binary64 fraction
    Dim KernelCount         As Long            'Evaluations served by the double-double kernel
    Dim TailUs              As Double          'Summed time of the tail evaluations
    Dim PdfUs               As Double          'Summed time of the density evaluations
    Dim N                   As Long            'Repetitions of one timing, unused
    Dim OK                  As Boolean         'Result of one timing, unused
    Dim I                   As Long            'Iterate index

    Calls = 0
    Us = 0#
    Note = "same=0"

    On Error GoTo Failed

    Replica = ReplicaInverse(P, 1# - P, A, Xs, Count, Lower)
    If Count = 0 Then Exit Sub

    If PROB_TryGammaInvP(P, 1# - P, A, Production, M) Then
        Same = (ToBits(Production) = ToBits(Replica))
    End If

    If A >= GCP_FAST_MIN_SHAPE Then
        Code = GCP_ROUTED_Q
        If Lower Then Code = GCP_ROUTED_P
    Else
        Code = GCP_KERNEL_Q
        If Lower Then Code = GCP_KERNEL_P
    End If

    For I = 1 To Count
        Select Case TailSource(A, Xs(I), Lower)
            Case 1: SeriesCount = SeriesCount + 1
            Case 2: FractionCount = FractionCount + 1
            Case Else: KernelCount = KernelCount + 1
        End Select
        TailUs = TailUs + TimeCode(Code, A, Xs(I), GCP_TAIL_SECONDS, N, OK)
        PdfUs = PdfUs + TimeCode(GCP_LOG_PDF, A, Xs(I), GCP_TAIL_SECONDS, N, OK)
    Next I

    Calls = Count
    Us = TailUs
    Note = "same=0"
    If Same Then Note = "same=1"
    Note = Note & ";lower=0"
    If Lower Then Note = Left$(Note, Len(Note) - 1) & "1"
    Note = Note & ";series=" & CStr(SeriesCount) & ";fraction=" & CStr(FractionCount) & _
           ";kernel=" & CStr(KernelCount) & ";pdf_us=" & Micro(PdfUs)
    Exit Sub

Failed:
    Calls = 0
    Us = 0#
    Note = "same=0;error"
End Sub


Private Function ReplicaInverse( _
    ByVal Probability As Double, _
    ByVal ComplementProbability As Double, _
    ByVal A As Double, _
    ByRef Xs() As Double, _
    ByRef Count As Long, _
    ByRef SolveLower As Boolean) _
    As Double
'
'==============================================================================
' ReplicaInverse
'------------------------------------------------------------------------------
' PURPOSE
'   PROB_TryGammaInvP for 0 < Probability < 1, statement for statement, with
'   every iterate at which a tail is evaluated stored in Xs. Count is zero
'   when the routine would have failed.
'==============================================================================
'
    Dim Target              As Double          'Probability mass being matched
    Dim Z                   As Double          'Normal seed
    Dim T                   As Double          'Wilson-Hilferty working value
    Dim X                   As Double          'Current iterate
    Dim XNew                As Double          'Proposed iterate
    Dim Low                 As Double          'Bracket lower bound
    Dim High                As Double          'Bracket upper bound
    Dim HasHigh             As Boolean         'TRUE once an upper bound exists
    Dim Value               As Double          'Tail at the iterate
    Dim Residual            As Double          'Signed distance from the target
    Dim LogDensity          As Double          'Log of the density at the iterate
    Dim Density             As Double          'Density at the iterate
    Dim FailMsg             As String          'Diagnostic text
    Dim Stored              As Long            'Iterates stored
    Dim IterIdx             As Long            'Iteration index

    Count = 0
    If Probability <= 0# Or ComplementProbability <= 0# Then Exit Function

    If Probability <= ComplementProbability Then
        SolveLower = True
        Target = Probability
    Else
        SolveLower = False
        Target = ComplementProbability
    End If

    If Probability <= 0.5 Then
        Z = PROB_NormalInvCDFRaw(Probability)
    Else
        Z = -PROB_NormalInvCDFRaw(ComplementProbability)
    End If

    T = 1# - 1# / (9# * A) + Z / Sqr(9# * A)
    X = A * T * T * T

    If A < 1# Or X <= 0# Or Not PROB_IsFinite(X) Then
        If SolveLower Then
            If Not PROB_TryExp((Log(Probability) + PROB_LogGamma(A + 1#)) / A, X) Then X = A
        Else
            X = A
        End If
    End If

    If X <= 0# Or Not PROB_IsFinite(X) Then X = 0.00000001

    Low = 0#
    High = 0#
    HasHigh = False

    For IterIdx = 1 To GCP_INV_MAX_ITER
        Stored = Stored + 1
        Xs(Stored) = X

        If SolveLower Then
            If A >= GCP_FAST_MIN_SHAPE Then
                If Not PROB_TryGammaRoutedP(A, X, Value, FailMsg) Then Exit Function
            Else
                If Not PROB_TryGammaRegularizedP(A, X, Value, FailMsg) Then Exit Function
            End If
            Residual = Value - Target
        Else
            If A >= GCP_FAST_MIN_SHAPE Then
                If Not PROB_TryGammaRoutedQ(A, X, Value, FailMsg) Then Exit Function
            Else
                If Not PROB_TryGammaRegularizedQ(A, X, Value, FailMsg) Then Exit Function
            End If
            Residual = Target - Value
        End If

        If Residual < 0# Then
            If X > Low Then Low = X
        Else
            If (Not HasHigh) Or X < High Then High = X
            HasHigh = True
        End If

        If Not PROB_TryGammaLogPdf(X, A, Log(X), 0#, LogDensity, FailMsg) Then
            Density = 0#
        ElseIf Not PROB_TryExp(LogDensity, Density) Then
            Density = 0#
        End If

        If Density <= 0# Then
            If HasHigh Then
                XNew = 0.5 * (Low + High)
            Else
                XNew = 2# * X
            End If
        Else
            XNew = X - Residual / Density
            If HasHigh Then
                If XNew <= Low Or XNew >= High Then XNew = 0.5 * (Low + High)
            ElseIf XNew <= Low Then
                XNew = 2# * X
            End If
        End If

        If Abs(XNew - X) <= PROB_EPS * Abs(XNew) Or XNew = X Then
            Count = Stored
            ReplicaInverse = XNew
            Exit Function
        End If

        X = XNew
    Next IterIdx
End Function


Private Function TailSource( _
    ByVal A As Double, _
    ByVal X As Double, _
    ByVal Lower As Boolean) _
    As Long
'
'==============================================================================
' TailSource
'------------------------------------------------------------------------------
' PURPOSE
'   Which routine serves the solver's tail evaluation at (A, X): 1 the
'   binary64 series, 2 the binary64 continued fraction, 3 the double-double
'   kernel. It applies the tests of PROB_TryGammaRoutedP and
'   PROB_TryGammaRoutedQ to the same binary64 value.
'==============================================================================
'
    Dim Fast                As Double          'Binary64 tail
    Dim Probe               As String          'Diagnostic of a declined routine

    TailSource = 3
    If A < GCP_FAST_MIN_SHAPE Then Exit Function
    If Not (PROB_IsFinite(A) And PROB_IsFinite(X)) Then Exit Function
    If Not (A > 0# And A <= GCP_FAST_MAX_SHAPE And X > 0#) Then Exit Function

    If X < A + 1# Then
        If Not PROB_TryGammaB64SeriesP(A, X, Fast, Probe) Then Exit Function
        If Lower Then
            If Fast >= GCP_FAST_FLOOR And Fast <= 1# Then TailSource = 1
        Else
            If Fast >= 0# And Fast <= GCP_FAST_LOWER_MAX Then TailSource = 1
        End If
    Else
        If Not PROB_TryGammaB64ContinuedFractionQ(A, X, Fast, Probe) Then Exit Function
        If Lower Then
            If Fast >= 0# And Fast <= GCP_FAST_UPPER_MAX Then TailSource = 2
        Else
            If Fast >= GCP_FAST_FLOOR And Fast <= GCP_FAST_UPPER_MAX Then TailSource = 2
        End If
    End If
End Function


Private Function Micro(ByVal Us As Double) As String
'
'==============================================================================
' Micro
'------------------------------------------------------------------------------
' PURPOSE
'   A time in microseconds as text with three decimals and a point,
'   whatever the decimal separator of the host.
'==============================================================================
'
    Micro = Replace(Format$(Us, "0.000"), ",", ".")
End Function


Private Function ToBits(ByVal X As Double) As String
'
'==============================================================================
' ToBits
'------------------------------------------------------------------------------
' PURPOSE
'   The 64 bits of X as 16 hexadecimal digits, most significant first.
'==============================================================================
'
    Dim Words               As GCP_Words       'The two halves
    Dim Number              As GCP_Number      'The same eight bytes as a Double

    Number.Value = X
    LSet Words = Number
    ToBits = Right$("00000000" & Hex$(Words.High), 8) & _
             Right$("00000000" & Hex$(Words.Low), 8)
End Function


Private Function ResolveGridPath( _
    ByVal FileName As String, _
    ByVal Title As String) _
    As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Finds the grid file beside the workbook, or asks for it.
'==============================================================================
'
    Dim BookPath            As String          'Workbook folder
    Dim Candidate           As String          'Path next to the workbook
    Dim Picked              As Variant         'File-dialog result

    BookPath = ThisWorkbook.Path

    If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
        Candidate = BookPath & Application.PathSeparator & FileName
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate " & FileName & " automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate " & Title

    Picked = Application.GetOpenFilename( _
        FileFilter:=Title & " (*.csv),*.csv", _
        Title:="Select " & FileName)

    If VarType(Picked) = vbBoolean Then
        ResolveGridPath = vbNullString
    Else
        ResolveGridPath = CStr(Picked)
    End If
End Function
