Attribute VB_Name = "M_STATS_PROBDIST_SURVBOUND"

Option Explicit

'==============================================================================
' M_STATS_PROBDIST_SURVBOUND
'------------------------------------------------------------------------------
' PURPOSE
'   Fills observed_vba in survival_boundary_grid.csv, the sweep that locates
'   where the standard-normal survival core stops meeting its tight RELATIVE
'   accuracy contract as z grows.
'
' WHY THIS EXISTS
'   The survival core shared by NormalStandard_Survival, Normal_Survival and
'   Lognormal_Survival holds ~1E-17 ABSOLUTE accuracy deep into the tail, but
'   its RELATIVE error grows with z. This study measures where that happens so
'   the central contracts' domains are restricted from evidence rather than
'   guessed. It is the evidence behind SurvivalTailRel in
'   benchmark/numerical_limitations.csv.
'
' USAGE
'   Run Export_SurvivalBoundary and select survival_boundary_grid.csv.
'
' GRID FORMAT (study grid; 11 columns, no arg4)
'   function, vba_kernel, claim, metric, arg1, arg2, arg3, reference,
'   observed_vba, regime, evidence_set
'
'   Arg slots already match each function's signature order. This macro writes
'   observed_vba (column index 8) and reads the regime (index 9) for functions
'   that publish more than one output. It never reads the reference column, so
'   the observed side stays independent. Note this is the study grid, not the
'   12-column main grid.
'
' SCOPE
'   The grid currently holds only the three survival functions. The dispatch
'   below is shared with the older-families holdout and is deliberately wider,
'   so a regenerated grid that adds points for other functions still exports
'   without a code change.
'
' ERROR POLICY
'   A CVErr, a non-numeric result or a runtime fault is written as the token
'   ERROR, which the analysis treats as an unusable point rather than a pass.
'
' DEPENDENCIES
'   - K_STATS_NormalStandard_Survival, K_STATS_Normal_Survival,
'     K_STATS_Lognormal_Survival (the functions this study measures)
'   - Further K_STATS_* and PROB_* routines reachable through the shared
'     dispatch
'
' UPDATED
'   2026-07-23
'==============================================================================


Public Sub Export_SurvivalBoundary()
'
'==============================================================================
' Export_SurvivalBoundary
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the observed_vba column of the survival-boundary grid in place.
'
' BEHAVIOR
'   Reads the whole file and normalizes line endings before splitting, so
'   LF-only, CR-only and CRLF grids all parse. VBA Line Input is CR-delimited
'   and would swallow an entire LF-only file (.gitattributes stores *.csv as
'   eol=lf) as a single line, silently writing nothing.
'
' ERROR POLICY
'   Any failure closes the handles and reports once; the grid is left as found.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'Resolved grid path
    Dim Lines()             As String          'File lines
    Dim Raw                 As String          'File contents
    Dim FileNum             As Integer         'Input file handle
    Dim OutNum              As Integer         'Output file handle
    Dim Cols                As Variant         'Split fields of one row
    Dim Sep                 As String          'Field separator
    Dim Filled              As Long            'Rows written
    Dim I                   As Long            'Row index
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler
    'Resolve the grid path (robust to OneDrive / SharePoint, where
    'ThisWorkbook.Path returns an http URL that Open cannot use)
        Path = ResolveGridPath()
        If Len(Path) = 0 Then Exit Sub          'User cancelled the picker
        Filled = 0
        Sep = ","
    'Read the whole file
        FileNum = FreeFile
        Open Path For Input As #FileNum
        Raw = Input$(LOF(FileNum), FileNum)
        Close #FileNum
    'Normalize line endings and split
        Raw = Replace(Raw, vbCrLf, vbLf)
        Raw = Replace(Raw, vbCr, vbLf)
        Lines = Split(Raw, vbLf)
'------------------------------------------------------------------------------
' EVALUATE EACH ROW
'------------------------------------------------------------------------------
    'Row 0 is the header; data starts at row 1
        For I = 1 To UBound(Lines)
            If Len(Trim$(Lines(I))) = 0 Then GoTo ContinueRow

            Cols = Split(Lines(I), Sep)
            If UBound(Cols) < 10 Then GoTo ContinueRow

            'observed_vba is index 8; the regime at index 9 disambiguates
            'functions that publish more than one output
            Cols(8) = EvaluateSurvivalBoundary(Trim$(Cols(0)), Trim$(CStr(Cols(9))), _
                                               CStr(Cols(4)), CStr(Cols(5)), CStr(Cols(6)))
            Lines(I) = Join(Cols, Sep)
            Filled = Filled + 1
ContinueRow:
        Next I
'------------------------------------------------------------------------------
' WRITE BACK
'------------------------------------------------------------------------------
        OutNum = FreeFile
        Open Path For Output As #OutNum
        For I = 0 To UBound(Lines)
            If I < UBound(Lines) Or Len(Lines(I)) > 0 Then Print #OutNum, Lines(I)
        Next I
        Close #OutNum

    MsgBox "Survival-boundary export complete: " & Filled & _
           " observation(s) written to" & vbCrLf & Path, _
           vbInformation, "Survival boundary export"
    Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    On Error Resume Next
    Close #FileNum
    Close #OutNum
    MsgBox "Survival-boundary export failed: " & Err.Description, _
           vbExclamation, "Survival boundary export"
End Sub


Private Function EvaluateSurvivalBoundary( _
    ByVal FuncName As String, _
    ByVal Regime As String, _
    ByVal A1Text As String, _
    ByVal A2Text As String, _
    ByVal A3Text As String) _
    As String
'
'==============================================================================
' EvaluateSurvivalBoundary
'------------------------------------------------------------------------------
' PURPOSE
'   Dispatches one grid row to its library function and returns a string token:
'   a full-precision number on success, or ERROR on any worksheet error.
'
' INPUTS
'   FuncName            contract function name from the grid
'   Regime              contract regime; selects the output for multi-output
'                       functions
'   A1Text - A3Text     arguments as text, since unused slots arrive empty
'
' RETURNS
'   Full-precision hi;lo token, or the literal ERROR.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant         'Raw function result
    Dim Arr                 As Variant         'Both lognormal parameters
    Dim A1                  As Double          'Parsed argument 1
    Dim A2                  As Double          'Parsed argument 2
    Dim A3                  As Double          'Parsed argument 3
    Dim Lb1                 As Long            'First dimension base
    Dim Lb2                 As Long            'Second dimension base
'------------------------------------------------------------------------------
' DISPATCH
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

        A1 = ParseDouble(A1Text)
        A2 = ParseDouble(A2Text)
        A3 = ParseDouble(A3Text)

    Select Case FuncName
        'Normal
        Case "Normal_Cumulative":                    V = K_STATS_Normal_Cumulative(A1, A2, A3)
        Case "Normal_Survival":                      V = K_STATS_Normal_Survival(A1, A2, A3)
        Case "Normal_Density":                       V = K_STATS_Normal_Density(A1, A2, A3)
        Case "Normal_InverseCumulative":             V = K_STATS_Normal_InverseCumulative(A1, A2, A3)
        Case "Normal_InverseSurvival":               V = K_STATS_Normal_InverseSurvival(A1, A2, A3)
        Case "Normal_ZScore":                        V = K_STATS_Normal_ZScore(A1, A2, A3)

        'Lognormal
        Case "Lognormal_Cumulative":                 V = K_STATS_Lognormal_Cumulative(A1, A2, A3)
        Case "Lognormal_Survival":                   V = K_STATS_Lognormal_Survival(A1, A2, A3)
        Case "Lognormal_Density":                    V = K_STATS_Lognormal_Density(A1, A2, A3)
        Case "Lognormal_InverseCumulative":          V = K_STATS_Lognormal_InverseCumulative(A1, A2, A3)
        Case "Lognormal_InverseSurvival":            V = K_STATS_Lognormal_InverseSurvival(A1, A2, A3)

        'F
        Case "F_Cumulative":                         V = K_STATS_F_Cumulative(A1, A2, A3)
        Case "F_Survival":                           V = K_STATS_F_Survival(A1, A2, A3)

        'Gamma
        Case "Gamma_Cumulative":                     V = K_STATS_Gamma_Cumulative(A1, A2, A3)
        Case "Gamma_Survival":                       V = K_STATS_Gamma_Survival(A1, A2, A3)
        Case "Gamma_Density":                        V = K_STATS_Gamma_Density(A1, A2, A3)
        Case "Gamma_InverseCumulative":              V = K_STATS_Gamma_InverseCumulative(A1, A2, A3)

        'Beta
        Case "Beta_Cumulative":                      V = K_STATS_Beta_Cumulative(A1, A2, A3)
        Case "Beta_Survival":                        V = K_STATS_Beta_Survival(A1, A2, A3)
        Case "Beta_Density":                         V = K_STATS_Beta_Density(A1, A2, A3)
        Case "Beta_InverseCumulative":               V = K_STATS_Beta_InverseCumulative(A1, A2, A3)

        'Weibull
        Case "Weibull_Cumulative":                   V = K_STATS_Weibull_Cumulative(A1, A2, A3)
        Case "Weibull_Survival":                     V = K_STATS_Weibull_Survival(A1, A2, A3)
        Case "Weibull_Density":                      V = K_STATS_Weibull_Density(A1, A2, A3)
        Case "Weibull_InverseCumulative":            V = K_STATS_Weibull_InverseCumulative(A1, A2, A3)

        'Uniform
        Case "Uniform_Cumulative":                   V = K_STATS_Uniform_Cumulative(A1, A2, A3)
        Case "Uniform_Survival":                     V = K_STATS_Uniform_Survival(A1, A2, A3)
        Case "Uniform_Density":                      V = K_STATS_Uniform_Density(A1, A2, A3)
        Case "Uniform_InverseCumulative":            V = K_STATS_Uniform_InverseCumulative(A1, A2, A3)

        'StudentT
        Case "StudentT_Cumulative":                  V = K_STATS_StudentT_Cumulative(A1, A2)
        Case "StudentT_Survival":                    V = K_STATS_StudentT_Survival(A1, A2)
        Case "StudentT_Density":                     V = K_STATS_StudentT_Density(A1, A2)
        Case "StudentT_InverseCumulative":           V = K_STATS_StudentT_InverseCumulative(A1, A2)

        'ChiSquare
        Case "ChiSquare_Cumulative":                 V = K_STATS_ChiSquare_Cumulative(A1, A2)
        Case "ChiSquare_Survival":                   V = K_STATS_ChiSquare_Survival(A1, A2)
        Case "ChiSquare_InverseCumulative":          V = K_STATS_ChiSquare_InverseCumulative(A1, A2)

        'Exponential
        Case "Exponential_Cumulative":               V = K_STATS_Exponential_Cumulative(A1, A2)
        Case "Exponential_Survival":                 V = K_STATS_Exponential_Survival(A1, A2)
        Case "Exponential_Density":                  V = K_STATS_Exponential_Density(A1, A2)
        Case "Exponential_InverseCumulative":        V = K_STATS_Exponential_InverseCumulative(A1, A2)

        'Lognormal
        Case "Lognormal_Mean":                       V = K_STATS_Lognormal_Mean(A1, A2)
        Case "Lognormal_Variance":                   V = K_STATS_Lognormal_Variance(A1, A2)
        Case "Lognormal_StdDev":                     V = K_STATS_Lognormal_StdDev(A1, A2)

        'Gamma
        Case "Gamma_Mean":                           V = K_STATS_Gamma_Mean(A1, A2)
        Case "Gamma_Variance":                       V = K_STATS_Gamma_Variance(A1, A2)
        Case "Gamma_StdDev":                         V = K_STATS_Gamma_StdDev(A1, A2)

        'Beta
        Case "Beta_Mean":                            V = K_STATS_Beta_Mean(A1, A2)
        Case "Beta_Variance":                        V = K_STATS_Beta_Variance(A1, A2)
        Case "Beta_StdDev":                          V = K_STATS_Beta_StdDev(A1, A2)

        'Weibull
        Case "Weibull_Mean":                         V = K_STATS_Weibull_Mean(A1, A2)
        Case "Weibull_Variance":                     V = K_STATS_Weibull_Variance(A1, A2)
        Case "Weibull_StdDev":                       V = K_STATS_Weibull_StdDev(A1, A2)

        'NormalStandard
        Case "NormalStandard_Cumulative":            V = K_STATS_NormalStandard_Cumulative(A1)
        Case "NormalStandard_Survival":              V = K_STATS_NormalStandard_Survival(A1)
        Case "NormalStandard_Density":               V = K_STATS_NormalStandard_Density(A1)
        Case "NormalStandard_InverseCumulative":     V = K_STATS_NormalStandard_InverseCumulative(A1)
        Case "NormalStandard_InverseCumulativeFast": V = K_STATS_NormalStandard_InverseCumulativeFast(A1)
        Case "NormalStandard_InverseSurvival":       V = K_STATS_NormalStandard_InverseSurvival(A1)
        Case "NormalStandard_IntervalProbability":   V = K_STATS_NormalStandard_IntervalProbability(A1, A2)

        'Special-function kernels
        Case "LogGamma":                             V = PROB_LogGamma(A1)
        Case "LogChoose":                            V = PROB_LogChoose(A1, A2)
        Case "LogGammaHalfDiff":                     V = PROB_LogGammaHalfDiff(A1)
        Case "StirlingError":                        V = PROB_StirlingError(A1)

        'Multi-output conversion; kept for a regenerated grid that includes it
        Case "Lognormal_ParametersFromMeanStdDev"
            Arr = K_STATS_Lognormal_ParametersFromMeanStdDev(A1, A2)
            If IsError(Arr) Then
                EvaluateSurvivalBoundary = "ERROR"
                Exit Function
            End If
            Lb1 = LBound(Arr, 1)
            Lb2 = LBound(Arr, 2)
            If Regime = "param_stddevlog" Then
                V = Arr(Lb1, Lb2 + 1)
            Else
                V = Arr(Lb1, Lb2)
            End If

        Case Else
            EvaluateSurvivalBoundary = "ERROR"
            Exit Function
    End Select
'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    If IsError(V) Then
        EvaluateSurvivalBoundary = "ERROR"
    ElseIf Not IsNumeric(V) Then
        EvaluateSurvivalBoundary = "ERROR"
    Else
        EvaluateSurvivalBoundary = FormatFullPrecision(CDbl(V))
    End If
    Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    EvaluateSurvivalBoundary = "ERROR"
End Function


Private Function ResolveGridPath() As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a usable LOCAL path to survival_boundary_grid.csv, or an empty
'   string if the user cancels.
'
' WHY THIS EXISTS
'   ThisWorkbook.Path returns an http(s) URL when the workbook lives on OneDrive
'   or SharePoint, and Open cannot read a URL. This prefers a local workbook
'   folder and otherwise asks the user.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim BookPath            As String          'Workbook folder
    Dim Candidate           As String          'Path next to the workbook
    Dim Picked              As Variant         'File-dialog result
'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    BookPath = ThisWorkbook.Path
    If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
        Candidate = BookPath & Application.PathSeparator & "survival_boundary_grid.csv"
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate survival_boundary_grid.csv automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate survival boundary grid"
    Picked = Application.GetOpenFilename( _
        FileFilter:="Survival boundary grid (*.csv),*.csv", _
        Title:="Select survival_boundary_grid.csv")

    If VarType(Picked) = vbBoolean Then
        ResolveGridPath = vbNullString
    Else
        ResolveGridPath = CStr(Picked)
    End If
End Function


Private Function ParseDouble(ByVal Text As String) As Double
'
'==============================================================================
' ParseDouble
'------------------------------------------------------------------------------
' PURPOSE
'   Parses a grid number written with a US decimal point, independent of the
'   local list/decimal separators.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim S                   As String          'Cleaned token
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    S = Trim$(Text)
    S = Replace(S, ",", ".")                   'Guard against a stray locale comma
    ParseDouble = Val(S)                       'Val always reads "." as decimal
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
' WHY TWO PARTS
'   Format$, Str$ and CDec all cap a Double at about 15 significant digits,
'   which is coarser than the accuracy this study measures. Writing the residual
'   X - hi as a second field carries the low-order bits hi dropped.
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
    Hi = Val(HiStr)                            'Val is locale-independent; CDbl is not
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
    Fmt15 = Replace(S, ",", ".")               'Force US decimal regardless of locale
End Function


