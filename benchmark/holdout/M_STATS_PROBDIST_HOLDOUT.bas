Attribute VB_Name = "M_STATS_PROBDIST_HOLDOUT"
Option Explicit

'==============================================================================
' M_STATS_PROBDIST_HOLDOUT
'------------------------------------------------------------------------------
' PURPOSE
'   Export macro for the independent holdout. Fills observed_vba in
'   holdout_grid.csv using FRESH points that were not used to set any
'   threshold, so a contract that passes here generalises beyond its own
'   fitting set and can be frozen.
'
' WHY THIS EXISTS
'   A threshold measured on the same grid that produced it proves only that the
'   fit was recorded correctly. Confirming it on unseen points is what turns
'   'measured provisional' into 'validated and frozen'.
'
' WORKFLOW
'   1. python generate_holdout.py                  -> holdout_grid.csv
'   2. Place holdout_grid.csv next to this workbook (or pick it in the dialog)
'      and run Export_Holdout.
'   3. python analyze_holdout.py                   -> holdout_summary.md
'
' GRID FORMAT (header row, then one row per evaluation)
'   function, vba_kernel, claim, metric, arg1, arg2, arg3, arg4, reference,
'   observed_vba, regime, evidence_set
'
'   This macro reads function/arg1..arg4/regime and writes observed_vba only.
'   It never reads the reference column, so the observed side stays independent.
'   The regime is passed through because one contracted function publishes two
'   outputs and the regime selects which one the row claims.
'
' ERROR POLICY
'   A function that returns a worksheet error (CVErr) or raises writes the token
'   ERROR into observed_vba; analyze_holdout.py treats a non-numeric observed
'   value as an unusable point rather than a pass.
'
' DEPENDENCIES
'   - K_STATS_Beta_*        (Density/Cumulative/Survival/InverseCumulative)
'   - K_STATS_F_*           (Cumulative/Survival/InverseCumulative/Density)
'   - K_STATS_ChiSquare_Density
'   - K_STATS_Normal_*      (InverseSurvival/IntervalProbability)
'   - K_STATS_NormalStandard_InverseSurvival/InverseCumulative
'   - K_STATS_Lognormal_InverseSurvival/ParametersFromMeanStdDev
'   - K_STATS_Binomial_* / Poisson_* / Geometric_* / NegativeBinomial_* /
'     Hypergeometric_* / DiscreteUniform_*
'   - PROB_LogBeta
'
' UPDATED
'   2026-07-23
'==============================================================================


Public Sub Export_Holdout()
'
'==============================================================================
' Export_Holdout
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the observed_vba column of the holdout grid in place.
'
' BEHAVIOR
'   Reads the whole file and normalizes line endings before splitting, so
'   LF-only, CR-only and CRLF grids all parse. VBA Line Input is CR-delimited
'   and would swallow an entire LF-only file (.gitattributes stores *.csv as
'   eol=lf) as a single line, silently writing nothing.
'
' ERROR POLICY
'   Any failure closes the handle and reports once; the grid is left as found.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'Resolved grid path
    Dim Lines()             As String          'File lines
    Dim Raw                 As String          'File contents
    Dim FileNo              As Integer         'Input file handle
    Dim OutNo               As Integer         'Output file handle
    Dim I                   As Long            'Row index
    Dim Cols                As Variant         'Split fields of one row
    Dim Sep                 As String          'Field separator
    Dim A1                  As Double          'Parsed argument 1
    Dim A2                  As Double          'Parsed argument 2
    Dim A3                  As Double          'Parsed argument 3
    Dim A4                  As Double          'Parsed argument 4
    Dim Filled              As Long            'Rows written
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
        FileNo = FreeFile
        Open Path For Input As #FileNo
        Raw = Input$(LOF(FileNo), FileNo)
        Close #FileNo
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
            If UBound(Cols) < 11 Then GoTo ContinueRow

            A1 = ParseDouble(Cols(4))
            A2 = ParseDouble(Cols(5))
            A3 = ParseDouble(Cols(6))
            A4 = ParseDouble(Cols(7))

            'observed_vba is column index 9 in the 12-column arg4 schema;
            'the regime (index 10) disambiguates multi-output functions
            Cols(9) = EvaluateHoldout(Trim$(Cols(0)), A1, A2, A3, A4, _
                                      Trim$(CStr(Cols(10))))
            Lines(I) = Join(Cols, Sep)
            Filled = Filled + 1
ContinueRow:
        Next I
'------------------------------------------------------------------------------
' WRITE BACK
'------------------------------------------------------------------------------
        OutNo = FreeFile
        Open Path For Output As #OutNo
        For I = 0 To UBound(Lines)
            If I < UBound(Lines) Or Len(Lines(I)) > 0 Then Print #OutNo, Lines(I)
        Next I
        Close #OutNo

    MsgBox "Holdout complete: " & Filled & " observation(s) written to" & _
           vbCrLf & Path & vbCrLf & vbCrLf & _
           "Now run:  python analyze_holdout.py", vbInformation, "Holdout"
    Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    On Error Resume Next
    Close #FileNo
    Close #OutNo
    MsgBox "Holdout export failed: " & Err.Description, vbExclamation, "Holdout"
End Sub


Private Function EvaluateHoldout( _
    ByVal FuncName As String, _
    ByVal A1 As Double, _
    ByVal A2 As Double, _
    ByVal A3 As Double, _
    ByVal A4 As Double, _
    ByVal Regime As String) _
    As String
'
'==============================================================================
' EvaluateHoldout
'------------------------------------------------------------------------------
' PURPOSE
'   Dispatches one holdout row to its library function and returns a string
'   token: a full-precision number on success, or ERROR on any worksheet error.
'
' INPUTS
'   FuncName    contract function name from the grid
'   A1 - A4     parsed arguments; unused positions arrive as 0
'   Regime      contract regime; selects the output for multi-output functions
'
' RETURNS
'   Full-precision hi;lo token, or the literal ERROR.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant         'Raw function result
'------------------------------------------------------------------------------
' DISPATCH
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

    Select Case FuncName
        'Beta and F, the original holdout core
        Case "Beta_Density":                       V = K_STATS_Beta_Density(A1, A2, A3)
        Case "Beta_Cumulative":                    V = K_STATS_Beta_Cumulative(A1, A2, A3)
        Case "Beta_Survival":                      V = K_STATS_Beta_Survival(A1, A2, A3)
        Case "Beta_InverseCumulative":             V = K_STATS_Beta_InverseCumulative(A1, A2, A3)
        Case "F_Cumulative":                       V = K_STATS_F_Cumulative(A1, A2, A3)
        Case "F_Survival":                         V = K_STATS_F_Survival(A1, A2, A3)
        Case "F_InverseCumulative":                V = K_STATS_F_InverseCumulative(A1, A2, A3)
        Case "PROB_LogBeta":                       V = PROB_LogBeta(A1, A2)

        'Density helpers and parameter conversion
        Case "ChiSquare_Density":                  V = K_STATS_ChiSquare_Density(A1, A2)
        Case "F_Density":                          V = K_STATS_F_Density(A1, A2, A3)
        Case "Normal_IntervalProbability":         V = K_STATS_Normal_IntervalProbability(A1, A2, 0#, A3)
        Case "Lognormal_ParametersFromMeanStdDev": V = ExtractLognormalParam(A1, A2, Regime)

        'Deep-tail and split-boundary inverses
        Case "NormalStandard_InverseSurvival":     V = K_STATS_NormalStandard_InverseSurvival(A1)
        Case "NormalStandard_InverseCumulative":   V = K_STATS_NormalStandard_InverseCumulative(A1)
        Case "Normal_InverseSurvival":             V = K_STATS_Normal_InverseSurvival(A1, A2, A3)
        Case "Lognormal_InverseSurvival":          V = K_STATS_Lognormal_InverseSurvival(A1, A2, A3)

        'Binomial
        Case "Binomial_PMF":                       V = K_STATS_Binomial_PMF(A1, A2, A3)
        Case "Binomial_LogPMF":                    V = K_STATS_Binomial_LogPMF(A1, A2, A3)
        Case "Binomial_Cumulative":                V = K_STATS_Binomial_Cumulative(A1, A2, A3)
        Case "Binomial_Survival":                  V = K_STATS_Binomial_Survival(A1, A2, A3)
        Case "Binomial_InverseCumulative":         V = K_STATS_Binomial_InverseCumulative(A1, A2, A3)
        Case "Binomial_Mean":                      V = K_STATS_Binomial_Mean(A1, A2)
        Case "Binomial_Variance":                  V = K_STATS_Binomial_Variance(A1, A2)
        Case "Binomial_StdDev":                    V = K_STATS_Binomial_StdDev(A1, A2)

        'Poisson
        Case "Poisson_PMF":                        V = K_STATS_Poisson_PMF(A1, A2)
        Case "Poisson_LogPMF":                     V = K_STATS_Poisson_LogPMF(A1, A2)
        Case "Poisson_Cumulative":                 V = K_STATS_Poisson_Cumulative(A1, A2)
        Case "Poisson_Survival":                   V = K_STATS_Poisson_Survival(A1, A2)
        Case "Poisson_InverseCumulative":          V = K_STATS_Poisson_InverseCumulative(A1, A2)
        Case "Poisson_Mean":                       V = K_STATS_Poisson_Mean(A1)
        Case "Poisson_Variance":                   V = K_STATS_Poisson_Variance(A1)
        Case "Poisson_StdDev":                     V = K_STATS_Poisson_StdDev(A1)

        'Geometric
        Case "Geometric_PMF":                      V = K_STATS_Geometric_PMF(A1, A2)
        Case "Geometric_LogPMF":                   V = K_STATS_Geometric_LogPMF(A1, A2)
        Case "Geometric_Cumulative":               V = K_STATS_Geometric_Cumulative(A1, A2)
        Case "Geometric_Survival":                 V = K_STATS_Geometric_Survival(A1, A2)
        Case "Geometric_InverseCumulative":        V = K_STATS_Geometric_InverseCumulative(A1, A2)
        Case "Geometric_Mean":                     V = K_STATS_Geometric_Mean(A1)
        Case "Geometric_Variance":                 V = K_STATS_Geometric_Variance(A1)
        Case "Geometric_StdDev":                   V = K_STATS_Geometric_StdDev(A1)

        'NegativeBinomial
        Case "NegativeBinomial_PMF":               V = K_STATS_NegativeBinomial_PMF(A1, A2, A3)
        Case "NegativeBinomial_LogPMF":            V = K_STATS_NegativeBinomial_LogPMF(A1, A2, A3)
        Case "NegativeBinomial_Cumulative":        V = K_STATS_NegativeBinomial_Cumulative(A1, A2, A3)
        Case "NegativeBinomial_Survival":          V = K_STATS_NegativeBinomial_Survival(A1, A2, A3)
        Case "NegativeBinomial_InverseCumulative": V = K_STATS_NegativeBinomial_InverseCumulative(A1, A2, A3)
        Case "NegativeBinomial_Mean":              V = K_STATS_NegativeBinomial_Mean(A1, A2)
        Case "NegativeBinomial_Variance":          V = K_STATS_NegativeBinomial_Variance(A1, A2)
        Case "NegativeBinomial_StdDev":            V = K_STATS_NegativeBinomial_StdDev(A1, A2)

        'Hypergeometric
        Case "Hypergeometric_PMF":                 V = K_STATS_Hypergeometric_PMF(A1, A2, A3, A4)
        Case "Hypergeometric_LogPMF":              V = K_STATS_Hypergeometric_LogPMF(A1, A2, A3, A4)
        Case "Hypergeometric_Cumulative":          V = K_STATS_Hypergeometric_Cumulative(A1, A2, A3, A4)
        Case "Hypergeometric_Survival":            V = K_STATS_Hypergeometric_Survival(A1, A2, A3, A4)
        Case "Hypergeometric_InverseCumulative":   V = K_STATS_Hypergeometric_InverseCumulative(A1, A2, A3, A4)
        Case "Hypergeometric_Mean":                V = K_STATS_Hypergeometric_Mean(A1, A2, A3)
        Case "Hypergeometric_Variance":            V = K_STATS_Hypergeometric_Variance(A1, A2, A3)
        Case "Hypergeometric_StdDev":              V = K_STATS_Hypergeometric_StdDev(A1, A2, A3)

        'DiscreteUniform
        Case "DiscreteUniform_PMF":                V = K_STATS_DiscreteUniform_PMF(A1, A2, A3)
        Case "DiscreteUniform_LogPMF":             V = K_STATS_DiscreteUniform_LogPMF(A1, A2, A3)
        Case "DiscreteUniform_Cumulative":         V = K_STATS_DiscreteUniform_Cumulative(A1, A2, A3)
        Case "DiscreteUniform_Survival":           V = K_STATS_DiscreteUniform_Survival(A1, A2, A3)
        Case "DiscreteUniform_InverseCumulative":  V = K_STATS_DiscreteUniform_InverseCumulative(A1, A2, A3)
        Case "DiscreteUniform_Mean":               V = K_STATS_DiscreteUniform_Mean(A1, A2)
        Case "DiscreteUniform_Variance":           V = K_STATS_DiscreteUniform_Variance(A1, A2)
        Case "DiscreteUniform_StdDev":             V = K_STATS_DiscreteUniform_StdDev(A1, A2)

        Case Else
            EvaluateHoldout = "ERROR"
            Exit Function
    End Select
'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    If IsError(V) Then
        EvaluateHoldout = "ERROR"
    ElseIf Not IsNumeric(V) Then
        EvaluateHoldout = "ERROR"
    Else
        EvaluateHoldout = FormatFullPrecision(CDbl(V))
    End If
    Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    EvaluateHoldout = "ERROR"
End Function


Private Function ExtractLognormalParam( _
    ByVal Mean As Double, _
    ByVal StdDev As Double, _
    ByVal Regime As String) _
    As Variant
'
'==============================================================================
' ExtractLognormalParam
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the single lognormal parameter this grid row claims.
'
' WHY THIS EXISTS
'   K_STATS_Lognormal_ParametersFromMeanStdDev publishes two outputs from one
'   call. The contract splits them into the param_meanlog and param_stddevlog
'   regimes, so the regime is what selects the element to compare.
'
' RETURNS
'   The selected parameter, or the propagated CVErr so the caller records ERROR.
'==============================================================================
'
    Dim Arr                 As Variant         'Both parameters
    Dim Lb1                 As Long            'First dimension base
    Dim Lb2                 As Long            'Second dimension base

    Arr = K_STATS_Lognormal_ParametersFromMeanStdDev(Mean, StdDev)
    If IsError(Arr) Then
        ExtractLognormalParam = Arr
        Exit Function
    End If

    Lb1 = LBound(Arr, 1)
    Lb2 = LBound(Arr, 2)
    If Regime = "param_stddevlog" Then
        ExtractLognormalParam = Arr(Lb1, Lb2 + 1)
    Else
        ExtractLognormalParam = Arr(Lb1, Lb2)
    End If
End Function


Private Function ResolveGridPath() As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a usable LOCAL path to holdout_grid.csv, or an empty string if the
'   user cancels.
'
' WHY THIS EXISTS
'   ThisWorkbook.Path returns an http(s) URL when the workbook lives on OneDrive
'   or SharePoint, and Open cannot use it. Falling back to the file picker keeps
'   the macro usable there.
'==============================================================================
'
    Dim BookPath            As String          'Workbook folder
    Dim Candidate           As String          'Path next to the workbook
    Dim Picked              As Variant         'File picker result

    BookPath = ThisWorkbook.Path
    If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
        Candidate = BookPath & Application.PathSeparator & "holdout_grid.csv"
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate holdout_grid.csv automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate holdout grid"
    Picked = Application.GetOpenFilename( _
        FileFilter:="Holdout grid (*.csv),*.csv", _
        Title:="Select holdout_grid.csv")

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
    Dim S                   As String          'Cleaned token

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
'   Emits the observed value as a two-part hi;lo token.
'
' WHY THIS EXISTS
'   A single 15-digit field cannot round-trip a Double. Writing the residual as
'   a second field lets the Python side rebuild the exact value in Decimal, so
'   the measured error is the library's, not the file format's.
'==============================================================================
'
    Dim HiStr               As String          'Leading 15-digit field
    Dim Hi                  As Double          'Value of that field
    Dim Lo                  As Double          'Remaining residual

    If X = 0# Then
        FormatFullPrecision = "0E+000;0E+000"
        Exit Function
    End If

    HiStr = Fmt15(X)
    Hi = Val(HiStr)
    Lo = X - Hi
    FormatFullPrecision = HiStr & ";" & Fmt15(Lo)
End Function


Private Function Fmt15(ByVal X As Double) As String
'
'==============================================================================
' Fmt15
'------------------------------------------------------------------------------
' PURPOSE
'   Formats one Double in 15-significant-digit scientific notation with a US
'   decimal point, whatever the local settings are.
'==============================================================================
'
    Dim S                   As String          'Formatted token

    If X = 0# Then
        Fmt15 = "0E+000"
        Exit Function
    End If

    S = Format$(X, "0.00000000000000E+000")
    Fmt15 = Replace(S, ",", ".")               'Locale guard
End Function




