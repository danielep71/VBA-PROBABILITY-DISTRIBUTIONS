Attribute VB_Name = "M_STATS_PROBDIST_ACCURACYEXPORT"
Option Explicit

'==============================================================================
' M_STATS_PROBDIST_ACCURACY_EXPORT
'------------------------------------------------------------------------------
' PURPOSE
'   Phase 2 of the reproducible accuracy harness. Reads the grid produced by
'   generate_reference_values.py, evaluates each library function at every grid
'   point, and writes the observed values back so compute_errors.py can measure
'   the error against the mpmath reference.
'
' WHY THIS EXISTS
'   The reference values are generated in Python (mpmath, 50 digits). The
'   library under test is VBA and can only be executed inside Excel. This macro
'   is the bridge: it fills the observed_vba column that Python cannot.
'
' WORKFLOW
'   1. python generate_reference_values.py         -> probability_accuracy_grid.csv
'   2. Place probability_accuracy_grid.csv next to this workbook (or set the path
'      in ACCURACY_GRID_PATH) and run Export_Accuracy_Observations.
'   3. python compute_errors.py                    -> accuracy_summary.md
'
' GRID FORMAT (header row, then one row per evaluation)
'   function, vba_kernel, claim, metric, arg1, arg2, arg3, reference, observed_vba
'
'   This macro reads columns function/arg1..arg3, writes observed_vba. It does
'   not read or trust the reference column, so the observed side is independent.
'
' ERROR POLICY
'   A function that returns a worksheet error (CVErr) or raises writes the token
'   ERROR into observed_vba; compute_errors.py treats a non-numeric observed
'   value as a failed point. Empty inputs are passed as the function defaults.
'
' DEPENDENCIES
'   - M_STATS_PROBDIST_SPECIALFUNCS (PROB_* kernels)
'   - M_STATS_PROBDIST_TFAMILY      (K_STATS_* UDFs)
'
' UPDATED
'   2026-07-18
'==============================================================================

Private Const ACCURACY_GRID_PATH As String = ""        'Empty => same folder as the workbook


Public Sub Export_Accuracy_Observations()
'
'==============================================================================
' Export_Accuracy_Observations
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the observed_vba column of the accuracy grid in place.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'Resolved grid path
    Dim Lines()             As String          'File lines
    Dim Raw                 As String          'File contents
    Dim FileNo              As Integer         'File handle
    Dim I                   As Long            'Row index
    Dim Cols                As Variant         'Split fields of one row
    Dim FuncName            As String          'Function under test
    Dim A1 As Double, A2 As Double, A3 As Double, A4 As Double  'Parsed arguments
    Dim HasA1 As Boolean, HasA2 As Boolean, HasA3 As Boolean, HasA4 As Boolean
    Dim Observed            As String          'Observed value token
    Dim Sep                 As String          'Field separator
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler
    'Resolve the grid path (robust to OneDrive / SharePoint, where
    'ThisWorkbook.Path returns an http URL that Open cannot use)
        Path = ResolveGridPath()
        If Len(Path) = 0 Then Exit Sub          'User cancelled the picker
    'Read the whole file
        FileNo = FreeFile
        Open Path For Input As #FileNo
        Raw = Input$(LOF(FileNo), FileNo)
        Close #FileNo
    'Normalize line endings and split
        Raw = Replace(Raw, vbCrLf, vbLf)
        Raw = Replace(Raw, vbCr, vbLf)
        Lines = Split(Raw, vbLf)
        Sep = ","
'------------------------------------------------------------------------------
' EVALUATE EACH ROW
'------------------------------------------------------------------------------
    'Row 0 is the header; data starts at row 1
        For I = 1 To UBound(Lines)
            If Len(Trim$(Lines(I))) = 0 Then GoTo ContinueRow

            Cols = Split(Lines(I), Sep)
            If UBound(Cols) < 11 Then GoTo ContinueRow

            'This macro owns only the main grid. Study-sourced rows
            '(evidence_set <> "main grid") are populated by their own study
            'harnesses and must be left byte-for-byte untouched.
            If Trim$(Cols(11)) <> "main grid" Then GoTo ContinueRow

            FuncName = Trim$(Cols(0))
            HasA1 = (Len(Trim$(Cols(4))) > 0)
            HasA2 = (Len(Trim$(Cols(5))) > 0)
            HasA3 = (Len(Trim$(Cols(6))) > 0)
            HasA4 = (Len(Trim$(Cols(7))) > 0)
            If HasA1 Then A1 = ParseDouble(Cols(4))
            If HasA2 Then A2 = ParseDouble(Cols(5))
            If HasA3 Then A3 = ParseDouble(Cols(6))
            If HasA4 Then A4 = ParseDouble(Cols(7))

            Observed = EvaluateOne(FuncName, A1, A2, A3, A4, HasA2, HasA3, HasA4)

            'Rebuild the row with observed_vba filled (column index 9 in the
            '12-column arg4 schema)
            Cols(9) = Observed
            Lines(I) = Join(Cols, Sep)
ContinueRow:
        Next I
'------------------------------------------------------------------------------
' WRITE BACK
'------------------------------------------------------------------------------
        FileNo = FreeFile
        Open Path For Output As #FileNo
        For I = 0 To UBound(Lines)
            If I < UBound(Lines) Or Len(Lines(I)) > 0 Then Print #FileNo, Lines(I)
        Next I
        Close #FileNo

    MsgBox "Accuracy observations written to:" & vbCrLf & Path & vbCrLf & vbCrLf & _
           "Now run:  python compute_errors.py", vbInformation, "Accuracy export complete"
    Exit Sub
Err_Handler:
    On Error Resume Next
    Close #FileNo
    MsgBox "Accuracy export failed: " & Err.Description, vbExclamation
End Sub


Private Function ResolveGridPath() As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a usable LOCAL path to probability_accuracy_grid.csv, or an empty
'   string if the user cancels.
'
' WHY THIS EXISTS
'   ThisWorkbook.Path returns an http(s) URL when the workbook lives on OneDrive
'   or SharePoint, and VBA's Open statement cannot read a URL. This resolver
'   prefers an explicit local path, then the workbook folder only when that is a
'   real local path containing the file, and finally falls back to a file picker.
'==============================================================================
'
    Dim Candidate           As String          'Path being tested
    Dim BookPath            As String          'Workbook folder
    Dim Picked              As Variant          'File-dialog result

    '1. Explicit constant wins when it points at a real file
        If Len(ACCURACY_GRID_PATH) > 0 Then
            If Len(Dir$(ACCURACY_GRID_PATH)) > 0 Then
                ResolveGridPath = ACCURACY_GRID_PATH
                Exit Function
            End If
        End If

    '2. Workbook folder, but only if it is a LOCAL path (URLs start with http)
        BookPath = ThisWorkbook.Path
        If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
            Candidate = BookPath & Application.PathSeparator & "probability_accuracy_grid.csv"
            If Len(Dir$(Candidate)) > 0 Then
                ResolveGridPath = Candidate
                Exit Function
            End If
        End If

    '3. Ask the user to locate the file
        MsgBox "Could not locate probability_accuracy_grid.csv automatically " & _
               "(the workbook may be on OneDrive/SharePoint). Please select it.", _
               vbInformation, "Locate accuracy grid"
        Picked = Application.GetOpenFilename( _
            FileFilter:="Accuracy grid (*.csv),*.csv", _
            Title:="Select probability_accuracy_grid.csv")
        If VarType(Picked) = vbBoolean Then
            ResolveGridPath = vbNullString
        Else
            ResolveGridPath = CStr(Picked)
        End If
End Function


Private Function EvaluateOne( _
    ByVal FuncName As String, _
    ByVal A1 As Double, _
    ByVal A2 As Double, _
    ByVal A3 As Double, _
    ByVal A4 As Double, _
    ByVal HasA2 As Boolean, _
    ByVal HasA3 As Boolean, _
    ByVal HasA4 As Boolean) _
    As String
'
'==============================================================================
' EvaluateOne
'------------------------------------------------------------------------------
' PURPOSE
'   Dispatches one grid row to its library function and returns a string token:
'   a full-precision number on success, or ERROR on any worksheet error.
'==============================================================================
'
    Dim V                   As Variant         'Raw function result

    On Error GoTo Fail

    Select Case FuncName
        Case "LogGamma":                     V = PROB_LogGamma(A1)
        Case "LogGammaHalfDiff":             V = PROB_LogGammaHalfDiff(A1)
        Case "StirlingError":                V = PROB_StirlingError(A1)
        Case "LogChoose":                    V = PROB_LogChoose(A1, A2)

        Case "StudentT_Density":             V = K_STATS_StudentT_Density(A1, A2)
        Case "StudentT_Cumulative":          V = K_STATS_StudentT_Cumulative(A1, A2)
        Case "StudentT_Survival":            V = K_STATS_StudentT_Survival(A1, A2)
        Case "StudentT_InverseCumulative":   V = K_STATS_StudentT_InverseCumulative(A1, A2)

        Case "ChiSquare_Cumulative":         V = K_STATS_ChiSquare_Cumulative(A1, A2)
        Case "ChiSquare_Survival":           V = K_STATS_ChiSquare_Survival(A1, A2)
        Case "ChiSquare_InverseCumulative":  V = K_STATS_ChiSquare_InverseCumulative(A1, A2)

        Case "F_Cumulative":                 V = K_STATS_F_Cumulative(A1, A2, A3)
        Case "F_Survival":                   V = K_STATS_F_Survival(A1, A2, A3)
        Case "F_InverseCumulative":          V = K_STATS_F_InverseCumulative(A1, A2, A3)

        Case "NormalStandard_Density":       V = K_STATS_NormalStandard_Density(A1)
        Case "NormalStandard_Cumulative":    V = K_STATS_NormalStandard_Cumulative(A1)
        Case "NormalStandard_Survival":      V = K_STATS_NormalStandard_Survival(A1)
        Case "NormalStandard_InverseCumulative":     V = K_STATS_NormalStandard_InverseCumulative(A1)
        Case "NormalStandard_InverseSurvival":       V = K_STATS_NormalStandard_InverseSurvival(A1)
        Case "NormalStandard_InverseCumulativeFast": V = K_STATS_NormalStandard_InverseCumulativeFast(A1)
        Case "NormalStandard_IntervalProbability":   V = K_STATS_NormalStandard_IntervalProbability(A1, A2)

        Case "Normal_Density":               V = K_STATS_Normal_Density(A1, A2, A3)
        Case "Normal_Cumulative":            V = K_STATS_Normal_Cumulative(A1, A2, A3)
        Case "Normal_Survival":              V = K_STATS_Normal_Survival(A1, A2, A3)
        Case "Normal_InverseCumulative":     V = K_STATS_Normal_InverseCumulative(A1, A2, A3)
        Case "Normal_InverseSurvival":       V = K_STATS_Normal_InverseSurvival(A1, A2, A3)
        Case "Normal_ZScore":                V = K_STATS_Normal_ZScore(A1, A2, A3)

        Case "Lognormal_Density":            V = K_STATS_Lognormal_Density(A1, A2, A3)
        Case "Lognormal_Cumulative":         V = K_STATS_Lognormal_Cumulative(A1, A2, A3)
        Case "Lognormal_Survival":           V = K_STATS_Lognormal_Survival(A1, A2, A3)
        Case "Lognormal_InverseCumulative":  V = K_STATS_Lognormal_InverseCumulative(A1, A2, A3)
        Case "Lognormal_InverseSurvival":    V = K_STATS_Lognormal_InverseSurvival(A1, A2, A3)
        Case "Lognormal_Mean":               V = K_STATS_Lognormal_Mean(A1, A2)
        Case "Lognormal_Variance":           V = K_STATS_Lognormal_Variance(A1, A2)
        Case "Lognormal_StdDev":             V = K_STATS_Lognormal_StdDev(A1, A2)
        Case "Lognormal_ParamMeanLog":       V = ExtractParam(A1, A2, 1)
        Case "Lognormal_ParamStdDevLog":     V = ExtractParam(A1, A2, 2)

        Case "Gamma_Density":                V = K_STATS_Gamma_Density(A1, A2, A3)
        Case "Gamma_Cumulative":             V = K_STATS_Gamma_Cumulative(A1, A2, A3)
        Case "Gamma_Survival":               V = K_STATS_Gamma_Survival(A1, A2, A3)
        Case "Gamma_InverseCumulative":      V = K_STATS_Gamma_InverseCumulative(A1, A2, A3)
        Case "Gamma_Mean":                   V = K_STATS_Gamma_Mean(A1, A2)
        Case "Gamma_Variance":               V = K_STATS_Gamma_Variance(A1, A2)
        Case "Gamma_StdDev":                 V = K_STATS_Gamma_StdDev(A1, A2)

        Case "Beta_Density":                 V = K_STATS_Beta_Density(A1, A2, A3)
        Case "Beta_Cumulative":              V = K_STATS_Beta_Cumulative(A1, A2, A3)
        Case "Beta_Survival":                V = K_STATS_Beta_Survival(A1, A2, A3)
        Case "Beta_InverseCumulative":       V = K_STATS_Beta_InverseCumulative(A1, A2, A3)
        Case "Beta_Mean":                    V = K_STATS_Beta_Mean(A1, A2)
        Case "Beta_Variance":                V = K_STATS_Beta_Variance(A1, A2)
        Case "Beta_StdDev":                  V = K_STATS_Beta_StdDev(A1, A2)

        Case "Exponential_Density":          V = K_STATS_Exponential_Density(A1, A2)
        Case "Exponential_Cumulative":       V = K_STATS_Exponential_Cumulative(A1, A2)
        Case "Exponential_Survival":         V = K_STATS_Exponential_Survival(A1, A2)
        Case "Exponential_InverseCumulative": V = K_STATS_Exponential_InverseCumulative(A1, A2)

        Case "Weibull_Density":              V = K_STATS_Weibull_Density(A1, A2, A3)
        Case "Weibull_Cumulative":           V = K_STATS_Weibull_Cumulative(A1, A2, A3)
        Case "Weibull_Survival":             V = K_STATS_Weibull_Survival(A1, A2, A3)
        Case "Weibull_InverseCumulative":    V = K_STATS_Weibull_InverseCumulative(A1, A2, A3)
        Case "Weibull_Mean":                 V = K_STATS_Weibull_Mean(A1, A2)
        Case "Weibull_Variance":             V = K_STATS_Weibull_Variance(A1, A2)
        Case "Weibull_StdDev":               V = K_STATS_Weibull_StdDev(A1, A2)

        Case "Uniform_Density":              V = K_STATS_Uniform_Density(A1, A2, A3)
        Case "Uniform_Cumulative":           V = K_STATS_Uniform_Cumulative(A1, A2, A3)
        Case "Uniform_Survival":             V = K_STATS_Uniform_Survival(A1, A2, A3)
        Case "Uniform_InverseCumulative":    V = K_STATS_Uniform_InverseCumulative(A1, A2, A3)


        Case "Binomial_PMF":                V = K_STATS_Binomial_PMF(A1, A2, A3)
        Case "Binomial_LogPMF":             V = K_STATS_Binomial_LogPMF(A1, A2, A3)
        Case "Binomial_Cumulative":         V = K_STATS_Binomial_Cumulative(A1, A2, A3)
        Case "Binomial_Survival":           V = K_STATS_Binomial_Survival(A1, A2, A3)
        Case "Binomial_InverseCumulative":  V = K_STATS_Binomial_InverseCumulative(A1, A2, A3)
        Case "Binomial_Mean":               V = K_STATS_Binomial_Mean(A1, A2)
        Case "Binomial_Variance":           V = K_STATS_Binomial_Variance(A1, A2)
        Case "Binomial_StdDev":             V = K_STATS_Binomial_StdDev(A1, A2)

        Case "Poisson_PMF":                 V = K_STATS_Poisson_PMF(A1, A2)
        Case "Poisson_LogPMF":              V = K_STATS_Poisson_LogPMF(A1, A2)
        Case "Poisson_Cumulative":          V = K_STATS_Poisson_Cumulative(A1, A2)
        Case "Poisson_Survival":            V = K_STATS_Poisson_Survival(A1, A2)
        Case "Poisson_InverseCumulative":   V = K_STATS_Poisson_InverseCumulative(A1, A2)
        Case "Poisson_Mean":                V = K_STATS_Poisson_Mean(A1)
        Case "Poisson_Variance":            V = K_STATS_Poisson_Variance(A1)
        Case "Poisson_StdDev":              V = K_STATS_Poisson_StdDev(A1)

        Case "Geometric_PMF":               V = K_STATS_Geometric_PMF(A1, A2)
        Case "Geometric_LogPMF":            V = K_STATS_Geometric_LogPMF(A1, A2)
        Case "Geometric_Cumulative":        V = K_STATS_Geometric_Cumulative(A1, A2)
        Case "Geometric_Survival":          V = K_STATS_Geometric_Survival(A1, A2)
        Case "Geometric_InverseCumulative": V = K_STATS_Geometric_InverseCumulative(A1, A2)
        Case "Geometric_Mean":              V = K_STATS_Geometric_Mean(A1)
        Case "Geometric_Variance":          V = K_STATS_Geometric_Variance(A1)
        Case "Geometric_StdDev":            V = K_STATS_Geometric_StdDev(A1)


        Case "NegativeBinomial_PMF":         V = K_STATS_NegativeBinomial_PMF(A1, A2, A3)
        Case "NegativeBinomial_LogPMF":      V = K_STATS_NegativeBinomial_LogPMF(A1, A2, A3)
        Case "NegativeBinomial_Cumulative":  V = K_STATS_NegativeBinomial_Cumulative(A1, A2, A3)
        Case "NegativeBinomial_Survival":    V = K_STATS_NegativeBinomial_Survival(A1, A2, A3)
        Case "NegativeBinomial_InverseCumulative": V = K_STATS_NegativeBinomial_InverseCumulative(A1, A2, A3)
        Case "NegativeBinomial_Mean":        V = K_STATS_NegativeBinomial_Mean(A1, A2)
        Case "NegativeBinomial_Variance":    V = K_STATS_NegativeBinomial_Variance(A1, A2)
        Case "NegativeBinomial_StdDev":      V = K_STATS_NegativeBinomial_StdDev(A1, A2)

        Case "Hypergeometric_PMF":           V = K_STATS_Hypergeometric_PMF(A1, A2, A3, A4)
        Case "Hypergeometric_LogPMF":        V = K_STATS_Hypergeometric_LogPMF(A1, A2, A3, A4)
        Case "Hypergeometric_Cumulative":    V = K_STATS_Hypergeometric_Cumulative(A1, A2, A3, A4)
        Case "Hypergeometric_Survival":      V = K_STATS_Hypergeometric_Survival(A1, A2, A3, A4)
        Case "Hypergeometric_InverseCumulative": V = K_STATS_Hypergeometric_InverseCumulative(A1, A2, A3, A4)
        Case "Hypergeometric_Mean":          V = K_STATS_Hypergeometric_Mean(A1, A2, A3)
        Case "Hypergeometric_Variance":      V = K_STATS_Hypergeometric_Variance(A1, A2, A3)
        Case "Hypergeometric_StdDev":        V = K_STATS_Hypergeometric_StdDev(A1, A2, A3)


        Case "DiscreteUniform_PMF":          V = K_STATS_DiscreteUniform_PMF(A1, A2, A3)
        Case "DiscreteUniform_LogPMF":       V = K_STATS_DiscreteUniform_LogPMF(A1, A2, A3)
        Case "DiscreteUniform_Cumulative":   V = K_STATS_DiscreteUniform_Cumulative(A1, A2, A3)
        Case "DiscreteUniform_Survival":     V = K_STATS_DiscreteUniform_Survival(A1, A2, A3)
        Case "DiscreteUniform_InverseCumulative": V = K_STATS_DiscreteUniform_InverseCumulative(A1, A2, A3)
        Case "DiscreteUniform_Mean":         V = K_STATS_DiscreteUniform_Mean(A1, A2)
        Case "DiscreteUniform_Variance":     V = K_STATS_DiscreteUniform_Variance(A1, A2)
        Case "DiscreteUniform_StdDev":       V = K_STATS_DiscreteUniform_StdDev(A1, A2)

        Case Else:                           EvaluateOne = "ERROR": Exit Function
    End Select

    'A worksheet-error Variant is a failed point
        If IsError(V) Then EvaluateOne = "ERROR": Exit Function

    'Full-precision, locale-independent decimal
        EvaluateOne = FormatFullPrecision(CDbl(V))
    Exit Function
Fail:
    EvaluateOne = "ERROR"
End Function


Private Function ExtractParam( _
    ByVal Mean As Double, _
    ByVal StdDev As Double, _
    ByVal Which As Long) _
    As Variant
'
'==============================================================================
' ExtractParam
'------------------------------------------------------------------------------
' PURPOSE
'   Calls K_STATS_Lognormal_ParametersFromMeanStdDev (which returns a 1x2 array)
'   and returns element Which (1 = MeanLog, 2 = StdDevLog), so each output can be
'   measured as its own grid row.
'==============================================================================
'
    Dim R                   As Variant         'Parameter array

    R = K_STATS_Lognormal_ParametersFromMeanStdDev(Mean, StdDev)
    If IsError(R) Then ExtractParam = CVErr(xlErrNum): Exit Function
    ExtractParam = R(1, Which)
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
'   Renders a Double as a two-part sum "hi;lo", where hi and lo are each written
'   to 15 significant digits (the most VBA emits reliably) and hi + lo, summed
'   in Double precision on the Python side, reproduces the original Double
'   exactly. This removes the ~15-digit export floor without asking VBA to write
'   more digits than it can.
'
' WHY TWO PARTS
'   Format$, Str$ and CDec all cap a Double at about 15 significant digits, which
'   is coarser than several published accuracy claims. Writing hi (the value to
'   15 digits) and lo (the exact residual X - hi, also to 15 digits) lets the
'   analysis recover the full Double: lo carries the low-order bits hi dropped.
'   compute_errors.py sums the parts.
'==============================================================================
'
    Dim HiStr               As String          'Value to 15 significant digits
    Dim Hi                  As Double          'The Double that HiStr denotes
    Dim Lo                  As Double          'Exact residual X - Hi

    If X = 0# Then FormatFullPrecision = "0E+000;0E+000": Exit Function

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
'   Formats X to exactly 15 significant digits in scientific notation with a US
'   decimal point. This is within VBA's reliable precision, so the output is not
'   silently re-rounded.
'==============================================================================
'
    Dim S                   As String          'Formatted value

    If X = 0# Then Fmt15 = "0E+000": Exit Function

    S = Format$(X, "0.00000000000000E+000")    '1 + 14 = 15 significant digits
    Fmt15 = Replace(S, ",", ".")               'Force US decimal regardless of locale
End Function




