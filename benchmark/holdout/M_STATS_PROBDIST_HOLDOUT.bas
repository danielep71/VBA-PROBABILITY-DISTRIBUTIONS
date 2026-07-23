Attribute VB_Name = "M_STATS_PROBDIST_HOLDOUT"

'==============================================================================
' M_STATS_PROBDIST_HOLDOUT
'------------------------------------------------------------------------------
' PURPOSE
'   Self-contained export macro for the independent holdout. Fills observed_vba
'   in holdout_grid.csv for the regime-specific contracts, using FRESH points not
'   used to set any threshold. Dispatches the public Beta/F functions (Variant)
'   and the PROB_LogBeta kernel (Double).
'
' USAGE
'   Run Export_Holdout and pick holdout_grid.csv in the dialog.
'
' DEPENDENCIES
'   - K_STATS_Beta_Density/Cumulative/Survival/InverseCumulative
'   - K_STATS_F_Cumulative/Survival/InverseCumulative
'   - PROB_LogBeta
'   - K_STATS_Binomial_* / Poisson_* / Geometric_* (PMF/LogPMF/Cumulative/
'     Survival/InverseCumulative/Mean/Variance/StdDev)
'
' UPDATED
'   2026-07-21
'==============================================================================
Option Explicit


Public Sub Export_Holdout()
    Dim Path                As String
    Dim FileNum             As Integer
    Dim OutNum              As Integer
    Dim Raw                 As String
    Dim Cols                As Variant
    Dim Lines()             As String
    Dim A1                  As Double
    Dim A2                  As Double
    Dim A3                  As Double
    Dim A4                  As Double
    Dim Filled              As Long
    Dim I                   As Long

    Path = ResolveGridPath()
    If Len(Path) = 0 Then Exit Sub

    Filled = 0

    'Read the whole file and normalize line endings, so LF-only, CR-only, and
    'CRLF grids all parse. VBA Line Input is CR-delimited and would swallow an
    'entire LF-only file (our .gitattributes stores *.csv as eol=lf) as one line.
    FileNum = FreeFile
    Open Path For Input As #FileNum
    Raw = Input$(LOF(FileNum), FileNum)
    Close #FileNum
    Raw = Replace(Raw, vbCrLf, vbLf)
    Raw = Replace(Raw, vbCr, vbLf)
    Lines = Split(Raw, vbLf)

    'Lines(0) is the header, kept verbatim; data starts at row 1.
    'Columns: function,vba_kernel,claim,metric,arg1,arg2,arg3,reference,observed_vba,regime,evidence_set
    For I = 1 To UBound(Lines)
        If Len(Trim$(Lines(I))) = 0 Then GoTo ContinueRow

        Cols = Split(Lines(I), ",")
        If UBound(Cols) >= 9 Then
            A1 = Val(Cols(4))
            A2 = Val(Cols(5))
            A3 = Val(Cols(6))
            A4 = Val(Cols(7))
            Cols(9) = EvaluateHoldout(Cols(0), A1, A2, A3, A4)
            Lines(I) = Join(Cols, ",")
            Filled = Filled + 1
        End If

ContinueRow:
    Next I

    OutNum = FreeFile
    Open Path For Output As #OutNum
    For I = 0 To UBound(Lines)
        If I < UBound(Lines) Or Len(Lines(I)) > 0 Then Print #OutNum, Lines(I)
    Next I
    Close #OutNum

    MsgBox "Holdout complete: " & Filled & " observation(s) written to " & vbCrLf & Path, _
           vbInformation, "Holdout"
End Sub


Private Function EvaluateHoldout( _
    ByVal FuncName As String, _
    ByVal A1 As Double, _
    ByVal A2 As Double, _
    ByVal A3 As Double, _
    ByVal A4 As Double) _
    As String
'
    Dim V                   As Variant

    On Error GoTo Err_Handler

    Select Case FuncName
        Case "Beta_Density":            V = K_STATS_Beta_Density(A1, A2, A3)
        Case "Beta_Cumulative":         V = K_STATS_Beta_Cumulative(A1, A2, A3)
        Case "Beta_Survival":           V = K_STATS_Beta_Survival(A1, A2, A3)
        Case "Beta_InverseCumulative":  V = K_STATS_Beta_InverseCumulative(A1, A2, A3)
        Case "F_Cumulative":            V = K_STATS_F_Cumulative(A1, A2, A3)
        Case "F_Survival":              V = K_STATS_F_Survival(A1, A2, A3)
        Case "F_InverseCumulative":     V = K_STATS_F_InverseCumulative(A1, A2, A3)
        Case "PROB_LogBeta":            V = PROB_LogBeta(A1, A2)
        Case "Binomial_PMF":            V = K_STATS_Binomial_PMF(A1, A2, A3)
        Case "Binomial_LogPMF":         V = K_STATS_Binomial_LogPMF(A1, A2, A3)
        Case "Binomial_Cumulative":     V = K_STATS_Binomial_Cumulative(A1, A2, A3)
        Case "Binomial_Survival":       V = K_STATS_Binomial_Survival(A1, A2, A3)
        Case "Binomial_InverseCumulative": V = K_STATS_Binomial_InverseCumulative(A1, A2, A3)
        Case "Binomial_Mean":           V = K_STATS_Binomial_Mean(A1, A2)
        Case "Binomial_Variance":       V = K_STATS_Binomial_Variance(A1, A2)
        Case "Binomial_StdDev":         V = K_STATS_Binomial_StdDev(A1, A2)
        Case "Poisson_PMF":             V = K_STATS_Poisson_PMF(A1, A2)
        Case "Poisson_LogPMF":          V = K_STATS_Poisson_LogPMF(A1, A2)
        Case "Poisson_Cumulative":      V = K_STATS_Poisson_Cumulative(A1, A2)
        Case "Poisson_Survival":        V = K_STATS_Poisson_Survival(A1, A2)
        Case "Poisson_InverseCumulative": V = K_STATS_Poisson_InverseCumulative(A1, A2)
        Case "Poisson_Mean":            V = K_STATS_Poisson_Mean(A1)
        Case "Poisson_Variance":        V = K_STATS_Poisson_Variance(A1)
        Case "Poisson_StdDev":          V = K_STATS_Poisson_StdDev(A1)
        Case "Geometric_PMF":           V = K_STATS_Geometric_PMF(A1, A2)
        Case "Geometric_LogPMF":        V = K_STATS_Geometric_LogPMF(A1, A2)
        Case "Geometric_Cumulative":    V = K_STATS_Geometric_Cumulative(A1, A2)
        Case "Geometric_Survival":      V = K_STATS_Geometric_Survival(A1, A2)
        Case "Geometric_InverseCumulative": V = K_STATS_Geometric_InverseCumulative(A1, A2)
        Case "Geometric_Mean":          V = K_STATS_Geometric_Mean(A1)
        Case "Geometric_Variance":      V = K_STATS_Geometric_Variance(A1)
        Case "Geometric_StdDev":        V = K_STATS_Geometric_StdDev(A1)
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
        Case "NormalStandard_InverseSurvival":  V = K_STATS_NormalStandard_InverseSurvival(A1)
        Case "NormalStandard_InverseCumulative": V = K_STATS_NormalStandard_InverseCumulative(A1)
        Case "Normal_InverseSurvival":       V = K_STATS_Normal_InverseSurvival(A1, A2, A3)
        Case "Lognormal_InverseSurvival":    V = K_STATS_Lognormal_InverseSurvival(A1, A2, A3)
        Case Else
            EvaluateHoldout = "ERROR"
            Exit Function
    End Select

    If IsError(V) Then
        EvaluateHoldout = "ERROR"
    ElseIf Not IsNumeric(V) Then
        EvaluateHoldout = "ERROR"
    Else
        EvaluateHoldout = FormatFullPrecision(CDbl(V))
    End If
    Exit Function

Err_Handler:
    EvaluateHoldout = "ERROR"
End Function


Private Function ResolveGridPath() As String
    Dim BookPath            As String
    Dim Candidate           As String
    Dim Picked              As Variant

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


Private Function FormatFullPrecision(ByVal X As Double) As String
    Dim HiStr               As String
    Dim Hi                  As Double
    Dim Lo                  As Double

    If X = 0# Then FormatFullPrecision = "0E+000;0E+000": Exit Function

    HiStr = Fmt15(X)
    Hi = Val(HiStr)
    Lo = X - Hi
    FormatFullPrecision = HiStr & ";" & Fmt15(Lo)
End Function


Private Function Fmt15(ByVal X As Double) As String
    Dim S                   As String

    If X = 0# Then Fmt15 = "0E+000": Exit Function

    S = Format$(X, "0.00000000000000E+000")
    Fmt15 = Replace(S, ",", ".")
End Function


