Attribute VB_Name = "M_STATS_PROBDIST_HOLDOUTOLDER"
'==============================================================================
' M_STATS_PROBDIST_HOLDOUTOLDER
'------------------------------------------------------------------------------
' PURPOSE
'   Fills observed_vba in holdout_older_grid.csv: a consolidated, fresh
'   (off-compliance-grid) holdout for the older families whose accuracy
'   contracts predate the independent-holdout discipline. Passing here validates
'   those thresholds on unseen data before their provenance is flipped to
'   "validated and frozen".
'
' USAGE
'   Run Export_HoldoutOlder and select holdout_older_grid.csv in the dialog.
'
' NOTES
'   - Arg slots (arg1, arg2, arg3) already match each function's signature order.
'   - Lognormal_ParametersFromMeanStdDev returns a 1x2 array; the element is
'     chosen by regime ("param_meanlog" / "param_stddevlog").
'
' UPDATED
'   2026-07-21
'==============================================================================
Option Explicit


Public Sub Export_HoldoutOlder()
    Dim Path                As String
    Dim FileNum             As Integer
    Dim OutNum              As Integer
    Dim Line                As String
    Dim Cols                As Variant
    Dim Lines()             As String
    Dim LineCount           As Long
    Dim IsHeader            As Boolean
    Dim Filled              As Long
    Dim I                   As Long

    Path = ResolveGridPath()
    If Len(Path) = 0 Then Exit Sub

    ReDim Lines(0 To 100000)
    LineCount = 0
    IsHeader = True
    Filled = 0

    FileNum = FreeFile
    Open Path For Input As #FileNum
    Do While Not EOF(FileNum)
        Line Input #FileNum, Line

        If IsHeader Then
            Lines(LineCount) = Line
            LineCount = LineCount + 1
            IsHeader = False
            GoTo ContinueLoop
        End If

        If Len(Trim$(Line)) = 0 Then GoTo ContinueLoop

        'Columns: function,vba_kernel,claim,metric,arg1,arg2,arg3,reference,observed_vba,regime,evidence_set
        Cols = Split(Line, ",")
        If UBound(Cols) >= 10 Then
            Cols(8) = EvaluateHoldout(Cols(0), Cols(9), Cols(4), Cols(5), Cols(6))
            Line = Join(Cols, ",")
            Filled = Filled + 1
        End If

        Lines(LineCount) = Line
        LineCount = LineCount + 1

ContinueLoop:
    Loop
    Close #FileNum

    OutNum = FreeFile
    Open Path For Output As #OutNum
    For I = 0 To LineCount - 1
        Print #OutNum, Lines(I)
    Next I
    Close #OutNum

    MsgBox "Older-families holdout export complete: " & Filled & " observation(s) written to " & vbCrLf & Path, _
           vbInformation, "Holdout older export"
End Sub


Private Function EvaluateHoldout( _
    ByVal FuncName As String, _
    ByVal Regime As String, _
    ByVal A1Text As String, _
    ByVal A2Text As String, _
    ByVal A3Text As String) _
    As String
'
    Dim V                   As Variant
    Dim Arr                 As Variant
    Dim A1                  As Double
    Dim A2                  As Double
    Dim A3                  As Double
    Dim Lb1                 As Long
    Dim Lb2                 As Long

    On Error GoTo Err_Handler

    A1 = Val(A1Text)
    A2 = Val(A2Text)
    A3 = Val(A3Text)

    Select Case FuncName
        Case "Normal_Cumulative": V = K_STATS_Normal_Cumulative(A1, A2, A3)
        Case "Normal_Survival": V = K_STATS_Normal_Survival(A1, A2, A3)
        Case "Normal_Density": V = K_STATS_Normal_Density(A1, A2, A3)
        Case "Normal_InverseCumulative": V = K_STATS_Normal_InverseCumulative(A1, A2, A3)
        Case "Normal_InverseSurvival": V = K_STATS_Normal_InverseSurvival(A1, A2, A3)
        Case "Normal_ZScore": V = K_STATS_Normal_ZScore(A1, A2, A3)
        Case "Lognormal_Cumulative": V = K_STATS_Lognormal_Cumulative(A1, A2, A3)
        Case "Lognormal_Survival": V = K_STATS_Lognormal_Survival(A1, A2, A3)
        Case "Lognormal_Density": V = K_STATS_Lognormal_Density(A1, A2, A3)
        Case "Lognormal_InverseCumulative": V = K_STATS_Lognormal_InverseCumulative(A1, A2, A3)
        Case "Lognormal_InverseSurvival": V = K_STATS_Lognormal_InverseSurvival(A1, A2, A3)
        Case "F_Cumulative": V = K_STATS_F_Cumulative(A1, A2, A3)
        Case "F_Survival": V = K_STATS_F_Survival(A1, A2, A3)
        Case "Gamma_Cumulative": V = K_STATS_Gamma_Cumulative(A1, A2, A3)
        Case "Gamma_Survival": V = K_STATS_Gamma_Survival(A1, A2, A3)
        Case "Gamma_Density": V = K_STATS_Gamma_Density(A1, A2, A3)
        Case "Gamma_InverseCumulative": V = K_STATS_Gamma_InverseCumulative(A1, A2, A3)
        Case "Beta_Cumulative": V = K_STATS_Beta_Cumulative(A1, A2, A3)
        Case "Beta_Survival": V = K_STATS_Beta_Survival(A1, A2, A3)
        Case "Beta_Density": V = K_STATS_Beta_Density(A1, A2, A3)
        Case "Beta_InverseCumulative": V = K_STATS_Beta_InverseCumulative(A1, A2, A3)
        Case "Weibull_Cumulative": V = K_STATS_Weibull_Cumulative(A1, A2, A3)
        Case "Weibull_Survival": V = K_STATS_Weibull_Survival(A1, A2, A3)
        Case "Weibull_Density": V = K_STATS_Weibull_Density(A1, A2, A3)
        Case "Weibull_InverseCumulative": V = K_STATS_Weibull_InverseCumulative(A1, A2, A3)
        Case "Uniform_Cumulative": V = K_STATS_Uniform_Cumulative(A1, A2, A3)
        Case "Uniform_Survival": V = K_STATS_Uniform_Survival(A1, A2, A3)
        Case "Uniform_Density": V = K_STATS_Uniform_Density(A1, A2, A3)
        Case "Uniform_InverseCumulative": V = K_STATS_Uniform_InverseCumulative(A1, A2, A3)
        Case "StudentT_Cumulative": V = K_STATS_StudentT_Cumulative(A1, A2)
        Case "StudentT_Survival": V = K_STATS_StudentT_Survival(A1, A2)
        Case "StudentT_Density": V = K_STATS_StudentT_Density(A1, A2)
        Case "StudentT_InverseCumulative": V = K_STATS_StudentT_InverseCumulative(A1, A2)
        Case "ChiSquare_Cumulative": V = K_STATS_ChiSquare_Cumulative(A1, A2)
        Case "ChiSquare_Survival": V = K_STATS_ChiSquare_Survival(A1, A2)
        Case "ChiSquare_InverseCumulative": V = K_STATS_ChiSquare_InverseCumulative(A1, A2)
        Case "Exponential_Cumulative": V = K_STATS_Exponential_Cumulative(A1, A2)
        Case "Exponential_Survival": V = K_STATS_Exponential_Survival(A1, A2)
        Case "Exponential_Density": V = K_STATS_Exponential_Density(A1, A2)
        Case "Exponential_InverseCumulative": V = K_STATS_Exponential_InverseCumulative(A1, A2)
        Case "Lognormal_Mean": V = K_STATS_Lognormal_Mean(A1, A2)
        Case "Lognormal_Variance": V = K_STATS_Lognormal_Variance(A1, A2)
        Case "Lognormal_StdDev": V = K_STATS_Lognormal_StdDev(A1, A2)
        Case "Gamma_Mean": V = K_STATS_Gamma_Mean(A1, A2)
        Case "Gamma_Variance": V = K_STATS_Gamma_Variance(A1, A2)
        Case "Gamma_StdDev": V = K_STATS_Gamma_StdDev(A1, A2)
        Case "Beta_Mean": V = K_STATS_Beta_Mean(A1, A2)
        Case "Beta_Variance": V = K_STATS_Beta_Variance(A1, A2)
        Case "Beta_StdDev": V = K_STATS_Beta_StdDev(A1, A2)
        Case "Weibull_Mean": V = K_STATS_Weibull_Mean(A1, A2)
        Case "Weibull_Variance": V = K_STATS_Weibull_Variance(A1, A2)
        Case "Weibull_StdDev": V = K_STATS_Weibull_StdDev(A1, A2)
        Case "NormalStandard_Cumulative": V = K_STATS_NormalStandard_Cumulative(A1)
        Case "NormalStandard_Survival": V = K_STATS_NormalStandard_Survival(A1)
        Case "NormalStandard_Density": V = K_STATS_NormalStandard_Density(A1)
        Case "NormalStandard_InverseCumulative": V = K_STATS_NormalStandard_InverseCumulative(A1)
        Case "NormalStandard_InverseCumulativeFast": V = K_STATS_NormalStandard_InverseCumulativeFast(A1)
        Case "NormalStandard_InverseSurvival": V = K_STATS_NormalStandard_InverseSurvival(A1)
        Case "NormalStandard_IntervalProbability": V = K_STATS_NormalStandard_IntervalProbability(A1, A2)

        Case "LogGamma": V = PROB_LogGamma(A1)
        Case "LogChoose": V = PROB_LogChoose(A1, A2)
        Case "LogGammaHalfDiff": V = PROB_LogGammaHalfDiff(A1)
        Case "StirlingError": V = PROB_StirlingError(A1)

        Case "Lognormal_ParametersFromMeanStdDev"
            Arr = K_STATS_Lognormal_ParametersFromMeanStdDev(A1, A2)
            If IsError(Arr) Then
                EvaluateHoldout = "ERROR"
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
        Candidate = BookPath & Application.PathSeparator & "holdout_older_grid.csv"
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate holdout_older_grid.csv automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate holdout grid"
    Picked = Application.GetOpenFilename( _
        FileFilter:="Holdout older grid (*.csv),*.csv", _
        Title:="Select holdout_older_grid.csv")
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
