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
'
' UPDATED
'   2026-07-19
'==============================================================================
Option Explicit


Public Sub Export_Holdout()
    Dim Path                As String
    Dim FileNum             As Integer
    Dim OutNum              As Integer
    Dim Line                As String
    Dim Cols                As Variant
    Dim Lines()             As String
    Dim LineCount           As Long
    Dim IsHeader            As Boolean
    Dim A1                  As Double
    Dim A2                  As Double
    Dim A3                  As Double
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
        If UBound(Cols) >= 8 Then
            A1 = Val(Cols(4))
            A2 = Val(Cols(5))
            A3 = Val(Cols(6))
            Cols(8) = EvaluateHoldout(Cols(0), A1, A2, A3)
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

    MsgBox "Holdout complete: " & Filled & " observation(s) written to " & vbCrLf & Path, _
           vbInformation, "Holdout"
End Sub


Private Function EvaluateHoldout( _
    ByVal FuncName As String, _
    ByVal A1 As Double, _
    ByVal A2 As Double, _
    ByVal A3 As Double) _
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
