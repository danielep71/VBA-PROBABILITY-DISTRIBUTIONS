Attribute VB_Name = "M_STATS_PROBDIST_DENSHELP"
'==============================================================================
' M_STATS_PROBDIST_DENSHELP
'------------------------------------------------------------------------------
' PURPOSE
'   Fills observed_vba for the four previously unbenchmarked public UDFs
'   (ChiSquare_Density, F_Density, Normal_IntervalProbability, and
'   Lognormal_ParametersFromMeanStdDev) in probability_accuracy_grid.csv. It
'   touches ONLY rows whose evidence_set is "density_helpers" and leaves every other
'   observation untouched, so it can run against the full main grid safely.
'
' USAGE
'   Run Export_DensityHelpers and select probability_accuracy_grid.csv in the dialog.
'
' NOTES
'   - Normal_IntervalProbability is exercised with Mean = 0 and StdDev = arg3.
'   - Lognormal_ParametersFromMeanStdDev returns a 1x2 array; the element is
'     chosen by regime ("param_meanlog" or "param_stddevlog").
'
' UPDATED
'   2026-07-21
'==============================================================================
Option Explicit


Public Sub Export_DensityHelpers()
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
            If Cols(10) = "density_helpers" Then
                Cols(8) = EvaluateF04(Cols(0), Cols(9), _
                                      Val(Cols(4)), Val(Cols(5)), Cols(6))
                Line = Join(Cols, ",")
                Filled = Filled + 1
            End If
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

    MsgBox "Density-helpers coverage export complete: " & Filled & " observation(s) written to " & vbCrLf & Path, _
           vbInformation, "Density helpers export"
End Sub


Private Function EvaluateF04( _
    ByVal FuncName As String, _
    ByVal Regime As String, _
    ByVal A1 As Double, _
    ByVal A2 As Double, _
    ByVal A3Text As String) _
    As String
'
    Dim V                   As Variant
    Dim Arr                 As Variant
    Dim Lb1                 As Long
    Dim Lb2                 As Long
    Dim A3                  As Double

    On Error GoTo Err_Handler

    Select Case FuncName
        Case "ChiSquare_Density"
            V = K_STATS_ChiSquare_Density(A1, A2)

        Case "F_Density"
            A3 = Val(A3Text)
            V = K_STATS_F_Density(A1, A2, A3)

        Case "Normal_IntervalProbability"
            A3 = Val(A3Text)                       'StdDev; Mean fixed at 0
            V = K_STATS_Normal_IntervalProbability(A1, A2, 0#, A3)

        Case "Lognormal_ParametersFromMeanStdDev"
            Arr = K_STATS_Lognormal_ParametersFromMeanStdDev(A1, A2)
            If IsError(Arr) Then
                EvaluateF04 = "ERROR"
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
            EvaluateF04 = "ERROR"
            Exit Function
    End Select

    If IsError(V) Then
        EvaluateF04 = "ERROR"
    ElseIf Not IsNumeric(V) Then
        EvaluateF04 = "ERROR"
    Else
        EvaluateF04 = FormatFullPrecision(CDbl(V))
    End If
    Exit Function

Err_Handler:
    EvaluateF04 = "ERROR"
End Function


Private Function ResolveGridPath() As String
    Dim BookPath            As String
    Dim Candidate           As String
    Dim Picked              As Variant

    BookPath = ThisWorkbook.Path
    If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
        Candidate = BookPath & Application.PathSeparator & "probability_accuracy_grid.csv"
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate probability_accuracy_grid.csv automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate main grid"
    Picked = Application.GetOpenFilename( _
        FileFilter:="Main accuracy grid (*.csv),*.csv", _
        Title:="Select probability_accuracy_grid.csv")
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
