Attribute VB_Name = "M_STATS_PROBDIST_FENVGAP"
'==============================================================================
' M_STATS_PROBDIST_FENVGAP
'------------------------------------------------------------------------------
' PURPOSE
'   Self-contained export macro for the F accuracy-envelope study. Fills
'   observed_vba in f_envelope_gap_grid.csv for F_Cumulative and F_Survival across a
'   fine sweep of the incomplete-beta shape parameter, so the exact df boundary
'   where F crosses its 1.1E-10 accuracy contract can be measured from real VBA.
'
' USAGE
'   Run Export_FEnvelopeGap and pick f_envelope_gap_grid.csv in the dialog.
'
' DEPENDENCIES
'   - K_STATS_F_Cumulative, K_STATS_F_Survival
'
' UPDATED
'   2026-07-19
'==============================================================================
Option Explicit


Public Sub Export_FEnvelopeGap()
    Dim Path                As String
    Dim FileNum             As Integer
    Dim OutNum              As Integer
    Dim Line                As String
    Dim Cols                As Variant
    Dim Lines()             As String
    Dim LineCount           As Long
    Dim IsHeader            As Boolean
    Dim XVal                As Double
    Dim D1                  As Double
    Dim D2                  As Double
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

        'Columns: function,vba_kernel,claim,metric,arg1(x),arg2(d1),arg3(d2),reference,observed_vba,regime,evidence_set
        Cols = Split(Line, ",")
        If UBound(Cols) >= 8 Then
            XVal = Val(Cols(4))
            D1 = Val(Cols(5))
            D2 = Val(Cols(6))
            Cols(8) = EvaluateF(Cols(0), XVal, D1, D2)
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

    MsgBox "F envelope GAP study complete: " & Filled & " observation(s) written to " & vbCrLf & Path, _
           vbInformation, "F envelope"
End Sub


Private Function EvaluateF( _
    ByVal FuncName As String, _
    ByVal XVal As Double, _
    ByVal D1 As Double, _
    ByVal D2 As Double) _
    As String
'
    Dim V                   As Variant

    On Error GoTo Err_Handler

    Select Case FuncName
        Case "F_Cumulative": V = K_STATS_F_Cumulative(XVal, D1, D2)
        Case "F_Survival":   V = K_STATS_F_Survival(XVal, D1, D2)
        Case Else
            EvaluateF = "ERROR"
            Exit Function
    End Select

    If IsError(V) Then
        EvaluateF = "ERROR"
    ElseIf Not IsNumeric(V) Then
        EvaluateF = "ERROR"
    Else
        EvaluateF = FormatFullPrecision(CDbl(V))
    End If
    Exit Function

Err_Handler:
    EvaluateF = "ERROR"
End Function


Private Function ResolveGridPath() As String
    Dim BookPath            As String
    Dim Candidate           As String
    Dim Picked              As Variant

    BookPath = ThisWorkbook.Path
    If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
        Candidate = BookPath & Application.PathSeparator & "f_envelope_gap_grid.csv"
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate f_envelope_gap_grid.csv automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate F envelope gap grid"
    Picked = Application.GetOpenFilename( _
        FileFilter:="F envelope grid (*.csv),*.csv", _
        Title:="Select f_envelope_gap_grid.csv")
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
