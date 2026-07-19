Attribute VB_Name = "M_STATS_PROBDIST_BETAF_INV"
'==============================================================================
' M_STATS_PROBDIST_BETAF_INV
'------------------------------------------------------------------------------
' PURPOSE
'   Self-contained export macro for the unbalanced Beta/F INVERSE study. Fills
'   the observed_vba column of beta_f_inverse_grid.csv with the quantile returned
'   by the public inverse functions. The analysis computes both the quantile
'   error and the forward-probability residual (quantile pushed back through the
'   true CDF), since inverse solvers amplify normalization error differently.
'
' USAGE
'   Run Export_BetaF_Inverse and pick beta_f_inverse_grid.csv in the dialog.
'
' NOTE
'   arg1 is the target probability p; arg2, arg3 are the shape / df parameters.
'   The K_STATS_ inverse functions return Variant and may return CVErr.
'
' DEPENDENCIES
'   - K_STATS_Beta_InverseCumulative, K_STATS_F_InverseCumulative
'
' UPDATED
'   2026-07-18
'==============================================================================
Option Explicit


Public Sub Export_BetaF_Inverse()
'
'==============================================================================
' Export_BetaF_Inverse
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the observed_vba column of the inverse grid in place with the returned
'   quantile x.
'==============================================================================
'
    Dim Path                As String
    Dim FileNum             As Integer
    Dim OutNum              As Integer
    Dim Line                As String
    Dim Cols                As Variant
    Dim Lines()             As String
    Dim LineCount           As Long
    Dim IsHeader            As Boolean
    Dim P                   As Double
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

        'Columns: function, vba_kernel, arg1(p), arg2, arg3, reference, observed_vba
        Cols = Split(Line, ",")
        If UBound(Cols) >= 6 Then
            P = Val(Cols(2))                   'Val is locale-independent
            A2 = Val(Cols(3))
            A3 = Val(Cols(4))
            Cols(6) = EvaluateInverse(Cols(0), P, A2, A3)
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

    MsgBox "Beta/F inverse study complete: " & Filled & _
           " observation(s) written to " & vbCrLf & Path, _
           vbInformation, "Beta/F inverse study"
End Sub


Private Function EvaluateInverse( _
    ByVal FuncName As String, _
    ByVal P As Double, _
    ByVal A2 As Double, _
    ByVal A3 As Double) _
    As String
'
'==============================================================================
' EvaluateInverse
'------------------------------------------------------------------------------
' PURPOSE
'   Dispatches the public inverse functions. A CVErr, non-numeric or runtime
'   fault is reported as "ERROR".
'==============================================================================
'
    Dim V                   As Variant

    On Error GoTo Err_Handler

    Select Case FuncName
        Case "Beta_InverseCumulative": V = K_STATS_Beta_InverseCumulative(P, A2, A3)
        Case "F_InverseCumulative":    V = K_STATS_F_InverseCumulative(P, A2, A3)
        Case Else
            EvaluateInverse = "ERROR"
            Exit Function
    End Select

    If IsError(V) Then
        EvaluateInverse = "ERROR"
    ElseIf Not IsNumeric(V) Then
        EvaluateInverse = "ERROR"
    Else
        EvaluateInverse = FormatFullPrecision(CDbl(V))
    End If
    Exit Function

Err_Handler:
    EvaluateInverse = "ERROR"
End Function


Private Function ResolveGridPath() As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a usable LOCAL path to beta_f_inverse_grid.csv, or "" if cancelled.
'==============================================================================
'
    Dim BookPath            As String
    Dim Candidate           As String
    Dim Picked              As Variant

    BookPath = ThisWorkbook.Path
    If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
        Candidate = BookPath & Application.PathSeparator & "beta_f_inverse_grid.csv"
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate beta_f_inverse_grid.csv automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate inverse grid"
    Picked = Application.GetOpenFilename( _
        FileFilter:="Inverse grid (*.csv),*.csv", _
        Title:="Select beta_f_inverse_grid.csv")
    If VarType(Picked) = vbBoolean Then
        ResolveGridPath = vbNullString
    Else
        ResolveGridPath = CStr(Picked)
    End If
End Function


Private Function FormatFullPrecision(ByVal X As Double) As String
'
'==============================================================================
' FormatFullPrecision
'------------------------------------------------------------------------------
' PURPOSE
'   Renders a Double as "hi;lo" so hi + lo reproduces the original Double.
'==============================================================================
'
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
'
'==============================================================================
' Fmt15
'------------------------------------------------------------------------------
' PURPOSE
'   Formats X to 15 significant digits, US decimal point regardless of locale.
'==============================================================================
'
    Dim S                   As String

    If X = 0# Then Fmt15 = "0E+000": Exit Function

    S = Format$(X, "0.00000000000000E+000")
    Fmt15 = Replace(S, ",", ".")
End Function
