Attribute VB_Name = "M_STATS_PROBDIST_BETAF_UNBAL"
'==============================================================================
' M_STATS_PROBDIST_BETAF_UNBAL
'------------------------------------------------------------------------------
' PURPOSE
'   Self-contained export macro for the step-6 unbalanced Beta/F study. Fills the
'   observed_vba column of beta_f_unbalanced_grid.csv by calling the public Beta
'   and F worksheet functions at strongly disparate shapes / degrees of freedom,
'   so their function-level relative error can be measured directly (rather than
'   inferred from the PROB_LogBeta proxy).
'
' USAGE
'   Run Export_BetaF_Unbalanced and pick beta_f_unbalanced_grid.csv in the dialog.
'
' NOTE
'   The K_STATS_ functions return Variant and may return CVErr on invalid input.
'   A CVErr or non-numeric result is written as "ERROR" so the analysis flags it.
'
' DEPENDENCIES
'   - K_STATS_Beta_Density, K_STATS_Beta_Cumulative, K_STATS_Beta_Survival
'   - K_STATS_F_Cumulative, K_STATS_F_Survival
'
' UPDATED
'   2026-07-18
'==============================================================================
Option Explicit


Public Sub Export_BetaF_Unbalanced()
'
'==============================================================================
' Export_BetaF_Unbalanced
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the observed_vba column of the unbalanced Beta/F grid in place.
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

        'Columns: function, vba_kernel, arg1, arg2, arg3, reference, observed_vba
        Cols = Split(Line, ",")
        If UBound(Cols) >= 6 Then
            A1 = Val(Cols(2))                  'Val is locale-independent
            A2 = Val(Cols(3))
            A3 = Val(Cols(4))
            Cols(6) = EvaluateFunction(Cols(0), A1, A2, A3)
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

    MsgBox "Unbalanced Beta/F study complete: " & Filled & _
           " observation(s) written to " & vbCrLf & Path, _
           vbInformation, "Beta/F unbalanced study"
End Sub


Private Function EvaluateFunction( _
    ByVal FuncName As String, _
    ByVal A1 As Double, _
    ByVal A2 As Double, _
    ByVal A3 As Double) _
    As String
'
'==============================================================================
' EvaluateFunction
'------------------------------------------------------------------------------
' PURPOSE
'   Dispatches the public Beta/F functions. A CVErr, non-numeric or runtime
'   fault is reported as "ERROR".
'==============================================================================
'
    Dim V                   As Variant

    On Error GoTo Err_Handler

    Select Case FuncName
        Case "Beta_Density":    V = K_STATS_Beta_Density(A1, A2, A3)
        Case "Beta_Cumulative": V = K_STATS_Beta_Cumulative(A1, A2, A3)
        Case "Beta_Survival":   V = K_STATS_Beta_Survival(A1, A2, A3)
        Case "F_Cumulative":    V = K_STATS_F_Cumulative(A1, A2, A3)
        Case "F_Survival":      V = K_STATS_F_Survival(A1, A2, A3)
        Case Else
            EvaluateFunction = "ERROR"
            Exit Function
    End Select

    If IsError(V) Then
        EvaluateFunction = "ERROR"
    ElseIf Not IsNumeric(V) Then
        EvaluateFunction = "ERROR"
    Else
        EvaluateFunction = FormatFullPrecision(CDbl(V))
    End If
    Exit Function

Err_Handler:
    EvaluateFunction = "ERROR"
End Function


Private Function ResolveGridPath() As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a usable LOCAL path to beta_f_unbalanced_grid.csv, or "" if
'   cancelled. ThisWorkbook.Path is a URL on OneDrive/SharePoint, which Open
'   cannot read, so this prefers a local workbook folder and otherwise asks.
'==============================================================================
'
    Dim BookPath            As String
    Dim Candidate           As String
    Dim Picked              As Variant

    BookPath = ThisWorkbook.Path
    If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
        Candidate = BookPath & Application.PathSeparator & "beta_f_unbalanced_grid.csv"
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate beta_f_unbalanced_grid.csv automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate Beta/F grid"
    Picked = Application.GetOpenFilename( _
        FileFilter:="Beta/F grid (*.csv),*.csv", _
        Title:="Select beta_f_unbalanced_grid.csv")
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
'   Renders a Double as "hi;lo" so hi + lo reproduces the original Double when
'   summed in Double precision on the Python side.
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
