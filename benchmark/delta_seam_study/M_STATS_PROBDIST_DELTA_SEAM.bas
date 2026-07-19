Attribute VB_Name = "M_STATS_PROBDIST_DELTA_SEAM"
'==============================================================================
' M_STATS_PROBDIST_DELTA_SEAM
'------------------------------------------------------------------------------
' PURPOSE
'   Self-contained export macro for the PROB_LogGammaDelta seam study. Fills the
'   observed_vba column of delta_seam_grid.csv, measuring three quantities per
'   (Large, Small) point so the delta kernel can be validated and the LogBeta
'   crossover chosen from measured VBA error envelopes:
'
'       LogGammaDelta   -> PROB_LogGammaDelta(Large, Small)
'       LogBeta_ident   -> PROB_LogGamma(Large)+PROB_LogGamma(Small)-PROB_LogGamma(Large+Small)
'       LogBeta_stable  -> PROB_LogGamma(Small)-PROB_LogGammaDelta(Large, Small)
'
' USAGE
'   Run Export_Delta_Seam and pick delta_seam_grid.csv in the dialog.
'
' DEPENDENCIES
'   - PROB_LogGamma, PROB_LogGammaDelta (M_STATS_PROBDIST_SPECIALFUNCS)
'
' UPDATED
'   2026-07-18
'==============================================================================
Option Explicit


Public Sub Export_Delta_Seam()
'
'==============================================================================
' Export_Delta_Seam
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the observed_vba column of the seam-study grid in place.
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

        'Columns: quantity, arg1(Large), arg2(Small), reference, observed_vba
        Cols = Split(Line, ",")
        If UBound(Cols) >= 4 Then
            A1 = Val(Cols(1))                  'Val is locale-independent
            A2 = Val(Cols(2))
            Cols(4) = EvaluateQuantity(Cols(0), A1, A2)
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

    MsgBox "Delta seam study complete: " & Filled & _
           " observation(s) written to " & vbCrLf & Path, _
           vbInformation, "Delta seam study"
End Sub


Private Function EvaluateQuantity( _
    ByVal Quantity As String, _
    ByVal LargeArg As Double, _
    ByVal SmallArg As Double) _
    As String
'
'==============================================================================
' EvaluateQuantity
'------------------------------------------------------------------------------
' PURPOSE
'   Dispatches the three measured quantities. Any runtime fault is reported as
'   "ERROR" so the analysis can flag the row.
'==============================================================================
'
    Dim Value               As Double

    On Error GoTo Err_Handler

    Select Case Quantity
        Case "LogGammaDelta"
            Value = PROB_LogGammaDelta(LargeArg, SmallArg)
        Case "LogBeta_ident"
            Value = PROB_LogGamma(LargeArg) + PROB_LogGamma(SmallArg) - _
                    PROB_LogGamma(LargeArg + SmallArg)
        Case "LogBeta_stable"
            Value = PROB_LogGamma(SmallArg) - PROB_LogGammaDelta(LargeArg, SmallArg)
        Case Else
            EvaluateQuantity = "ERROR"
            Exit Function
    End Select

    EvaluateQuantity = FormatFullPrecision(Value)
    Exit Function

Err_Handler:
    EvaluateQuantity = "ERROR"
End Function


Private Function ResolveGridPath() As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a usable LOCAL path to delta_seam_grid.csv, or "" if cancelled.
'   ThisWorkbook.Path is a URL on OneDrive/SharePoint, which Open cannot read,
'   so this prefers a local workbook folder and otherwise asks the user.
'==============================================================================
'
    Dim BookPath            As String
    Dim Candidate           As String
    Dim Picked              As Variant

    BookPath = ThisWorkbook.Path
    If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
        Candidate = BookPath & Application.PathSeparator & "delta_seam_grid.csv"
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate delta_seam_grid.csv automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate seam grid"
    Picked = Application.GetOpenFilename( _
        FileFilter:="Seam grid (*.csv),*.csv", _
        Title:="Select delta_seam_grid.csv")
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
