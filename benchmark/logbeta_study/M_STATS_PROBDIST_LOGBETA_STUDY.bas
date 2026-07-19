Attribute VB_Name = "M_STATS_PROBDIST_LOGBETA_STUDY"
'==============================================================================
' M_STATS_PROBDIST_LOGBETA_STUDY
'------------------------------------------------------------------------------
' PURPOSE
'   Self-contained export macro for the unbalanced-Beta switch study. Fills the
'   observed_vba column of logbeta_switch_grid.csv by calling PROB_LogBeta at
'   each (A, B) pair, so the asymptotic-versus-general switch in PROB_LogBeta can
'   be characterized against the mpmath reference.
'
' WHY THIS EXISTS AS A SEPARATE MODULE
'   It deliberately does not touch the main M_STATS_PROBDIST_ACCURACYEXPORT
'   macro. The study is isolated so it can be added and removed without
'   affecting the 66-function regression export. The small hi;lo formatter and
'   path resolver are duplicated here on purpose, so this module depends only on
'   PROB_LogBeta (a Public kernel), nothing else in the harness.
'
' OUTPUT
'   Each observation is written as a two-part "hi;lo" sum (two 15-digit numbers
'   VBA emits reliably); analyze_logbeta_switch.py sums them to recover the full
'   Double. Rows whose function column is not "LogBeta" are left untouched.
'
' USAGE
'   Run Export_LogBeta_Study and pick logbeta_switch_grid.csv in the dialog.
'
' DEPENDENCIES
'   - PROB_LogBeta  (M_STATS_PROBDIST_SPECIALFUNCS)
'
' UPDATED
'   2026-07-18
'==============================================================================
Option Explicit

Private Const LOGBETA_GRID_PATH As String = ""    'Empty => resolve via workbook folder or picker


Public Sub Export_LogBeta_Study()
'
'==============================================================================
' Export_LogBeta_Study
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the observed_vba column of the LogBeta switch-study grid in place.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'Resolved grid path
    Dim FileNum             As Integer         'Input channel
    Dim OutNum              As Integer         'Output channel
    Dim Line                As String          'Current CSV line
    Dim Cols                As Variant         'Split fields
    Dim Lines()             As String          'All rewritten lines
    Dim LineCount           As Long            'Number of buffered lines
    Dim IsHeader            As Boolean         'First-row flag
    Dim A                   As Double          'First argument
    Dim B                   As Double          'Second argument
    Dim Observed            As String          'Formatted observation
    Dim FilledCount         As Long            'Rows populated

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Resolve the grid path (robust to OneDrive / SharePoint)
        Path = ResolveGridPath()
        If Len(Path) = 0 Then Exit Sub          'User cancelled the picker

    'Prepare the line buffer
        ReDim Lines(0 To 100000)
        LineCount = 0
        IsHeader = True
        FilledCount = 0

'------------------------------------------------------------------------------
' READ AND FILL
'------------------------------------------------------------------------------
    'Read every line, filling observed_vba on LogBeta rows
        FileNum = FreeFile
        Open Path For Input As #FileNum
        Do While Not EOF(FileNum)
            Line Input #FileNum, Line

            'Pass the header through unchanged
                If IsHeader Then
                    Lines(LineCount) = Line
                    LineCount = LineCount + 1
                    IsHeader = False
                    GoTo ContinueLoop
                End If

            'Skip empty trailing lines
                If Len(Trim$(Line)) = 0 Then GoTo ContinueLoop

            'Split the row: function, vba_kernel, claim, metric, arg1, arg2,
            'arg3, reference, observed_vba
                Cols = Split(Line, ",")

            'Only process LogBeta rows; pass anything else through unchanged
                If UBound(Cols) >= 8 Then
                    If Cols(0) = "LogBeta" Then
                        A = Val(Cols(4))         'Val is locale-independent
                        B = Val(Cols(5))
                        Observed = EvaluateLogBeta(A, B)
                        Cols(8) = Observed
                        Line = Join(Cols, ",")
                        FilledCount = FilledCount + 1
                    End If
                End If

            Lines(LineCount) = Line
            LineCount = LineCount + 1

ContinueLoop:
        Loop
        Close #FileNum

'------------------------------------------------------------------------------
' WRITE BACK
'------------------------------------------------------------------------------
    'Rewrite the file with the filled observations
        OutNum = FreeFile
        Open Path For Output As #OutNum
        Dim I As Long
        For I = 0 To LineCount - 1
            Print #OutNum, Lines(I)
        Next I
        Close #OutNum

    'Report
        MsgBox "LogBeta study complete: " & FilledCount & _
               " observation(s) written to " & vbCrLf & Path, _
               vbInformation, "LogBeta switch study"
End Sub


Private Function EvaluateLogBeta( _
    ByVal A As Double, _
    ByVal B As Double) _
    As String
'
'==============================================================================
' EvaluateLogBeta
'------------------------------------------------------------------------------
' PURPOSE
'   Calls PROB_LogBeta(A, B) and returns the result as a hi;lo pair. Any runtime
'   fault is reported as "ERROR" so the analysis can flag the row.
'==============================================================================
'
    Dim Value               As Double          'PROB_LogBeta result

    On Error GoTo Err_Handler

    Value = PROB_LogBeta(A, B)
    EvaluateLogBeta = FormatFullPrecision(Value)
    Exit Function

Err_Handler:
    EvaluateLogBeta = "ERROR"
End Function


Private Function ResolveGridPath() As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a usable LOCAL path to logbeta_switch_grid.csv, or an empty string
'   if the user cancels.
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
        If Len(LOGBETA_GRID_PATH) > 0 Then
            If Len(Dir$(LOGBETA_GRID_PATH)) > 0 Then
                ResolveGridPath = LOGBETA_GRID_PATH
                Exit Function
            End If
        End If

    '2. Workbook folder, but only if it is a LOCAL path (URLs start with http)
        BookPath = ThisWorkbook.Path
        If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
            Candidate = BookPath & Application.PathSeparator & "logbeta_switch_grid.csv"
            If Len(Dir$(Candidate)) > 0 Then
                ResolveGridPath = Candidate
                Exit Function
            End If
        End If

    '3. Ask the user to locate the file
        MsgBox "Could not locate logbeta_switch_grid.csv automatically " & _
               "(the workbook may be on OneDrive/SharePoint). Please select it.", _
               vbInformation, "Locate LogBeta grid"
        Picked = Application.GetOpenFilename( _
            FileFilter:="LogBeta grid (*.csv),*.csv", _
            Title:="Select logbeta_switch_grid.csv")
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
'   Renders a Double as a two-part sum "hi;lo", where hi and lo are each written
'   to 15 significant digits and hi + lo, summed in Double precision on the
'   Python side, reproduces the original Double exactly.
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
'   decimal point.
'==============================================================================
'
    Dim S                   As String          'Formatted value

    If X = 0# Then Fmt15 = "0E+000": Exit Function

    S = Format$(X, "0.00000000000000E+000")    '1 + 14 = 15 significant digits
    Fmt15 = Replace(S, ",", ".")               'Force US decimal regardless of locale
End Function
