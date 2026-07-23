Attribute VB_Name = "M_STATS_PROBDIST_LOGBETA_STUDY"
Option Explicit

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
'   affecting the main regression export. The small hi;lo formatter and path
'   resolver are duplicated here on purpose, so this module depends only on
'   PROB_LogBeta (a Public kernel), nothing else in the harness.
'
' GRID FORMAT (study grid; 9 columns, no regime or evidence_set)
'   function, vba_kernel, claim, metric, arg1, arg2, arg3, reference,
'   observed_vba
'
'   arg1 is A and arg2 is B. This macro writes observed_vba (column index 8)
'   and never reads the reference column, so the observed side stays
'   independent. Rows whose function column is not "LogBeta" are passed through
'   untouched.
'
' OUTPUT
'   Each observation is written as a two-part "hi;lo" sum (two 15-digit numbers
'   VBA emits reliably); analyze_logbeta_switch.py sums them to recover the full
'   Double.
'
' USAGE
'   Run Export_LogBeta_Study and pick logbeta_switch_grid.csv in the dialog.
'
' ERROR POLICY
'   A runtime fault inside the kernel call is written as the token ERROR so the
'   analysis can flag that row. A failure elsewhere closes the handles and
'   reports once, leaving the grid as found.
'
' DEPENDENCIES
'   - PROB_LogBeta  (M_STATS_PROBDIST_SPECIALFUNCS)
'
' UPDATED
'   2026-07-23
'==============================================================================

Private Const LOGBETA_GRID_PATH As String = ""    'Empty => workbook folder or picker


Public Sub Export_LogBeta_Study()
'
'==============================================================================
' Export_LogBeta_Study
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the observed_vba column of the LogBeta switch-study grid in place.
'
' BEHAVIOR
'   Reads the whole file and normalizes line endings before splitting, so
'   LF-only, CR-only and CRLF grids all parse. VBA Line Input is CR-delimited
'   and would swallow an entire LF-only file (.gitattributes stores *.csv as
'   eol=lf) as a single line, silently writing nothing.
'
' ERROR POLICY
'   Any failure closes the handles and reports once; the grid is left as found.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'Resolved grid path
    Dim Lines()             As String          'All rewritten lines
    Dim Raw                 As String          'File contents
    Dim FileNum             As Integer         'Input channel
    Dim OutNum              As Integer         'Output channel
    Dim Cols                As Variant         'Split fields
    Dim Sep                 As String          'Field separator
    Dim A                   As Double          'First argument
    Dim B                   As Double          'Second argument
    Dim Observed            As String          'Formatted observation
    Dim FilledCount         As Long            'Rows populated
    Dim I                   As Long            'Row index
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler
    'Resolve the grid path (robust to OneDrive / SharePoint)
        Path = ResolveGridPath()
        If Len(Path) = 0 Then Exit Sub          'User cancelled the picker
        FilledCount = 0
        Sep = ","
    'Read the whole file
        FileNum = FreeFile
        Open Path For Input As #FileNum
        Raw = Input$(LOF(FileNum), FileNum)
        Close #FileNum
    'Normalize line endings and split
        Raw = Replace(Raw, vbCrLf, vbLf)
        Raw = Replace(Raw, vbCr, vbLf)
        Lines = Split(Raw, vbLf)
'------------------------------------------------------------------------------
' READ AND FILL
'------------------------------------------------------------------------------
    'Row 0 is the header; data starts at row 1
        For I = 1 To UBound(Lines)
            If Len(Trim$(Lines(I))) = 0 Then GoTo ContinueRow

            Cols = Split(Lines(I), Sep)
            If UBound(Cols) < 8 Then GoTo ContinueRow

            'Only LogBeta rows are owned here; anything else passes through
            If Trim$(Cols(0)) <> "LogBeta" Then GoTo ContinueRow

            A = ParseDouble(Cols(4))
            B = ParseDouble(Cols(5))
            Observed = EvaluateLogBeta(A, B)

            Cols(8) = Observed
            Lines(I) = Join(Cols, Sep)
            FilledCount = FilledCount + 1
ContinueRow:
        Next I
'------------------------------------------------------------------------------
' WRITE BACK
'------------------------------------------------------------------------------
    'Rewrite the file with the filled observations
        OutNum = FreeFile
        Open Path For Output As #OutNum
        For I = 0 To UBound(Lines)
            If I < UBound(Lines) Or Len(Lines(I)) > 0 Then Print #OutNum, Lines(I)
        Next I
        Close #OutNum

    MsgBox "LogBeta study complete: " & FilledCount & _
           " observation(s) written to" & vbCrLf & Path, _
           vbInformation, "LogBeta switch study"
    Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    On Error Resume Next
    Close #FileNum
    Close #OutNum
    MsgBox "LogBeta study failed: " & Err.Description, vbExclamation, _
           "LogBeta switch study"
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
'   Calls PROB_LogBeta(A, B) and returns the result as a hi;lo pair.
'
' INPUTS
'   A, B    the two Beta arguments from the grid
'
' RETURNS
'   Full-precision hi;lo token, or the literal ERROR on a runtime fault.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Value               As Double          'PROB_LogBeta result
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

        Value = PROB_LogBeta(A, B)
        EvaluateLogBeta = FormatFullPrecision(Value)
    Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
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
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Candidate           As String          'Path being tested
    Dim BookPath            As String          'Workbook folder
    Dim Picked              As Variant         'File-dialog result
'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
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
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim S                   As String          'Cleaned token
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
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
'   to 15 significant digits and hi + lo, summed in Double precision on the
'   Python side, reproduces the original Double exactly.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HiStr               As String          'Value to 15 significant digits
    Dim Hi                  As Double          'The Double that HiStr denotes
    Dim Lo                  As Double          'Exact residual X - Hi
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    If X = 0# Then
        FormatFullPrecision = "0E+000;0E+000"
        Exit Function
    End If

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
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim S                   As String          'Formatted value
'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    If X = 0# Then
        Fmt15 = "0E+000"
        Exit Function
    End If

    S = Format$(X, "0.00000000000000E+000")    '1 + 14 = 15 significant digits
    Fmt15 = Replace(S, ",", ".")               'Force US decimal regardless of locale
End Function


