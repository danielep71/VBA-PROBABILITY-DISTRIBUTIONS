Attribute VB_Name = "M_STATS_PROBDIST_GBF"
Option Explicit

'==============================================================================
' M_STATS_PROBDIST_GBF
'------------------------------------------------------------------------------
' PURPOSE
'   Export macro for the #307 binary64 route fit set (FIT_DESIGN.md). Fills
'   the observed side of b64_fit_grid.csv: at each (A, X) it calls the
'   repaired binary64 lower series and continued fraction, records each
'   routine's status and value, and times the routine that answered.
'
' HISTORICAL
'   The routines it calls were generated into production only at the fit
'   candidate 88ff6ba (#312) and left production after #307 closed. The study
'   grid gbf_b64_fit is no longer registered for export; this module compiles
'   only against that candidate's source.
'
' USAGE
'   Run Export_GammaB64Fit. It finds the grid beside the workbook, or asks
'   for it. ci/Export-ExcelGrids.ps1 runs it on the self-hosted runner with
'   the grid staged beside the export workbook (#225).
'
' GRID FORMAT (study grid; 14 columns)
'   point_id, shape, x, ref_p, ref_q, echo_shape, echo_x, series_status,
'   series_p, cf_status, cf_q, route_us, route_calls, note
'
'   This macro writes the last nine columns. It never reads the references,
'   so the observed side stays independent. Diagnostics are written with
'   their commas replaced, so the row keeps its 14 fields.
'
' TIMING
'   route_us is the mean microseconds per call of the routine that returned
'   OK (the series if both did), repeated, doubling the count from 16, until
'   at least 0.2 s elapsed; route_calls is that count. A row where neither
'   routine answered has no timing.
'
' ERROR POLICY
'   A routine that returns False is recorded as DECLINED with its diagnostic;
'   a runtime fault as ERROR. Neither is ever scored as a value.
'
' DEPENDENCIES
'   - PROB_TryGammaB64SeriesP              (M_STATS_PROBDIST_SPECIALFUNCS)
'   - PROB_TryGammaB64ContinuedFractionQ   (M_STATS_PROBDIST_SPECIALFUNCS)
'
' UPDATED
'   2026-11-10
'==============================================================================


Public Sub Export_GammaB64Fit()
'
'==============================================================================
' Export_GammaB64Fit
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the observed columns of the #307 binary64 route fit set in place.
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
    Dim Lines()             As String          'File lines
    Dim Raw                 As String          'File contents
    Dim FileNum             As Integer         'Input file handle
    Dim OutNum              As Integer         'Output file handle
    Dim Cols                As Variant         'Split fields of one row
    Dim Sep                 As String          'Field separator
    Dim A                   As Double          'Shape (column 1)
    Dim X                   As Double          'Argument (column 2)
    Dim V                   As Double          'Routine result
    Dim Msg                 As String          'Routine diagnostic
    Dim Note                As String          'Both diagnostics
    Dim Answered            As String          'Routine timed, or empty
    Dim Calls               As Long            'Timed call count
    Dim Filled              As Long            'Rows written
    Dim I                   As Long            'Row index
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler
    'Resolve the grid path (robust to OneDrive / SharePoint, where
    'ThisWorkbook.Path returns an http URL that Open cannot use)
        Path = ResolveGridPath()
        If Len(Path) = 0 Then Exit Sub          'User cancelled the picker
        Filled = 0
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
' EVALUATE EACH ROW
'------------------------------------------------------------------------------
    'Row 0 is the header; data starts at row 1
        For I = 1 To UBound(Lines)
            If Len(Trim$(Lines(I))) = 0 Then GoTo ContinueRow

            Cols = Split(Lines(I), Sep)
            If UBound(Cols) < 13 Then GoTo ContinueRow

            A = ParseDouble(Cols(1))
            X = ParseDouble(Cols(2))

            Cols(5) = FormatFullPrecision(A)
            Cols(6) = FormatFullPrecision(X)
            Answered = ""
            Cols(7) = RoutineStatus("series", A, X, V, Msg)
            Cols(8) = IIf(Cols(7) = "OK", FormatFullPrecision(V), "")
            If Cols(7) = "OK" Then Answered = "series"
            Note = "series: " & Msg
            Cols(9) = RoutineStatus("cf", A, X, V, Msg)
            Cols(10) = IIf(Cols(9) = "OK", FormatFullPrecision(V), "")
            If Cols(9) = "OK" And Len(Answered) = 0 Then Answered = "cf"
            Note = Note & "; cf: " & Msg
            If Len(Answered) > 0 Then
                Cols(11) = Replace(Format$(TimeRoutine(Answered, A, X, Calls), "0.000"), ",", ".")
                Cols(12) = CStr(Calls)
            Else
                Cols(11) = ""
                Cols(12) = ""
            End If
            Cols(13) = Replace(Note, ",", ";")
            Lines(I) = Join(Cols, Sep)
            Filled = Filled + 1
ContinueRow:
        Next I
'------------------------------------------------------------------------------
' WRITE BACK
'------------------------------------------------------------------------------
        OutNum = FreeFile
        Open Path For Output As #OutNum
        For I = 0 To UBound(Lines)
            If I < UBound(Lines) Or Len(Lines(I)) > 0 Then Print #OutNum, Lines(I)
        Next I
        Close #OutNum

    MsgBox "Gamma binary64 route fit complete: " & Filled & _
           " observation(s) written to" & vbCrLf & Path, _
           vbInformation, "Gamma binary64 route fit"
    Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    On Error Resume Next
    Close #FileNum
    Close #OutNum
    MsgBox "Gamma binary64 route fit failed: " & Err.Description, _
           vbExclamation, "Gamma binary64 route fit"
End Sub


Private Function RoutineStatus( _
    ByVal Routine As String, _
    ByVal A As Double, _
    ByVal X As Double, _
    ByRef Value As Double, _
    ByRef Msg As String) _
    As String
'
'==============================================================================
' RoutineStatus
'------------------------------------------------------------------------------
' PURPOSE
'   Calls one repaired binary64 routine and returns OK, DECLINED or ERROR.
'   Value and Msg carry its result and diagnostic.
'
' INPUTS
'   Routine     "series" for PROB_TryGammaB64SeriesP, "cf" for
'               PROB_TryGammaB64ContinuedFractionQ
'   A, X        kernel shape and argument
'==============================================================================
'
    On Error GoTo Err_Handler
    Value = 0#: Msg = ""
    If Routine = "series" Then
        RoutineStatus = IIf(PROB_TryGammaB64SeriesP(A, X, Value, Msg), "OK", "DECLINED")
    Else
        RoutineStatus = IIf(PROB_TryGammaB64ContinuedFractionQ(A, X, Value, Msg), "OK", "DECLINED")
    End If
    Exit Function

Err_Handler:
    Msg = "runtime error " & CStr(Err.Number)
    RoutineStatus = "ERROR"
End Function


Private Function TimeRoutine( _
    ByVal Routine As String, _
    ByVal A As Double, _
    ByVal X As Double, _
    ByRef Calls As Long) _
    As Double
'
'==============================================================================
' TimeRoutine
'------------------------------------------------------------------------------
' PURPOSE
'   Mean microseconds per call of one routine at (A, X): the count doubles
'   from 16 until at least 0.2 s has elapsed, as the #298 cost benchmark does.
'==============================================================================
'
    Dim K As Long, T0 As Double, Elapsed As Double, V As Double, M As String, OK As Boolean
    Calls = 16
    Do
        T0 = Timer
        For K = 1 To Calls
            If Routine = "series" Then
                OK = PROB_TryGammaB64SeriesP(A, X, V, M)
            Else
                OK = PROB_TryGammaB64ContinuedFractionQ(A, X, V, M)
            End If
        Next K
        Elapsed = Timer - T0
        If Elapsed < 0# Then Elapsed = Elapsed + 86400#
        If Elapsed >= 0.2 Or Calls >= 536870912 Then Exit Do
        Calls = Calls * 2
    Loop
    TimeRoutine = Elapsed / Calls * 1000000#
End Function


Private Function ResolveGridPath() As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a usable LOCAL path to b64_fit_grid.csv, or an
'   empty string if the user cancels.
'
' WHY THIS EXISTS
'   ThisWorkbook.Path returns an http(s) URL when the workbook lives on OneDrive
'   or SharePoint, and Open cannot read a URL. This prefers a local workbook
'   folder and otherwise asks the user.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim BookPath            As String          'Workbook folder
    Dim Candidate           As String          'Path next to the workbook
    Dim Picked              As Variant          'File-dialog result
'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    BookPath = ThisWorkbook.Path
    If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
        Candidate = BookPath & Application.PathSeparator & _
                    "b64_fit_grid.csv"
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate b64_fit_grid.csv automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate Gamma binary64 route fit"
    Picked = Application.GetOpenFilename( _
        FileFilter:="Gamma binary64 route fit (*.csv),*.csv", _
        Title:="Select b64_fit_grid.csv")

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
'   Renders a Double as a two-part sum "hi;lo", so hi + lo summed in Double
'   precision on the Python side reproduces the original Double exactly.
'
' WHY TWO PARTS
'   Format$, Str$ and CDec all cap a Double at about 15 significant digits,
'   which is coarser than the accuracy this study measures. Writing the residual
'   X - hi as a second field carries the low-order bits hi dropped.
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
'   Formats X to 15 significant digits in scientific notation with a US decimal
'   point, whatever the local settings are.
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
