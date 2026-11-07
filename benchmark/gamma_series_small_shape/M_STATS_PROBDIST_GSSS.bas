Attribute VB_Name = "M_STATS_PROBDIST_GSSS"
Option Explicit

'==============================================================================
' M_STATS_PROBDIST_GSSS
'------------------------------------------------------------------------------
' PURPOSE
'   Export macro for the #23 small-shape incomplete-gamma study. Fills the
'   observed side of gamma_series_small_shape_grid.csv: the frozen holdout of
'   HOLDOUT_DESIGN.md (Gamma and Chi-square cumulative routes to the same
'   kernel), its scope probes and the issue's minimal counterexample.
'
' USAGE
'   Run Export_GammaSeriesSmallShape. It finds the grid beside the workbook,
'   or asks for it. ci/Export-ExcelGrids.ps1 runs it on the self-hosted
'   runner with the grid staged beside the export workbook (#225).
'
' GRID FORMAT (study grid; 13 columns)
'   function, vba_kernel, claim, metric, arg1, arg2, arg3, reference, regime,
'   evidence_set, echo_arg1, echo_arg2, observed_vba
'
'   arg1 is X, arg2 the shape (Gamma) or the degrees of freedom (Chi-square),
'   arg3 the Gamma scale. This macro writes the last three columns: the
'   binary64 values it parsed from arg1 and arg2, and the result. It never
'   reads the reference column, so the observed side stays independent. The
'   echoes let the analyzer confirm that each reference is evaluated at the
'   binary64 inputs the function received.
'
' ERROR POLICY
'   The K_STATS_ functions return Variant and may return CVErr. A CVErr, a
'   non-numeric result or a runtime fault is written as the token ERROR so the
'   analysis flags the row instead of counting it as a pass.
'
' DEPENDENCIES
'   - K_STATS_Gamma_Cumulative          (M_STATS_PROBDIST_CONTINUOUS)
'   - K_STATS_ChiSquare_Cumulative      (M_STATS_PROBDIST_TFAMILY)
'
' UPDATED
'   2026-11-07
'==============================================================================


Public Sub Export_GammaSeriesSmallShape()
'
'==============================================================================
' Export_GammaSeriesSmallShape
'------------------------------------------------------------------------------
' PURPOSE
'   Fills echo_arg1, echo_arg2 and observed_vba of the #23 study grid in place.
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
    Dim A1                  As Double          'X (arg1)
    Dim A2                  As Double          'Shape or degrees of freedom (arg2)
    Dim A3                  As Double          'Gamma scale (arg3)
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
            If UBound(Cols) < 12 Then GoTo ContinueRow

            A1 = ParseDouble(Cols(4))
            A2 = ParseDouble(Cols(5))
            A3 = ParseDouble(Cols(6))

            Cols(10) = FormatFullPrecision(A1)
            Cols(11) = FormatFullPrecision(A2)
            Cols(12) = EvaluateGammaSeries(Trim$(Cols(0)), A1, A2, A3)
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

    MsgBox "Gamma small-shape study complete: " & Filled & _
           " observation(s) written to" & vbCrLf & Path, _
           vbInformation, "Gamma small-shape study"
    Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    On Error Resume Next
    Close #FileNum
    Close #OutNum
    MsgBox "Gamma small-shape study failed: " & Err.Description, _
           vbExclamation, "Gamma small-shape study"
End Sub


Private Function EvaluateGammaSeries( _
    ByVal FuncName As String, _
    ByVal A1 As Double, _
    ByVal A2 As Double, _
    ByVal A3 As Double) _
    As String
'
'==============================================================================
' EvaluateGammaSeries
'------------------------------------------------------------------------------
' PURPOSE
'   Dispatches the public Gamma and Chi-square cumulative functions and returns
'   a string token: a full-precision number on success, or ERROR on any error.
'
' INPUTS
'   FuncName    contract function name from the grid
'   A1          X
'   A2          shape (Gamma) or degrees of freedom (Chi-square)
'   A3          Gamma scale; unused for Chi-square
'
' RETURNS
'   Full-precision hi;lo token, or the literal ERROR.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant         'Raw function result
'------------------------------------------------------------------------------
' DISPATCH
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

    Select Case FuncName
        Case "Gamma_Cumulative":            V = K_STATS_Gamma_Cumulative(A1, A2, A3)
        Case "ChiSquare_Cumulative":        V = K_STATS_ChiSquare_Cumulative(A1, A2)

        Case Else
            EvaluateGammaSeries = "ERROR"
            Exit Function
    End Select
'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    If IsError(V) Then
        EvaluateGammaSeries = "ERROR"
    ElseIf Not IsNumeric(V) Then
        EvaluateGammaSeries = "ERROR"
    Else
        EvaluateGammaSeries = FormatFullPrecision(CDbl(V))
    End If
    Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    EvaluateGammaSeries = "ERROR"
End Function


Private Function ResolveGridPath() As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a usable LOCAL path to gamma_series_small_shape_grid.csv, or an
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
                    "gamma_series_small_shape_grid.csv"
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate gamma_series_small_shape_grid.csv automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate Gamma small-shape grid"
    Picked = Application.GetOpenFilename( _
        FileFilter:="Gamma small-shape grid (*.csv),*.csv", _
        Title:="Select gamma_series_small_shape_grid.csv")

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
