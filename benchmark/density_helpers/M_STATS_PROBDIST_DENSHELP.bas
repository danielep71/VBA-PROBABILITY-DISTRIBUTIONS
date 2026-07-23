Attribute VB_Name = "M_STATS_PROBDIST_DENSHELP"
Option Explicit

'==============================================================================
' M_STATS_PROBDIST_DENSHELP
'------------------------------------------------------------------------------
' PURPOSE
'   Fills observed_vba for four public UDFs that the main export does not cover
'   (ChiSquare_Density, F_Density, Normal_IntervalProbability and
'   Lognormal_ParametersFromMeanStdDev) in probability_accuracy_grid.csv.
'
' SCOPE
'   Touches ONLY rows whose evidence_set is "density_helpers" and leaves every
'   other observation byte-for-byte untouched, so it is safe to run against the
'   full main grid.
'
' USAGE
'   Run Export_DensityHelpers and select probability_accuracy_grid.csv.
'
' GRID FORMAT (12-column arg4 schema; index in brackets)
'   [0] function     [1] vba_kernel  [2] claim      [3] metric
'   [4] arg1         [5] arg2        [6] arg3       [7] arg4
'   [8] reference    [9] observed_vba               [10] regime
'   [11] evidence_set
'
'   This macro reads function, arg1..arg3, regime and evidence_set, and writes
'   observed_vba only. Because it writes into the SHARED main grid, the header
'   is validated before anything is written: a schema change that moved these
'   columns would otherwise write into the reference column silently.
'
' NOTES
'   - Normal_IntervalProbability is exercised with Mean = 0 and StdDev = arg3.
'   - Lognormal_ParametersFromMeanStdDev returns a 1x2 array; the element is
'     chosen by regime ("param_meanlog" or "param_stddevlog").
'
' ERROR POLICY
'   A CVErr, a non-numeric result or a runtime fault is written as the token
'   ERROR; compute_errors.py treats a non-numeric observed value as unusable.
'
' DEPENDENCIES
'   - K_STATS_ChiSquare_Density, K_STATS_F_Density
'   - K_STATS_Normal_IntervalProbability
'   - K_STATS_Lognormal_ParametersFromMeanStdDev
'
' UPDATED
'   2026-07-23
'==============================================================================

'Column indices in the 12-column main-grid schema
Private Const COL_FUNCTION      As Long = 0
Private Const COL_ARG1          As Long = 4
Private Const COL_ARG2          As Long = 5
Private Const COL_ARG3          As Long = 6
Private Const COL_OBSERVED      As Long = 9
Private Const COL_REGIME        As Long = 10
Private Const COL_EVIDENCE      As Long = 11


Public Sub Export_DensityHelpers()
'
'==============================================================================
' Export_DensityHelpers
'------------------------------------------------------------------------------
' PURPOSE
'   Fills observed_vba for the density_helpers rows of the main grid in place.
'
' BEHAVIOR
'   Reads the whole file and normalizes line endings before splitting, so
'   LF-only, CR-only and CRLF grids all parse. VBA Line Input is CR-delimited
'   and would swallow an entire LF-only file (.gitattributes stores *.csv as
'   eol=lf) as a single line, silently writing nothing.
'
' ERROR POLICY
'   The header is validated first; on mismatch nothing is written. Any failure
'   closes the handles and reports once, leaving the grid as found.
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
' VALIDATE HEADER
'------------------------------------------------------------------------------
    'This macro writes into the shared main grid. If the schema ever moves,
    'refuse rather than write an observation into the wrong column.
        If Not HeaderIsExpected(Lines(0), Sep) Then
            MsgBox "Grid header does not match the expected 12-column schema." & _
                   vbCrLf & "Nothing was written.", vbExclamation, "Density helpers export"
            Exit Sub
        End If
'------------------------------------------------------------------------------
' EVALUATE MATCHING ROWS
'------------------------------------------------------------------------------
    'Row 0 is the header; data starts at row 1
        For I = 1 To UBound(Lines)
            If Len(Trim$(Lines(I))) = 0 Then GoTo ContinueRow

            Cols = Split(Lines(I), Sep)
            If UBound(Cols) < COL_EVIDENCE Then GoTo ContinueRow

            'Own only the density_helpers rows
            If Trim$(Cols(COL_EVIDENCE)) <> "density_helpers" Then GoTo ContinueRow

            Cols(COL_OBSERVED) = EvaluateDensityHelper( _
                Trim$(Cols(COL_FUNCTION)), Trim$(CStr(Cols(COL_REGIME))), _
                ParseDouble(Cols(COL_ARG1)), ParseDouble(Cols(COL_ARG2)), _
                CStr(Cols(COL_ARG3)))
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

    MsgBox "Density-helpers coverage export complete: " & Filled & _
           " observation(s) written to" & vbCrLf & Path, _
           vbInformation, "Density helpers export"
    Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    On Error Resume Next
    Close #FileNum
    Close #OutNum
    MsgBox "Density-helpers export failed: " & Err.Description, vbExclamation, _
           "Density helpers export"
End Sub


Private Function HeaderIsExpected( _
    ByVal HeaderLine As String, _
    ByVal Sep As String) _
    As Boolean
'
'==============================================================================
' HeaderIsExpected
'------------------------------------------------------------------------------
' PURPOSE
'   Confirms the grid header carries the expected column names at the exact
'   indices this macro writes to.
'
' WHY THIS EXISTS
'   An earlier schema change moved observed_vba, regime and evidence_set by one
'   position. The macro kept its old indices, so it silently matched nothing -
'   and had it matched, it would have written an observation into the reference
'   column. Validating the header turns that class of defect into a refusal.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim H                   As Variant         'Split header fields
'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    H = Split(HeaderLine, Sep)
    If UBound(H) < COL_EVIDENCE Then Exit Function

    If Trim$(H(COL_FUNCTION)) <> "function" Then Exit Function
    If Trim$(H(COL_ARG1)) <> "arg1" Then Exit Function
    If Trim$(H(COL_ARG2)) <> "arg2" Then Exit Function
    If Trim$(H(COL_ARG3)) <> "arg3" Then Exit Function
    If Trim$(H(COL_OBSERVED)) <> "observed_vba" Then Exit Function
    If Trim$(H(COL_REGIME)) <> "regime" Then Exit Function
    If Trim$(H(COL_EVIDENCE)) <> "evidence_set" Then Exit Function

    HeaderIsExpected = True
End Function


Private Function EvaluateDensityHelper( _
    ByVal FuncName As String, _
    ByVal Regime As String, _
    ByVal A1 As Double, _
    ByVal A2 As Double, _
    ByVal A3Text As String) _
    As String
'
'==============================================================================
' EvaluateDensityHelper
'------------------------------------------------------------------------------
' PURPOSE
'   Dispatches the four density-helper functions and returns a string token: a
'   full-precision number on success, or ERROR on any worksheet error.
'
' INPUTS
'   FuncName    contract function name from the grid
'   Regime      contract regime; selects the output for the parameter conversion
'   A1, A2      parsed arguments 1 and 2
'   A3Text      argument 3 as text, since two of the four leave it empty
'
' RETURNS
'   Full-precision hi;lo token, or the literal ERROR.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant         'Raw function result
    Dim Arr                 As Variant         'Both lognormal parameters
    Dim Lb1                 As Long            'First dimension base
    Dim Lb2                 As Long            'Second dimension base
    Dim A3                  As Double          'Parsed argument 3
'------------------------------------------------------------------------------
' DISPATCH
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

    Select Case FuncName
        Case "ChiSquare_Density"
            V = K_STATS_ChiSquare_Density(A1, A2)

        Case "F_Density"
            A3 = ParseDouble(A3Text)
            V = K_STATS_F_Density(A1, A2, A3)

        Case "Normal_IntervalProbability"
            A3 = ParseDouble(A3Text)           'StdDev; Mean fixed at 0
            V = K_STATS_Normal_IntervalProbability(A1, A2, 0#, A3)

        Case "Lognormal_ParametersFromMeanStdDev"
            'Returns both parameters; the regime selects which one this row claims
            Arr = K_STATS_Lognormal_ParametersFromMeanStdDev(A1, A2)
            If IsError(Arr) Then
                EvaluateDensityHelper = "ERROR"
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
            EvaluateDensityHelper = "ERROR"
            Exit Function
    End Select
'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    If IsError(V) Then
        EvaluateDensityHelper = "ERROR"
    ElseIf Not IsNumeric(V) Then
        EvaluateDensityHelper = "ERROR"
    Else
        EvaluateDensityHelper = FormatFullPrecision(CDbl(V))
    End If
    Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    EvaluateDensityHelper = "ERROR"
End Function


Private Function ResolveGridPath() As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a usable LOCAL path to probability_accuracy_grid.csv, or an empty
'   string if the user cancels.
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
    Dim Picked              As Variant         'File-dialog result
'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
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
'   which is coarser than the accuracy the harness measures. Writing the
'   residual X - hi as a second field carries the bits hi dropped.
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


