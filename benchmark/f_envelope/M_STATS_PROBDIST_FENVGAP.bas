Attribute VB_Name = "M_STATS_PROBDIST_FENVGAP"
Option Explicit

'==============================================================================
' M_STATS_PROBDIST_FENVGAP
'------------------------------------------------------------------------------
' PURPOSE
'   Self-contained export macro for the F accuracy-envelope study, GAP sweep.
'   Fills observed_vba in f_envelope_gap_grid.csv for F_Cumulative and
'   F_Survival.
'
' WHY THIS EXISTS
'   F_CDF(x; d1, d2) = I_y(d1/2, d2/2) with y = d1*x / (d1*x + d2), so the
'   degrading quantity is the incomplete-beta shape parameter max(d1/2, d2/2).
'   The base sweep covers that parameter from about 1E6 to 5E9. This grid fills
'   the region BELOW it, roughly 50 to 1E6, so the df boundary where F crosses
'   its 1.1E-10 accuracy contract is bracketed from both sides rather than
'   extrapolated. Both orientations are covered (large first vs large second
'   beta parameter), because the continued fraction can behave differently on
'   each side.
'
' USAGE
'   Run Export_FEnvelopeGap and pick f_envelope_gap_grid.csv in the dialog.
'
' GRID FORMAT (study grid; 11 columns, no arg4)
'   function, vba_kernel, claim, metric, arg1, arg2, arg3, reference,
'   observed_vba, regime, evidence_set
'
'   arg1 is x, arg2 is d1 and arg3 is d2. This macro writes observed_vba
'   (column index 8) and never reads the reference column, so the observed side
'   stays independent. Note this is the study grid, not the 12-column main grid.
'
' ERROR POLICY
'   The K_STATS_ functions return Variant and may return CVErr. A CVErr, a
'   non-numeric result or a runtime fault is written as the token ERROR so the
'   analysis flags the row instead of counting it as a pass.
'
' DEPENDENCIES
'   - K_STATS_F_Cumulative, K_STATS_F_Survival
'
' UPDATED
'   2026-07-23
'==============================================================================


Public Sub Export_FEnvelopeGap()
'
'==============================================================================
' Export_FEnvelopeGap
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the observed_vba column of the gap F envelope grid in place.
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
    Dim XVal                As Double          'Evaluation point x (arg1)
    Dim D1                  As Double          'Numerator df (arg2)
    Dim D2                  As Double          'Denominator df (arg3)
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
            If UBound(Cols) < 8 Then GoTo ContinueRow

            XVal = ParseDouble(Cols(4))
            D1 = ParseDouble(Cols(5))
            D2 = ParseDouble(Cols(6))

            Cols(8) = EvaluateF(Trim$(Cols(0)), XVal, D1, D2)
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

    MsgBox "F envelope GAP study complete: " & Filled & _
           " observation(s) written to" & vbCrLf & Path, _
           vbInformation, "F envelope gap"
    Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    On Error Resume Next
    Close #FileNum
    Close #OutNum
    MsgBox "F envelope gap study failed: " & Err.Description, _
           vbExclamation, "F envelope gap"
End Sub


Private Function EvaluateF( _
    ByVal FuncName As String, _
    ByVal XVal As Double, _
    ByVal D1 As Double, _
    ByVal D2 As Double) _
    As String
'
'==============================================================================
' EvaluateF
'------------------------------------------------------------------------------
' PURPOSE
'   Dispatches the two public F functions and returns a string token: a
'   full-precision number on success, or ERROR on any worksheet error.
'
' INPUTS
'   FuncName    contract function name from the grid
'   XVal        evaluation point
'   D1, D2      numerator and denominator degrees of freedom
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
        Case "F_Cumulative":  V = K_STATS_F_Cumulative(XVal, D1, D2)
        Case "F_Survival":    V = K_STATS_F_Survival(XVal, D1, D2)

        Case Else
            EvaluateF = "ERROR"
            Exit Function
    End Select
'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    If IsError(V) Then
        EvaluateF = "ERROR"
    ElseIf Not IsNumeric(V) Then
        EvaluateF = "ERROR"
    Else
        EvaluateF = FormatFullPrecision(CDbl(V))
    End If
    Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    EvaluateF = "ERROR"
End Function


Private Function ResolveGridPath() As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a usable LOCAL path to f_envelope_gap_grid.csv, or an empty string
'   if the user cancels.
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


