Attribute VB_Name = "M_STATS_PROBDIST_DELTA_SEAM"
Option Explicit

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
'       LogBeta_ident   -> PROB_LogGamma(Large) + PROB_LogGamma(Small)
'                          - PROB_LogGamma(Large + Small)
'       LogBeta_stable  -> PROB_LogGamma(Small) - PROB_LogGammaDelta(Large, Small)
'
' USAGE
'   Run Export_Delta_Seam and pick delta_seam_grid.csv in the dialog.
'
' GRID FORMAT (header row, then one row per evaluation)
'   quantity, arg1, arg2, reference, observed_vba
'
'   arg1 is the Large argument and arg2 the Small one. This macro writes
'   observed_vba (column index 4) and never reads the reference column, so the
'   observed side stays independent.
'
' ERROR POLICY
'   The PROB_ kernels return Double, so there is no CVErr path here; a runtime
'   fault is written as the token ERROR and the analysis flags that row.
'
' DEPENDENCIES
'   - PROB_LogGamma, PROB_LogGammaDelta (M_STATS_PROBDIST_SPECIALFUNCS)
'
' UPDATED
'   2026-07-23
'==============================================================================


Public Sub Export_Delta_Seam()
'
'==============================================================================
' Export_Delta_Seam
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the observed_vba column of the seam-study grid in place.
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
    Dim A1                  As Double          'Large argument (arg1)
    Dim A2                  As Double          'Small argument (arg2)
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
            If UBound(Cols) < 4 Then GoTo ContinueRow

            A1 = ParseDouble(Cols(1))
            A2 = ParseDouble(Cols(2))

            Cols(4) = EvaluateQuantity(Trim$(Cols(0)), A1, A2)
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

    MsgBox "Delta seam study complete: " & Filled & _
           " observation(s) written to" & vbCrLf & Path, _
           vbInformation, "Delta seam study"
    Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    On Error Resume Next
    Close #FileNum
    Close #OutNum
    MsgBox "Delta seam study failed: " & Err.Description, vbExclamation, _
           "Delta seam study"
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
'   Dispatches the three measured quantities and returns a full-precision token.
'
' INPUTS
'   Quantity    quantity name from the grid
'   LargeArg    the Large argument
'   SmallArg    the Small argument
'
' RETURNS
'   Full-precision hi;lo token, or the literal ERROR on a runtime fault.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Value               As Double          'Measured quantity
'------------------------------------------------------------------------------
' DISPATCH
'------------------------------------------------------------------------------
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
'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    EvaluateQuantity = FormatFullPrecision(Value)
    Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    EvaluateQuantity = "ERROR"
End Function


Private Function ResolveGridPath() As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a usable LOCAL path to delta_seam_grid.csv, or an empty string if
'   the user cancels.
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


