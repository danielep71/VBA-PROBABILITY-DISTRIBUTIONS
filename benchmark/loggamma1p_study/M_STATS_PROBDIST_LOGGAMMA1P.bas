Attribute VB_Name = "M_STATS_PROBDIST_LOGGAMMA1P"
Option Explicit

'==============================================================================
' M_STATS_PROBDIST_LOGGAMMA1P
'------------------------------------------------------------------------------
' PURPOSE
'   Self-contained export macro for the LogGamma1p study (ICR-P1-01 prerequisite,
'   issue #12). Fills the observed_vba column of loggamma1p_grid.csv, measuring
'   four quantities per point so the kernel can be validated, the defect it
'   replaces quantified, and the series/Lanczos seam chosen from measured data:
'
'       EchoX            -> X itself, round-tripped through the parser
'       LogGamma1p       -> PROB_TryLogGamma1p(X)
'       LogGamma1pOverX  -> PROB_TryLogGamma1p(X) / X
'       LogGammaNaive    -> PROB_LogGamma(1# + X)
'
' WHY EchoX EXISTS
'   The grid carries subnormal arguments down to the smallest positive Double.
'   If Val() mis-parses one of those literals, every other number in the row is
'   measuring an argument nobody asked for, and the study would report a kernel
'   defect that is really a parser defect. EchoX makes the parsed Double itself
'   an observation, so the analysis can rule that out before reading anything
'   else.
'
' WHY LogGamma1pOverX IS MEASURED HERE
'   The contract metric is the SCALED error, Abs(observed - reference) / X,
'   because the scaled Gamma inverse computes
'   [LogProbability + LogGamma1p(Shape)] / Shape. Dividing in VBA rather than in
'   Python captures the rounding of the division the caller will actually incur.
'
' USAGE
'   Run Export_LogGamma1p and pick loggamma1p_grid.csv in the dialog.
'
' GRID FORMAT (header row, then one row per evaluation)
'   quantity, regime, arg1, reference, observed_vba
'
'   arg1 is X. This macro writes observed_vba (column index 4) and never reads
'   the reference column, so the observed side stays independent.
'
' ERROR POLICY
'   PROB_TryLogGamma1p returns Boolean; a FALSE return and any runtime fault are
'   both written as the token ERROR, and the analysis flags that row.
'
' DEPENDENCIES
'   - PROB_TryLogGamma1p, PROB_LogGamma (M_STATS_PROBDIST_SPECIALFUNCS)
'
' UPDATED
'   2026-08-12
'==============================================================================


Public Sub Export_LogGamma1p()
'
'==============================================================================
' Export_LogGamma1p
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the observed_vba column of the LogGamma1p study grid in place.
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

            A1 = ParseDouble(Cols(2))

            Cols(4) = EvaluateQuantity(Trim$(Cols(0)), A1)
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

    MsgBox "LogGamma1p study complete: " & Filled & _
           " observation(s) written to" & vbCrLf & Path, _
           vbInformation, "LogGamma1p study"
    Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    On Error Resume Next
    Close #FileNum
    Close #OutNum
    MsgBox "LogGamma1p study failed: " & Err.Description, vbExclamation, _
           "LogGamma1p study"
End Sub


Private Function EvaluateQuantity( _
    ByVal Quantity As String, _
    ByVal X As Double) _
    As String
'
'==============================================================================
' EvaluateQuantity
'------------------------------------------------------------------------------
' PURPOSE
'   Dispatches the four measured quantities and returns a full-precision token.
'
' INPUTS
'   Quantity    quantity name from the grid
'   X           the increment above one
'
' RETURNS
'   Full-precision hi;lo token, or the literal ERROR on a FALSE return or a
'   runtime fault.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Value               As Double          'Measured quantity
    Dim Kernel              As Double          'PROB_TryLogGamma1p result
    Dim FailMsg             As String          'Kernel diagnostic
'------------------------------------------------------------------------------
' DISPATCH
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

    Select Case Quantity
        Case "EchoX"
            Value = X

        Case "LogGamma1p"
            If Not PROB_TryLogGamma1p(X, Kernel, FailMsg) Then
                EvaluateQuantity = "ERROR"
                Exit Function
            End If
            Value = Kernel

        Case "LogGamma1pOverX"
            If Not PROB_TryLogGamma1p(X, Kernel, FailMsg) Then
                EvaluateQuantity = "ERROR"
                Exit Function
            End If
            If X = 0# Then
                EvaluateQuantity = "ERROR"
                Exit Function
            End If
            Value = Kernel / X

        Case "LogGammaNaive"
            Value = PROB_LogGamma(1# + X)

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
'   Returns a usable LOCAL path to loggamma1p_grid.csv, or an empty string if
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
        Candidate = BookPath & Application.PathSeparator & "loggamma1p_grid.csv"
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate loggamma1p_grid.csv automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate LogGamma1p grid"
    Picked = Application.GetOpenFilename( _
        FileFilter:="LogGamma1p grid (*.csv),*.csv", _
        Title:="Select loggamma1p_grid.csv")

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
'
' LIMITATION
'   At subnormal magnitudes the residual X - hi is itself below the
'   representable grid, so the pair cannot round-trip exactly there. The
'   analysis compares arguments as Doubles rather than as decimals for exactly
'   this reason; see the EchoX quantity.
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
