Attribute VB_Name = "M_STATS_PROBDIST_LGREGIMES"
Option Explicit

'==============================================================================
' M_STATS_PROBDIST_LGREGIMES
'------------------------------------------------------------------------------
' PURPOSE
'   Self-contained export macro for the PROB_LogGamma regime study. Fills the
'   observed_vba column of loggamma_regimes_grid.csv, measuring:
'
'       EchoZ     -> Z itself, round-tripped through the parser
'       LogGamma  -> PROB_LogGamma(Z)
'
' WHY THIS STUDY EXISTS
'   The module has advertised one global claim, relative error below 6.1E-14
'   across Z in [1E-8, 1E+50]. That claim fails in two independent ways, and
'   this grid is the evidence for replacing it with regime-aware contracts:
'
'     A. The reflection path forms Sin(PROB_PI * Z) after PROB_PI * Z has
'        entered the subnormal range and lost significand bits, giving about
'        4.6E-02 of absolute log error at the smallest positive Double.
'     B. Log(Gamma(Z)) is zero at Z = 1 and Z = 2, so a global RELATIVE
'        contract is ill-conditioned by construction.
'
' RUN IT TWICE
'   Export once BEFORE the Phase 1 edit and keep the result as
'   loggamma_regimes_baseline.csv, then again after. No committed observation
'   covers Z below 1E-8, so without that baseline the subnormal improvement has
'   nothing to be measured against.
'
' USAGE
'   Run Export_LogGammaRegimes and pick loggamma_regimes_grid.csv in the dialog.
'
' GRID FORMAT (header row, then one row per evaluation)
'   quantity, regime, metric, arg1, bits, reference, observed_vba
'
'   arg1 is Z. This macro writes observed_vba (column index 6) and never reads
'   the reference column, so the observed side stays independent.
'
' ERROR POLICY
'   PROB_LogGamma returns Double and has no CVErr path; a runtime fault is
'   written as the token ERROR and the analysis flags that row. Z = 1 and Z = 2
'   are legitimate points that return exactly zero.
'
' DEPENDENCIES
'   - PROB_LogGamma (M_STATS_PROBDIST_SPECIALFUNCS)
'
' UPDATED
'   2026-08-12
'==============================================================================


Public Sub Export_LogGammaRegimes()
'
'==============================================================================
' Export_LogGammaRegimes
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the observed_vba column of the LogGamma regimes study grid in place.
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
            If UBound(Cols) < 6 Then GoTo ContinueRow

            A1 = ParseDouble(Cols(3))

            Cols(6) = EvaluateQuantity(Trim$(Cols(0)), A1)
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

    MsgBox "LogGamma regimes study complete: " & Filled & _
           " observation(s) written to" & vbCrLf & Path, _
           vbInformation, "LogGamma regimes study"
    Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    On Error Resume Next
    Close #FileNum
    Close #OutNum
    MsgBox "LogGamma regimes study failed: " & Err.Description, vbExclamation, _
           "LogGamma regimes study"
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
'   X           the LogGamma argument Z
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
        Case "EchoZ"
            Value = X

        Case "LogGamma"
            Value = PROB_LogGamma(X)

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
'   Returns a usable LOCAL path to loggamma_regimes_grid.csv, or an empty string if
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
        Candidate = BookPath & Application.PathSeparator & "loggamma_regimes_grid.csv"
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate loggamma_regimes_grid.csv automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate LogGamma regimes grid"
    Picked = Application.GetOpenFilename( _
        FileFilter:="LogGamma regimes grid (*.csv),*.csv", _
        Title:="Select loggamma_regimes_grid.csv")

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
'   this reason; see the EchoZ quantity.
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
