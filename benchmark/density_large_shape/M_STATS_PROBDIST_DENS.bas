Attribute VB_Name = "M_STATS_PROBDIST_DENS"
Option Explicit

'==============================================================================
' M_STATS_PROBDIST_DENS
'------------------------------------------------------------------------------
' PURPOSE
'   Self-contained export macro for the large-shape density study (CR-P1-01).
'   Fills observed_vba in density_large_shape_grid.csv for the Gamma, Chi-square,
'   Beta and F densities across shape/df from 1E2 to 1E20, so the boundary where
'   the current (cancellation-prone) log-density form loses accuracy can be
'   measured from real VBA and set beside the stable Loader target.
'
' USAGE
'   Run Export_DensityLargeShape and pick density_large_shape_grid.csv.
'
' GRID FORMAT (study grid; 11 columns, no arg4)
'   function, vba_kernel, claim, metric, arg1, arg2, arg3, reference,
'   observed_vba, regime, evidence_set
'
'   arg1 is X, arg2 and arg3 are the shape/df parameters (arg3 is empty for the
'   two-parameter Chi-square density). This macro writes observed_vba (column
'   index 8) and never reads the reference column, so the observed side stays
'   independent. This is the study grid, not the 12-column main grid.
'
' ERROR POLICY
'   The K_STATS_ functions return Variant and may return CVErr. A CVErr, a
'   non-numeric result or a runtime fault is written as the token ERROR so the
'   analysis flags the row instead of counting it as a pass.
'
' DEPENDENCIES
'   - K_STATS_Gamma_Density, K_STATS_ChiSquare_Density,
'     K_STATS_Beta_Density, K_STATS_F_Density
'
' UPDATED
'   2026-07-25
'==============================================================================


Public Sub Export_DensityLargeShape()
'
'==============================================================================
' Export_DensityLargeShape
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the observed_vba column of the large-shape density grid in place.
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
    Dim A1                  As Double          'Evaluation point X (arg1)
    Dim A2                  As Double          'First shape / df (arg2)
    Dim A3                  As Double          'Second shape / df (arg3)
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

            A1 = ParseDouble(Cols(4))
            A2 = ParseDouble(Cols(5))
            A3 = ParseDouble(Cols(6))

            Cols(8) = EvaluateDensity(Trim$(Cols(0)), A1, A2, A3)
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

    MsgBox "Large-shape density study complete: " & Filled & _
           " observation(s) written to" & vbCrLf & Path, _
           vbInformation, "Density large-shape"
    Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    On Error Resume Next
    Close #FileNum
    Close #OutNum
    MsgBox "Large-shape density study failed: " & Err.Description, _
           vbExclamation, "Density large-shape"
End Sub


Private Function EvaluateDensity( _
    ByVal FuncName As String, _
    ByVal A1 As Double, _
    ByVal A2 As Double, _
    ByVal A3 As Double) _
    As String
'
'==============================================================================
' EvaluateDensity
'------------------------------------------------------------------------------
' PURPOSE
'   Dispatches the public density functions and returns a string token: a
'   full-precision number on success, or ERROR on any error. Chi-square takes
'   two parameters (X, df); the others take three (X and two shapes/df).
'
' INPUTS
'   FuncName    contract function name from the grid
'   A1          evaluation point X
'   A2, A3      shape / degrees-of-freedom parameters (A3 unused by Chi-square)
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
        Case "Gamma_Density":      V = K_STATS_Gamma_Density(A1, A2, A3)
        Case "ChiSquare_Density":  V = K_STATS_ChiSquare_Density(A1, A2)
        Case "Beta_Density":       V = K_STATS_Beta_Density(A1, A2, A3)
        Case "F_Density":          V = K_STATS_F_Density(A1, A2, A3)

        Case Else
            EvaluateDensity = "ERROR"
            Exit Function
    End Select
'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    If IsError(V) Then
        EvaluateDensity = "ERROR"
    ElseIf Not IsNumeric(V) Then
        EvaluateDensity = "ERROR"
    Else
        EvaluateDensity = FormatFullPrecision(CDbl(V))
    End If
    Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    EvaluateDensity = "ERROR"
End Function


Private Function ResolveGridPath() As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a usable LOCAL path to density_large_shape_grid.csv, or an empty
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
    Dim Picked              As Variant          'File-dialog result
'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    BookPath = ThisWorkbook.Path
    If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
        Candidate = BookPath & Application.PathSeparator & "density_large_shape_grid.csv"
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate density_large_shape_grid.csv automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate density grid"
    Picked = Application.GetOpenFilename( _
        FileFilter:="Density grid (*.csv),*.csv", _
        Title:="Select density_large_shape_grid.csv")

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
'   local list/decimal separators. An empty field (Chi-square arg3) reads zero.
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
