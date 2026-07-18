Attribute VB_Name = "M_STATS_PROBDIST_ACCURACYEXPORT"
Option Explicit

'==============================================================================
' M_STATS_PROBDIST_ACCURACY_EXPORT
'------------------------------------------------------------------------------
' PURPOSE
'   Phase 2 of the reproducible accuracy harness. Reads the grid produced by
'   generate_reference_values.py, evaluates each library function at every grid
'   point, and writes the observed values back so compute_errors.py can measure
'   the error against the mpmath reference.
'
' WHY THIS EXISTS
'   The reference values are generated in Python (mpmath, 50 digits). The
'   library under test is VBA and can only be executed inside Excel. This macro
'   is the bridge: it fills the observed_vba column that Python cannot.
'
' WORKFLOW
'   1. python generate_reference_values.py         -> probability_accuracy_grid.csv
'   2. Place probability_accuracy_grid.csv next to this workbook (or set the path
'      in ACCURACY_GRID_PATH) and run Export_Accuracy_Observations.
'   3. python compute_errors.py                    -> accuracy_summary.md
'
' GRID FORMAT (header row, then one row per evaluation)
'   function, vba_kernel, claim, metric, arg1, arg2, arg3, reference, observed_vba
'
'   This macro reads columns function/arg1..arg3, writes observed_vba. It does
'   not read or trust the reference column, so the observed side is independent.
'
' ERROR POLICY
'   A function that returns a worksheet error (CVErr) or raises writes the token
'   ERROR into observed_vba; compute_errors.py treats a non-numeric observed
'   value as a failed point. Empty inputs are passed as the function defaults.
'
' DEPENDENCIES
'   - M_STATS_PROBDIST_SPECIALFUNCS (PROB_* kernels)
'   - M_STATS_PROBDIST_TFAMILY      (K_STATS_* UDFs)
'
' UPDATED
'   2026-07-18
'==============================================================================

Private Const ACCURACY_GRID_PATH As String = ""        'Empty => same folder as the workbook


Public Sub Export_Accuracy_Observations()
'
'==============================================================================
' Export_Accuracy_Observations
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the observed_vba column of the accuracy grid in place.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'Resolved grid path
    Dim Lines()             As String          'File lines
    Dim Raw                 As String          'File contents
    Dim FileNo              As Integer         'File handle
    Dim I                   As Long            'Row index
    Dim Cols                As Variant         'Split fields of one row
    Dim FuncName            As String          'Function under test
    Dim A1 As Double, A2 As Double, A3 As Double  'Parsed arguments
    Dim HasA1 As Boolean, HasA2 As Boolean, HasA3 As Boolean
    Dim Observed            As String          'Observed value token
    Dim Sep                 As String          'Field separator
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler
    'Resolve the grid path (robust to OneDrive / SharePoint, where
    'ThisWorkbook.Path returns an http URL that Open cannot use)
        Path = ResolveGridPath()
        If Len(Path) = 0 Then Exit Sub          'User cancelled the picker
    'Read the whole file
        FileNo = FreeFile
        Open Path For Input As #FileNo
        Raw = Input$(LOF(FileNo), FileNo)
        Close #FileNo
    'Normalize line endings and split
        Raw = Replace(Raw, vbCrLf, vbLf)
        Raw = Replace(Raw, vbCr, vbLf)
        Lines = Split(Raw, vbLf)
        Sep = ","
'------------------------------------------------------------------------------
' EVALUATE EACH ROW
'------------------------------------------------------------------------------
    'Row 0 is the header; data starts at row 1
        For I = 1 To UBound(Lines)
            If Len(Trim$(Lines(I))) = 0 Then GoTo ContinueRow

            Cols = Split(Lines(I), Sep)
            If UBound(Cols) < 8 Then GoTo ContinueRow

            FuncName = Trim$(Cols(0))
            HasA1 = (Len(Trim$(Cols(4))) > 0)
            HasA2 = (Len(Trim$(Cols(5))) > 0)
            HasA3 = (Len(Trim$(Cols(6))) > 0)
            If HasA1 Then A1 = ParseDouble(Cols(4))
            If HasA2 Then A2 = ParseDouble(Cols(5))
            If HasA3 Then A3 = ParseDouble(Cols(6))

            Observed = EvaluateOne(FuncName, A1, A2, A3, HasA2, HasA3)

            'Rebuild the row with observed_vba (last field) filled
            Cols(8) = Observed
            Lines(I) = Join(Cols, Sep)
ContinueRow:
        Next I
'------------------------------------------------------------------------------
' WRITE BACK
'------------------------------------------------------------------------------
        FileNo = FreeFile
        Open Path For Output As #FileNo
        For I = 0 To UBound(Lines)
            If I < UBound(Lines) Or Len(Lines(I)) > 0 Then Print #FileNo, Lines(I)
        Next I
        Close #FileNo

    MsgBox "Accuracy observations written to:" & vbCrLf & Path & vbCrLf & vbCrLf & _
           "Now run:  python compute_errors.py", vbInformation, "Accuracy export complete"
    Exit Sub
Err_Handler:
    On Error Resume Next
    Close #FileNo
    MsgBox "Accuracy export failed: " & Err.Description, vbExclamation
End Sub


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
'   or SharePoint, and VBA's Open statement cannot read a URL. This resolver
'   prefers an explicit local path, then the workbook folder only when that is a
'   real local path containing the file, and finally falls back to a file picker.
'==============================================================================
'
    Dim Candidate           As String          'Path being tested
    Dim BookPath            As String          'Workbook folder
    Dim Picked              As Variant          'File-dialog result

    '1. Explicit constant wins when it points at a real file
        If Len(ACCURACY_GRID_PATH) > 0 Then
            If Len(Dir$(ACCURACY_GRID_PATH)) > 0 Then
                ResolveGridPath = ACCURACY_GRID_PATH
                Exit Function
            End If
        End If

    '2. Workbook folder, but only if it is a LOCAL path (URLs start with http)
        BookPath = ThisWorkbook.Path
        If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
            Candidate = BookPath & Application.PathSeparator & "probability_accuracy_grid.csv"
            If Len(Dir$(Candidate)) > 0 Then
                ResolveGridPath = Candidate
                Exit Function
            End If
        End If

    '3. Ask the user to locate the file
        MsgBox "Could not locate probability_accuracy_grid.csv automatically " & _
               "(the workbook may be on OneDrive/SharePoint). Please select it.", _
               vbInformation, "Locate accuracy grid"
        Picked = Application.GetOpenFilename( _
            FileFilter:="Accuracy grid (*.csv),*.csv", _
            Title:="Select probability_accuracy_grid.csv")
        If VarType(Picked) = vbBoolean Then
            ResolveGridPath = vbNullString
        Else
            ResolveGridPath = CStr(Picked)
        End If
End Function


Private Function EvaluateOne( _
    ByVal FuncName As String, _
    ByVal A1 As Double, _
    ByVal A2 As Double, _
    ByVal A3 As Double, _
    ByVal HasA2 As Boolean, _
    ByVal HasA3 As Boolean) _
    As String
'
'==============================================================================
' EvaluateOne
'------------------------------------------------------------------------------
' PURPOSE
'   Dispatches one grid row to its library function and returns a string token:
'   a full-precision number on success, or ERROR on any worksheet error.
'==============================================================================
'
    Dim V                   As Variant         'Raw function result

    On Error GoTo Fail

    Select Case FuncName
        Case "LogGamma":                     V = PROB_LogGamma(A1)
        Case "LogGammaHalfDiff":             V = PROB_LogGammaHalfDiff(A1)
        Case "StirlingError":                V = PROB_StirlingError(A1)
        Case "LogChoose":                    V = PROB_LogChoose(A1, A2)

        Case "StudentT_Density":             V = K_STATS_StudentT_Density(A1, A2)
        Case "StudentT_Cumulative":          V = K_STATS_StudentT_Cumulative(A1, A2)
        Case "StudentT_Survival":            V = K_STATS_StudentT_Survival(A1, A2)
        Case "StudentT_InverseCumulative":   V = K_STATS_StudentT_InverseCumulative(A1, A2)

        Case "ChiSquare_Cumulative":         V = K_STATS_ChiSquare_Cumulative(A1, A2)
        Case "ChiSquare_Survival":           V = K_STATS_ChiSquare_Survival(A1, A2)
        Case "ChiSquare_InverseCumulative":  V = K_STATS_ChiSquare_InverseCumulative(A1, A2)

        Case "F_Cumulative":                 V = K_STATS_F_Cumulative(A1, A2, A3)
        Case "F_Survival":                   V = K_STATS_F_Survival(A1, A2, A3)
        Case "F_InverseCumulative":          V = K_STATS_F_InverseCumulative(A1, A2, A3)

        Case Else:                           EvaluateOne = "ERROR": Exit Function
    End Select

    'A worksheet-error Variant is a failed point
        If IsError(V) Then EvaluateOne = "ERROR": Exit Function

    'Full-precision, locale-independent decimal
        EvaluateOne = FormatFullPrecision(CDbl(V))
    Exit Function
Fail:
    EvaluateOne = "ERROR"
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
    Dim S                   As String          'Cleaned token
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
'   Renders a Double as the shortest decimal string (15 to 17 significant
'   digits) that reads back to the identical Double, so no precision is lost
'   when the CSV is parsed in Python.
'
' WHY NOT Format$
'   Format$ and Str$ round a Double to about 15 significant digits before
'   producing the string, which is coarser than several published accuracy
'   claims. This routine builds the mantissa with the Decimal type, which is not
'   capped at 15 digits, and returns the first digit count whose value round-
'   trips exactly through CDbl.
'==============================================================================
'
    Dim Digits              As Long            'Trial significant-digit count
    Dim Candidate           As String          'Formatted candidate

    If X = 0# Then FormatFullPrecision = "0.0E+000": Exit Function

    'Return the shortest representation that round-trips exactly
        For Digits = 15 To 17
            Candidate = SciFormat(X, Digits)
            If CDbl(Candidate) = X Then
                FormatFullPrecision = Candidate
                Exit Function
            End If
        Next Digits

    'Fall back to the fullest representation (extreme exponents may not
    'round-trip because 10 ^ E is not exactly representable there)
        FormatFullPrecision = Candidate
End Function


Private Function SciFormat( _
    ByVal X As Double, _
    ByVal Sig As Long) _
    As String
'
'==============================================================================
' SciFormat
'------------------------------------------------------------------------------
' PURPOSE
'   Formats X in scientific notation with Sig significant digits and a US
'   decimal point, using Decimal arithmetic for the mantissa so the digit count
'   is not capped at 15.
'==============================================================================
'
    Dim Sign                As String          'Leading minus or empty
    Dim Ax                  As Double          'Absolute value
    Dim E                   As Long            'Decimal exponent
    Dim Mant                As Double          'Mantissa in [1, 10)
    Dim MantStr             As String          'Rounded mantissa text

    If X < 0# Then
        Sign = "-": Ax = -X
    Else
        Sign = vbNullString: Ax = X
    End If

    'Decimal exponent, then normalize the mantissa into [1, 10)
        E = Int(Log(Ax) / Log(10#))
        Mant = Ax / (10# ^ E)
        Do While Mant >= 10#
            Mant = Mant / 10#: E = E + 1
        Loop
        Do While Mant < 1#
            Mant = Mant * 10#: E = E - 1
        Loop

    'Round the mantissa to Sig significant digits with Decimal (uncapped)
        MantStr = Replace(CStr(Round(CDec(Mant), Sig - 1)), ",", ".")
        If InStr(MantStr, ".") = 0 Then MantStr = MantStr & ".0"

    SciFormat = Sign & MantStr & "E" & IIf(E >= 0, "+", "-") & Format$(Abs(E), "000")
End Function


