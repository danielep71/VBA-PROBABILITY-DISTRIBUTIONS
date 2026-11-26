Attribute VB_Name = "M_STATS_PROBDIST_XFM"
Option Explicit

'One Double and its two 32-bit words, for exact transport through text
Private Type XFM_Words
    Low                     As Long
    High                    As Long
End Type

Private Type XFM_Number
    Value                   As Double
End Type

'==============================================================================
' M_STATS_PROBDIST_XFM
'------------------------------------------------------------------------------
' PURPOSE
'   Export macro for the map from Excel's distribution functions to the
'   library's (#330, benchmark/excel_function_map/README.md). At each row it
'   evaluates Excel's own formula and the library expression that replaces
'   it, and records both results exactly.
'
' USAGE
'   Export_ExcelFunctionMap fills excel_function_map_grid.csv. It finds the
'   file beside the workbook, or asks for it; ci/Export-ExcelGrids.ps1 runs
'   it on the self-hosted runner with the file staged beside the export
'   workbook.
'
' GRID FORMAT (study grid; 7 columns)
'   point_id, kind, excel_formula, kstats_call, reference,
'   observed_excel, observed_kstats
'
'   The macro writes the last two. The two formulas are stored with ";"
'   between arguments; the macro puts the commas back before it evaluates
'   them. It never reads the reference.
'
' EVALUATION
'   Both formulas go through the Evaluate method of a worksheet of this
'   workbook, so each is computed as a cell would compute it and the
'   library's functions resolve to this workbook's own modules.
'
' ERROR POLICY
'   A formula that returns a worksheet error, or a value that is not a
'   number, is recorded as ERROR. A result is recorded as the 16 hexadecimal
'   digits of its Double.
'
' UPDATED
'   2026-11-16
'==============================================================================


Public Sub Export_ExcelFunctionMap()
'
'==============================================================================
' Export_ExcelFunctionMap
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the two observation columns of excel_function_map_grid.csv.
'==============================================================================
'
    Dim Path                As String          'File written
    Dim Filled              As Long            'Rows written, or -1

    Filled = ExportFile("excel_function_map_grid.csv", "Excel function map", Path)
    If Filled < 0 Then Exit Sub

    MsgBox "Excel function map complete: " & Filled & _
           " observation(s) written to" & vbCrLf & Path, _
           vbInformation, "Excel function map"
End Sub


Public Sub Export_ExcelFloor()

'
'------------------------------------------------------------------------------
'                         EXPORT THE EXCEL FLOOR GRID
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the two observation columns of excel_floor_grid.csv
'
' WHY THIS EXISTS
'   The release rule is that the library is nowhere less accurate than Excel
'   (#493). The floor grid holds, for every comparable point of the certified
'   grids, the library call and each Excel formulation of the same quantity;
'   this macro evaluates both sides as cells would, like the function map
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Locates excel_floor_grid.csv beside the workbook, or asks for it
'   - Evaluates every formula of the grid through ExportFile
'   - Reports the number of rows written
'
' ERROR POLICY
'   ExportFile reports a failure itself and leaves the grid as found; this
'   macro then stops without a message of its own
'
' DEPENDENCIES
'   ExportFile
'
' NOTES
'   ci/Export-ExcelGrids.ps1 runs it on the runner as the study export
'   xfl_floor; benchmark/excel_function_map/floor_check.py judges the result
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'File written
    Dim Filled              As Long            'Rows written, or -1

'------------------------------------------------------------------------------
' EXPORT THE GRID
'------------------------------------------------------------------------------
    'Evaluate every formula of the grid
        Filled = ExportFile("excel_floor_grid.csv", "Excel floor", Path)
    'Stop when the export failed or was cancelled
        If Filled < 0 Then Exit Sub

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Report the rows written
        MsgBox "Excel floor complete: " & Filled & _
            " observation(s) written to" & vbCrLf & Path, _
            vbInformation, "Excel floor"

End Sub


Public Sub Export_PrefactorHoldout()

'
'------------------------------------------------------------------------------
'                      EXPORT THE PREFACTOR HOLDOUT GRID
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the two observation columns of prefactor_holdout_grid.csv
'
' WHY THIS EXISTS
'   The incomplete-beta prefactor rule (#175) is judged on points no source
'   has evaluated in Excel, and against Excel's own functions at the same
'   points. The grid has the format of the Excel floor grid, so the same
'   evaluator fills both sides
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Locates prefactor_holdout_grid.csv beside the workbook, or asks for it
'   - Evaluates every formula of the grid through ExportFile
'   - Reports the number of rows written
'
' ERROR POLICY
'   ExportFile reports a failure itself and leaves the grid as found; this
'   macro then stops without a message of its own
'
' DEPENDENCIES
'   ExportFile
'
' NOTES
'   ci/Export-ExcelGrids.ps1 runs it on the runner as the study export
'   xfp_prefactor; benchmark/beta_prefactor_study/prefactor_study.py judges
'   the result
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'File written
    Dim Filled              As Long            'Rows written, or -1

'------------------------------------------------------------------------------
' EXPORT THE GRID
'------------------------------------------------------------------------------
    'Evaluate every formula of the grid
        Filled = ExportFile("prefactor_holdout_grid.csv", "Prefactor holdout", Path)
    'Stop when the export failed or was cancelled
        If Filled < 0 Then Exit Sub

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Report the rows written
        MsgBox "Prefactor holdout complete: " & Filled & _
            " observation(s) written to" & vbCrLf & Path, _
            vbInformation, "Prefactor holdout"

End Sub


Public Sub Export_PrefactorHoldout2()

'
'------------------------------------------------------------------------------
'                   EXPORT THE SECOND PREFACTOR HOLDOUT GRID
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the two observation columns of prefactor_holdout2_grid.csv
'
' WHY THIS EXISTS
'   The first candidate of the incomplete-beta prefactor rule (#175) was
'   rejected on the first holdout, which is revealed since. The second
'   registration is judged on points of its own, again against Excel's
'   functions at the same points
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Locates prefactor_holdout2_grid.csv beside the workbook, or asks for it
'   - Evaluates every formula of the grid through ExportFile
'   - Reports the number of rows written
'
' ERROR POLICY
'   ExportFile reports a failure itself and leaves the grid as found; this
'   macro then stops without a message of its own
'
' DEPENDENCIES
'   ExportFile
'
' NOTES
'   ci/Export-ExcelGrids.ps1 runs it on the runner as the study export
'   xfp_prefactor2; benchmark/beta_prefactor_study/prefactor_study2.py judges
'   the result
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'File written
    Dim Filled              As Long            'Rows written, or -1

'------------------------------------------------------------------------------
' EXPORT THE GRID
'------------------------------------------------------------------------------
    'Evaluate every formula of the grid
        Filled = ExportFile("prefactor_holdout2_grid.csv", "Prefactor holdout 2", Path)
    'Stop when the export failed or was cancelled
        If Filled < 0 Then Exit Sub

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Report the rows written
        MsgBox "Prefactor holdout 2 complete: " & Filled & _
            " observation(s) written to" & vbCrLf & Path, _
            vbInformation, "Prefactor holdout 2"

End Sub


Public Sub Export_DeviancePairHoldout()

'
'------------------------------------------------------------------------------
'                    EXPORT THE DEVIANCE PAIR HOLDOUT GRID
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the two observation columns of deviance_pair_holdout_grid.csv
'
' WHY THIS EXISTS
'   The rule that carries the Loader deviance as a pair of Doubles (#498) is
'   judged on points no source has evaluated in Excel, and against Excel's
'   own functions at the same points. The grid has the format of the Excel
'   floor grid, so the same evaluator fills both sides
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Locates deviance_pair_holdout_grid.csv beside the workbook, or asks for it
'   - Evaluates every formula of the grid through ExportFile
'   - Reports the number of rows written
'
' ERROR POLICY
'   ExportFile reports a failure itself and leaves the grid as found; this
'   macro then stops without a message of its own
'
' DEPENDENCIES
'   ExportFile
'
' NOTES
'   ci/Export-ExcelGrids.ps1 runs it on the runner as the study export
'   xdp_pair; benchmark/deviance_pair_study/pair_study.py judges
'   the result
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'File written
    Dim Filled              As Long            'Rows written, or -1

'------------------------------------------------------------------------------
' EXPORT THE GRID
'------------------------------------------------------------------------------
    'Evaluate every formula of the grid
        Filled = ExportFile("deviance_pair_holdout_grid.csv", _
            "Deviance pair holdout", Path)
    'Stop when the export failed or was cancelled
        If Filled < 0 Then Exit Sub

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Report the rows written
        MsgBox "Deviance pair holdout complete: " & Filled & _
            " observation(s) written to" & vbCrLf & Path, _
            vbInformation, "Deviance pair holdout"

End Sub


Public Sub Export_DensityPairHoldout()

'
'------------------------------------------------------------------------------
'                    EXPORT THE DENSITY PAIR HOLDOUT GRID
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the two observation columns of density_pair_holdout_grid.csv
'
' WHY THIS EXISTS
'   The rule that carries the Beta, F and Student t densities as pairs of
'   Doubles (#501) is judged on points no source has evaluated in Excel, and
'   against Excel's own functions at the same points. The grid has the format
'   of the Excel floor grid, so the same evaluator fills both sides
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Locates density_pair_holdout_grid.csv beside the workbook, or asks for it
'   - Evaluates every formula of the grid through ExportFile
'   - Reports the number of rows written
'
' ERROR POLICY
'   ExportFile reports a failure itself and leaves the grid as found; this
'   macro then stops without a message of its own
'
' DEPENDENCIES
'   ExportFile
'
' NOTES
'   ci/Export-ExcelGrids.ps1 runs it on the runner as the study export
'   xdn_pair; benchmark/density_pair_study/density_study.py judges
'   the result
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'File written
    Dim Filled              As Long            'Rows written, or -1

'------------------------------------------------------------------------------
' EXPORT THE GRID
'------------------------------------------------------------------------------
    'Evaluate every formula of the grid
        Filled = ExportFile("density_pair_holdout_grid.csv", _
            "Density pair holdout", Path)
    'Stop when the export failed or was cancelled
        If Filled < 0 Then Exit Sub

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Report the rows written
        MsgBox "Density pair holdout complete: " & Filled & _
            " observation(s) written to" & vbCrLf & Path, _
            vbInformation, "Density pair holdout"

End Sub


Public Sub Export_HypergeometricPairHoldout()

'
'------------------------------------------------------------------------------
'                 EXPORT THE HYPERGEOMETRIC PAIR HOLDOUT GRID
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the two observation columns of hypergeometric_pair_holdout_grid.csv
'
' WHY THIS EXISTS
'   The rule that carries the Hypergeometric mass function as pairs of
'   Doubles (#496) is judged on points no source has evaluated in Excel, and
'   against Excel's own functions at the same points. The grid has the format
'   of the Excel floor grid, so the same evaluator fills both sides
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Locates hypergeometric_pair_holdout_grid.csv beside the workbook, or
'     asks for it
'   - Evaluates every formula of the grid through ExportFile
'   - Reports the number of rows written
'
' ERROR POLICY
'   ExportFile reports a failure itself and leaves the grid as found; this
'   macro then stops without a message of its own
'
' DEPENDENCIES
'   ExportFile
'
' NOTES
'   ci/Export-ExcelGrids.ps1 runs it on the runner as the study export
'   xhg_pair; benchmark/hypergeometric_pair_study/hyper_study.py judges
'   the result
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'File written
    Dim Filled              As Long            'Rows written, or -1

'------------------------------------------------------------------------------
' EXPORT THE GRID
'------------------------------------------------------------------------------
    'Evaluate every formula of the grid
        Filled = ExportFile("hypergeometric_pair_holdout_grid.csv", _
            "Hypergeometric pair holdout", Path)
    'Stop when the export failed or was cancelled
        If Filled < 0 Then Exit Sub

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Report the rows written
        MsgBox "Hypergeometric pair holdout complete: " & Filled & _
            " observation(s) written to" & vbCrLf & Path, _
            vbInformation, "Hypergeometric pair holdout"

End Sub


Private Function ExportFile( _
    ByVal FileName As String, _
    ByVal Title As String, _
    ByRef Path As String) _
    As Long
'
'==============================================================================
' ExportFile
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the observation columns of the grid file, row by row. Returns the
'   number of rows written, or -1 when the file was not chosen or the export
'   failed; the failure is reported here, the success by the caller.
'
' BEHAVIOR
'   - Reads the whole file, evaluates every data row, rewrites the file.
'   - A row with fewer than 7 fields is left as it is and not counted.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Lines()             As String          'File lines
    Dim Raw                 As String          'File contents
    Dim FileNum             As Integer         'Input file handle
    Dim OutNum              As Integer         'Output file handle
    Dim Cols                As Variant         'Split fields of one row
    Dim Filled              As Long            'Rows written
    Dim I                   As Long            'Row index
'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

        Path = ResolveGridPath(FileName, Title)
        ExportFile = -1
        If Len(Path) = 0 Then Exit Function     'User cancelled the picker

        Filled = 0

        FileNum = FreeFile
        Open Path For Input As #FileNum
        Raw = Input$(LOF(FileNum), FileNum)
        Close #FileNum

        Raw = Replace(Raw, vbCrLf, vbLf)
        Raw = Replace(Raw, vbCr, vbLf)
        Lines = Split(Raw, vbLf)
'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Row 0 is the header; data starts at row 1
        For I = 1 To UBound(Lines)
            If Len(Trim$(Lines(I))) = 0 Then GoTo ContinueRow

            Cols = Split(Lines(I), ",")
            If UBound(Cols) < 6 Then GoTo ContinueRow

            Cols(5) = Observed(CStr(Cols(2)))
            Cols(6) = Observed(CStr(Cols(3)))

            Lines(I) = Join(Cols, ",")
            Filled = Filled + 1
ContinueRow:
        Next I
'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        OutNum = FreeFile
        Open Path For Output As #OutNum
        For I = 0 To UBound(Lines)
            If I < UBound(Lines) Or Len(Lines(I)) > 0 Then Print #OutNum, Lines(I)
        Next I
        Close #OutNum

    ExportFile = Filled
    Exit Function

Err_Handler:
    On Error Resume Next
    Close #FileNum
    Close #OutNum
    MsgBox Title & " failed: " & Err.Description, _
           vbExclamation, Title
    ExportFile = -1
End Function


Private Function Observed(ByVal Stored As String) As String
'
'==============================================================================
' Observed
'------------------------------------------------------------------------------
' PURPOSE
'   Evaluates one stored formula as a cell of this workbook would and
'   returns the bits of its result, or ERROR.
'==============================================================================
'
    Dim V                   As Variant         'Result of the formula

    On Error GoTo Failed

    V = ThisWorkbook.Worksheets(1).Evaluate(Replace(Trim$(Stored), ";", ","))

    If IsError(V) Then
        Observed = "ERROR"
    ElseIf VarType(V) = vbDouble Or VarType(V) = vbInteger Or VarType(V) = vbLong Then
        Observed = ToBits(CDbl(V))
    Else
        Observed = "ERROR"
    End If
    Exit Function

Failed:
    Observed = "ERROR"
End Function


Private Function ToBits(ByVal X As Double) As String
'
'==============================================================================
' ToBits
'------------------------------------------------------------------------------
' PURPOSE
'   The 64 bits of X as 16 hexadecimal digits, most significant first.
'==============================================================================
'
    Dim Words               As XFM_Words       'The two halves
    Dim Number              As XFM_Number      'The same eight bytes as a Double

    Number.Value = X
    LSet Words = Number
    ToBits = Right$("00000000" & Hex$(Words.High), 8) & _
             Right$("00000000" & Hex$(Words.Low), 8)
End Function


Private Function ResolveGridPath( _
    ByVal FileName As String, _
    ByVal Title As String) _
    As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Finds the grid file beside the workbook, or asks for it.
'==============================================================================
'
    Dim BookPath            As String          'Workbook folder
    Dim Candidate           As String          'Path next to the workbook
    Dim Picked              As Variant         'File-dialog result

    BookPath = ThisWorkbook.Path

    If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
        Candidate = BookPath & Application.PathSeparator & FileName
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate " & FileName & " automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate " & Title

    Picked = Application.GetOpenFilename( _
        FileFilter:=Title & " (*.csv),*.csv", _
        Title:="Select " & FileName)

    If VarType(Picked) = vbBoolean Then
        ResolveGridPath = vbNullString
    Else
        ResolveGridPath = CStr(Picked)
    End If
End Function
