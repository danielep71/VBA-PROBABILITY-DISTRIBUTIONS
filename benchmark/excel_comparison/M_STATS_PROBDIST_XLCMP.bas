Attribute VB_Name = "M_STATS_PROBDIST_XLCMP"
Option Explicit
'
'==============================================================================
' M_STATS_PROBDIST_XLCMP
'------------------------------------------------------------------------------
' PURPOSE
'   Fills observed_excel and observed_kstats in
'   benchmark/excel_comparison/excel_comparison_grid.csv, so the library and
'   Excel's native functions can be compared against the same 50-digit
'   references.
'
' WHY THIS EXISTS
'   The README claims better deep-tail and large-shape accuracy than Excel's
'   native statistical functions. The claim was asserted rather than shown. This
'   macro produces the evidence, including on the points where Excel is expected
'   to do just as well - a comparison that only reported the wins would not be
'   evidence.
'
' BEHAVIOR
'   Each row names an Excel formula and the corresponding K_STATS_ call. Excel
'   values go through Application.Evaluate so the sheet formula is exercised as
'   a user would write it; a native function that raises or returns an Excel
'   error is recorded as the token ERROR, which is itself a measurement.
'
' ERROR POLICY
'   Reports failures through a message box; the file handle is always closed.
'
' UPDATED
'   2026-08-06
'==============================================================================
'
Private Const GRID_FILE As String = "excel_comparison_grid.csv"


Public Sub Export_ExcelComparison()
'
'==============================================================================
' Export_ExcelComparison
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the comparison grid, evaluates both columns and writes them back.
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Path                As String          'Resolved grid location
    Dim Buffer              As String          'Whole file
    Dim Lines()             As String          'Split on vbLf
    Dim Parts()             As String          'Current row
    Dim Idx                 As Long            'Row index
    Dim FileNo              As Integer         'File handle
    Dim DoneRows            As Long            'Rows evaluated
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo Err_Handler

    FileNo = 0
    DoneRows = 0

    Path = ResolveGridPath()
    If Len(Path) = 0 Then Exit Sub
'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    FileNo = FreeFile
    Open Path For Binary As #FileNo
    Buffer = Space$(LOF(FileNo))
    Get #FileNo, , Buffer
    Close #FileNo
    FileNo = 0

    Buffer = Replace$(Buffer, vbCrLf, vbLf)
    Lines = Split(Buffer, vbLf)
'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    For Idx = 1 To UBound(Lines)
        If Len(Trim$(Lines(Idx))) > 0 Then
            'PIPE-delimited by design: the labels and Excel formulas in this
            'grid all contain commas, so a comma split would shred the row and
            'write observations into the wrong columns.
            Parts = Split(Lines(Idx), "|")

            If UBound(Parts) >= 6 Then
                Parts(4) = EvaluateExpression(Parts(1))
                Parts(5) = EvaluateExpression(Parts(2))
                Lines(Idx) = Join(Parts, "|")
                DoneRows = DoneRows + 1
            End If
        End If
    Next Idx
'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    FileNo = FreeFile
    Open Path For Output As #FileNo
    Print #FileNo, Join(Lines, vbLf);
    Close #FileNo
    FileNo = 0

    MsgBox "Excel comparison complete." & vbCrLf & _
           "Rows evaluated: " & DoneRows & vbCrLf & vbCrLf & _
           "Now run:  python analyze_excel_comparison.py", _
           vbInformation, "Excel comparison"
    Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    On Error Resume Next
    If FileNo <> 0 Then Close #FileNo
    MsgBox "Excel comparison failed: " & Err.Description, vbExclamation
End Sub


Private Function EvaluateExpression(ByVal Expr As String) As String
'
'==============================================================================
' EvaluateExpression
'------------------------------------------------------------------------------
' PURPOSE
'   Evaluates one worksheet expression and returns the value at full precision,
'   or the token ERROR when it is not a finite number.
'
' NOTE
'   Application.Evaluate is used for BOTH columns so the two are measured the
'   same way: through the worksheet layer, exactly as a user would call them.
'==============================================================================
'
    Dim V                   As Variant         'Evaluation result

    On Error GoTo Err_Handler

    'A "-" marks a case where Excel offers no equivalent function. That is a
    'measurement in its own right - the most consequential difference of all -
    'so it is recorded explicitly rather than left blank or treated as an error.
    If Trim$(Expr) = "-" Then
        EvaluateExpression = "NONE"
        Exit Function
    End If

    V = Application.Evaluate(Trim$(Expr))

    If IsError(V) Then
        EvaluateExpression = "ERROR"
    ElseIf Not IsNumeric(V) Then
        EvaluateExpression = "ERROR"
    Else
        EvaluateExpression = FormatFullPrecision(CDbl(V))
    End If
    Exit Function

Err_Handler:
    EvaluateExpression = "ERROR"
End Function


Private Function ResolveGridPath() As String
'
'==============================================================================
' ResolveGridPath
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a usable LOCAL path to the grid, or an empty string if cancelled.
'
' WHY THIS EXISTS
'   ThisWorkbook.Path returns an http(s) URL on OneDrive or SharePoint, and Dir$
'   raises on such a path rather than returning empty, so the URL case must be
'   excluded BEFORE Dir$ is called.
'==============================================================================
'
    Dim BookPath            As String          'Workbook folder
    Dim Candidate           As String          'Path next to the workbook
    Dim Picked              As Variant         'File-dialog result

    BookPath = ThisWorkbook.Path
    If Len(BookPath) > 0 And LCase$(Left$(BookPath, 4)) <> "http" Then
        Candidate = BookPath & Application.PathSeparator & GRID_FILE
        If Len(Dir$(Candidate)) > 0 Then
            ResolveGridPath = Candidate
            Exit Function
        End If
    End If

    MsgBox "Could not locate " & GRID_FILE & " automatically " & _
           "(the workbook may be on OneDrive/SharePoint). Please select it.", _
           vbInformation, "Locate comparison grid"
    Picked = Application.GetOpenFilename( _
        FileFilter:="Comparison grid (*.csv),*.csv", _
        Title:="Select " & GRID_FILE)

    If VarType(Picked) = vbBoolean Then
        ResolveGridPath = vbNullString
    Else
        ResolveGridPath = CStr(Picked)
    End If
End Function


Private Function FormatFullPrecision(ByVal X As Double) As String
'
'==============================================================================
' FormatFullPrecision
'------------------------------------------------------------------------------
' PURPOSE
'   Writes X as a hi;lo pair so no accuracy is lost through 15-digit text.
'==============================================================================
'
    Dim Hi                  As Double          'Leading 15-digit part
    Dim Lo                  As Double          'Remainder

    If X = 0# Then
        FormatFullPrecision = "0"
        Exit Function
    End If

    Hi = Val(Fmt15(X))                         'Val is locale-independent
    Lo = X - Hi

    If Lo = 0# Then
        FormatFullPrecision = Fmt15(Hi)
    Else
        FormatFullPrecision = Fmt15(Hi) & ";" & Fmt15(Lo)
    End If
End Function


Private Function Fmt15(ByVal X As Double) As String
'
'==============================================================================
' Fmt15
'------------------------------------------------------------------------------
' PURPOSE
'   Formats X with 15 significant digits using an invariant decimal point.
'==============================================================================
'
    Dim Text                As String          'Formatted value
    Dim Sep                 As String          'Locale decimal separator

    Text = Format$(X, "0.00000000000000E+00")
    Sep = Mid$(CStr(1.5), 2, 1)
    If Sep <> "." Then Text = Replace$(Text, Sep, ".")

    Fmt15 = Text
End Function
