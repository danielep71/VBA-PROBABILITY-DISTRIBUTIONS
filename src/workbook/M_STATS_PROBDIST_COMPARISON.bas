Attribute VB_Name = "M_STATS_PROBDIST_COMPARISON"
'==============================================================================
' MODULE: M_STATS_PROBDIST_COMPARISON
' PURPOSE: The Comparison sheet (Excel vs VBA vs Python): its reset, its chart
' and the line that names the Excel computing the sheet. The values are
' worksheet formulas: Excel's own functions, calls to the production API and
' the correct digits and grades against stored references. This module
' computes nothing.
'==============================================================================
Option Explicit

Private Const CASE_CELL As String = "D6"
Private Const REFERENCE_GRADE_CELL As String = "D7"
Private Const REPORT_GRADE_CELL As String = "D8"
Private Const PYTHON_CELL As String = "D9"
Private Const BUILD_CELL As String = "D34"
Private Const DEFAULT_CASE As String = "normal survival, z = 8"
Private Const DEFAULT_REFERENCE_GRADE As Double = 12
Private Const DEFAULT_REPORT_GRADE As Double = 6
Private Const FULL_DIGITS As Double = 16
Private Const CHART_NAME As String = "ComparisonDigits"

Public Function ComparisonSheet() As Worksheet
    Dim Candidate As Worksheet
    For Each Candidate In ThisWorkbook.Worksheets
        If Candidate.CodeName = "shComparison" Then
            Set ComparisonSheet = Candidate
            Exit Function
        End If
    Next Candidate
    Err.Raise vbObjectError + 270, "ComparisonSheet", "Comparison worksheet is missing."
End Function


Public Function ComparisonIsActive() As Boolean
    ' True when this workbook's active sheet is the Comparison sheet: the
    ' Ribbon commands then act on it instead of the Explorer.
    On Error Resume Next
    If ActiveWorkbook Is ThisWorkbook Then
        ComparisonIsActive = (ActiveSheet.CodeName = "shComparison")
    End If
End Function


Public Function ComparisonReady() As Boolean
    ' True when the sheet exists with its chart, so the commands can run.
    On Error GoTo Failed
    ComparisonReady = (ComparisonSheet().ChartObjects.Count = 1)
    Exit Function
Failed:
    ComparisonReady = False
End Function


Public Sub ResetComparisonInputs(Optional ByVal Confirm As Boolean = True)
    ' Restores the case shown, the study's two grade thresholds and an empty
    ' =PY() cell.
    Dim Sheet As Worksheet
    Dim SavedEvents As Boolean
    Dim Failure As String
    Set Sheet = ComparisonSheet()
    If Confirm Then
        If MsgBox("Reset only the Comparison inputs to the first tail case and the study's grades?", _
                  vbOKCancel + vbQuestion, "Excel vs VBA vs Python") <> vbOK Then Exit Sub
    End If
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Sheet.Range(CASE_CELL).Value2 = DEFAULT_CASE
    Sheet.Range(REFERENCE_GRADE_CELL).Value2 = DEFAULT_REFERENCE_GRADE
    Sheet.Range(REPORT_GRADE_CELL).Value2 = DEFAULT_REPORT_GRADE
    Sheet.Range(PYTHON_CELL).ClearContents
    Application.EnableEvents = SavedEvents
    RefreshComparison
    Exit Sub
Failed:
    Failure = Err.Description
    Application.EnableEvents = SavedEvents
    Err.Raise vbObjectError + 271, "ResetComparisonInputs", Failure
End Sub


Public Sub ComparisonInputsChanged(ByVal Target As Range)
    ' Called by Workbook_SheetChange. The values, digits and grades follow the
    ' sheet's own formulas; a change of a grade threshold also refreshes the
    ' chart, whose title states the two thresholds.
    Dim Sheet As Worksheet
    On Error GoTo Failed
    Set Sheet = ComparisonSheet()
    If Not Target.Worksheet Is Sheet Then Exit Sub
    If Application.Intersect(Target, Sheet.Range(REFERENCE_GRADE_CELL & ":" & REPORT_GRADE_CELL)) Is Nothing Then Exit Sub
    RefreshComparison
    Exit Sub
Failed:
    MsgBox "The Comparison sheet could not follow that change: " & Err.Description, vbExclamation, _
           "Excel vs VBA vs Python"
End Sub


Public Sub ComparisonOpened()
    ' Called by Workbook_Open: names the Excel that now computes the sheet,
    ' when it is not the one that saved it. The chart keeps the type and title
    ' it was saved with, so nothing else is set here: a chart formatted at
    ' every opening would leave the workbook marked as changed.
    Dim Sheet As Worksheet
    Dim WasSaved As Boolean
    On Error Resume Next
    WasSaved = ThisWorkbook.Saved
    Set Sheet = ComparisonSheet()
    If CStr(Sheet.Range(BUILD_CELL).Value2) <> ComparisonExcelBuild() Then
        Sheet.Range(BUILD_CELL).Value2 = ComparisonExcelBuild()
        ThisWorkbook.Saved = WasSaved
    End If
    On Error GoTo 0
End Sub


Public Function ComparisonExcelBuild() As String
    ' The Excel that computes the Excel column: version, build, bitness and
    ' operating system, as the issue asks every observation to record.
    Dim Bitness As String
#If Win64 Then
    Bitness = "64-bit"
#Else
    Bitness = "32-bit"
#End If
    ComparisonExcelBuild = "Microsoft Excel " & Application.Version & ", build " & CStr(Application.Build) & _
                           ", " & Bitness & ", " & Application.OperatingSystem
End Function


Public Sub RefreshComparison()
    ' Recalculates the sheet, writes the Excel that computed it and sets the
    ' bar chart of correct digits: one bar per implementation and case, the
    ' cases in the order of the table, the axis from 0 to 16 digits and the
    ' two grade thresholds in the title.
    Dim Sheet As Worksheet
    Dim Item As ChartObject
    Dim SavedEvents As Boolean
    Dim WasProtected As Boolean
    Dim Failure As String
    Dim Colors As Variant
    Dim Index As Long
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Set Sheet = ComparisonSheet()
    WasProtected = Sheet.ProtectContents
    Application.EnableEvents = False
    If WasProtected Then Sheet.Unprotect
    ' Written only when another Excel opens the file: a write at every opening
    ' would leave the workbook marked as changed.
    If CStr(Sheet.Range(BUILD_CELL).Value2) <> ComparisonExcelBuild() Then Sheet.Range(BUILD_CELL).Value2 = ComparisonExcelBuild()
    Sheet.Calculate
    If Sheet.ChartObjects.Count = 1 Then
        Set Item = Sheet.ChartObjects(CHART_NAME)
        If Item.Chart.SeriesCollection.Count = 3 Then
            Colors = Array(RGB(237, 125, 49), RGB(31, 56, 100), RGB(112, 173, 71))
            With Item.Chart
                .ChartType = xlBarClustered
                For Index = 1 To 3
                    .SeriesCollection(Index).Format.Fill.ForeColor.RGB = Colors(Index - 1)
                Next Index
                .ChartGroups(1).GapWidth = 60
                .ChartGroups(1).Overlap = 0
                .HasTitle = True
                .ChartTitle.Text = "Correct digits against the 50-digit reference (reference grade from " & _
                                   CStr(Sheet.Range(REFERENCE_GRADE_CELL).Value2) & ", report from " & _
                                   CStr(Sheet.Range(REPORT_GRADE_CELL).Value2) & ")"
                .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
                .HasLegend = True
                .Legend.Position = xlLegendPositionBottom
                .Axes(xlValue).MinimumScale = 0#
                .Axes(xlValue).MaximumScale = FULL_DIGITS
                .Axes(xlValue).MajorUnit = 2#
                .Axes(xlValue).TickLabels.NumberFormat = "0"
                .Axes(xlCategory).ReversePlotOrder = True
                .Axes(xlCategory).TickLabels.Font.Size = 7
                .Axes(xlCategory).TickLabelSpacing = 1
            End With
        End If
    End If
    If WasProtected Then Sheet.Protect
    Application.EnableEvents = SavedEvents
    Exit Sub
Failed:
    Failure = Err.Description
    On Error Resume Next
    If WasProtected Then Sheet.Protect
    Application.EnableEvents = SavedEvents
    On Error GoTo 0
    Err.Raise vbObjectError + 272, "RefreshComparison", Failure
End Sub
