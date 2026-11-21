Attribute VB_Name = "M_STATS_PROBDIST_TAILS"
'==============================================================================
' MODULE: M_STATS_PROBDIST_TAILS
' PURPOSE: The Tails sheet (Extreme Tails and Numerical Limits): its reset, its
' two charts and the line that names the Excel computing the sheet. The values
' are worksheet formulas: calls to the production API, the alternative path
' (1 - CDF, or LN of the mass), the correct digits and the statuses against
' stored mpmath references. This module computes nothing.
'==============================================================================
Option Explicit

Private Const SCENARIO_CELL As String = "D6"
Private Const REFERENCE_GRADE_CELL As String = "D7"
Private Const REPORT_GRADE_CELL As String = "D8"
Private Const BUILD_CELL As String = "D34"
Private Const DEFAULT_SCENARIO As String = "Standard Normal"
Private Const DEFAULT_REFERENCE_GRADE As Double = 12
Private Const DEFAULT_REPORT_GRADE As Double = 6
Private Const FULL_DIGITS As Double = 16
Private Const DIGITS_CHART As String = "TailsDigits"
Private Const VALUES_CHART As String = "TailsValues"

Public Function TailsSheet() As Worksheet
    Dim Candidate As Worksheet
    For Each Candidate In ThisWorkbook.Worksheets
        If Candidate.CodeName = "shTails" Then
            Set TailsSheet = Candidate
            Exit Function
        End If
    Next Candidate
    Err.Raise vbObjectError + 280, "TailsSheet", "Tails worksheet is missing."
End Function


Public Function TailsIsActive() As Boolean
    ' True when this workbook's active sheet is the Tails sheet: the Ribbon
    ' commands then act on it instead of the Explorer.
    On Error Resume Next
    If ActiveWorkbook Is ThisWorkbook Then
        TailsIsActive = (ActiveSheet.CodeName = "shTails")
    End If
End Function


Public Function TailsReady() As Boolean
    ' True when the sheet exists with its two charts, so the commands can run.
    On Error GoTo Failed
    TailsReady = (TailsSheet().ChartObjects.Count = 2)
    Exit Function
Failed:
    TailsReady = False
End Function


Public Sub ResetTailsInputs(Optional ByVal Confirm As Boolean = True)
    ' Restores the scenario shown and the study's two grade thresholds.
    Dim Sheet As Worksheet
    Dim SavedEvents As Boolean
    Dim Failure As String
    Set Sheet = TailsSheet()
    If Confirm Then
        If MsgBox("Reset only the Tails inputs to the Standard Normal scenario and the study's grades?", _
                  vbOKCancel + vbQuestion, "Extreme Tails and Numerical Limits") <> vbOK Then Exit Sub
    End If
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Sheet.Range(SCENARIO_CELL).Value2 = DEFAULT_SCENARIO
    Sheet.Range(REFERENCE_GRADE_CELL).Value2 = DEFAULT_REFERENCE_GRADE
    Sheet.Range(REPORT_GRADE_CELL).Value2 = DEFAULT_REPORT_GRADE
    Application.EnableEvents = SavedEvents
    RefreshTails
    Exit Sub
Failed:
    Failure = Err.Description
    Application.EnableEvents = SavedEvents
    Err.Raise vbObjectError + 281, "ResetTailsInputs", Failure
End Sub


Public Sub TailsInputsChanged(ByVal Target As Range)
    ' Called by Workbook_SheetChange. The table, the values and the statuses
    ' follow the sheet's own formulas; a new scenario or threshold refreshes
    ' the two charts, whose titles state them.
    Dim Sheet As Worksheet
    On Error GoTo Failed
    Set Sheet = TailsSheet()
    If Not Target.Worksheet Is Sheet Then Exit Sub
    If Application.Intersect(Target, Sheet.Range(SCENARIO_CELL & ":" & REPORT_GRADE_CELL)) Is Nothing Then Exit Sub
    RefreshTails
    Exit Sub
Failed:
    MsgBox "The Tails sheet could not follow that change: " & Err.Description, vbExclamation, _
           "Extreme Tails and Numerical Limits"
End Sub


Public Sub TailsOpened()
    ' Called by Workbook_Open: names the Excel that now computes the sheet,
    ' when it is not the one that saved it. The charts keep the types and
    ' titles they were saved with, so nothing else is set here: formatting
    ' them at every opening would leave the workbook marked as changed.
    Dim Sheet As Worksheet
    Dim WasSaved As Boolean
    On Error Resume Next
    WasSaved = ThisWorkbook.Saved
    Set Sheet = TailsSheet()
    If CStr(Sheet.Range(BUILD_CELL).Value2) <> TailsExcelBuild() Then
        Sheet.Range(BUILD_CELL).Value2 = TailsExcelBuild()
        ThisWorkbook.Saved = WasSaved
    End If
    On Error GoTo 0
End Sub


Public Function TailsExcelBuild() As String
    ' The Excel that computes the sheet: version, build, bitness and operating
    ' system, as the issue asks every observation to record.
    Dim Bitness As String
#If Win64 Then
    Bitness = "64-bit"
#Else
    Bitness = "32-bit"
#End If
    TailsExcelBuild = "Microsoft Excel " & Application.Version & ", build " & CStr(Application.Build) & _
                      ", " & Bitness & ", " & Application.OperatingSystem
End Function


Public Sub RefreshTails()
    ' Recalculates the sheet, writes the Excel that computed it and sets the
    ' two charts: the correct digits of the two paths against the depth of the
    ' tail (0 to 16 digits), and log10 of the values against it, each with the
    ' limit of the Double range drawn. Zero and unrepresentable values are not
    ' plotted: the table states them.
    Dim Sheet As Worksheet
    Dim SavedEvents As Boolean
    Dim WasProtected As Boolean
    Dim Failure As String
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Set Sheet = TailsSheet()
    WasProtected = Sheet.ProtectContents
    Application.EnableEvents = False
    If WasProtected Then Sheet.Unprotect
    If CStr(Sheet.Range(BUILD_CELL).Value2) <> TailsExcelBuild() Then Sheet.Range(BUILD_CELL).Value2 = TailsExcelBuild()
    Sheet.Calculate
    If Sheet.ChartObjects.Count = 2 Then
        FormatChart Sheet.ChartObjects(DIGITS_CHART).Chart, _
            "Correct digits, " & CStr(Sheet.Range(SCENARIO_CELL).Value2) & " (reference grade from " & _
            CStr(Sheet.Range(REFERENCE_GRADE_CELL).Value2) & ", report from " & CStr(Sheet.Range(REPORT_GRADE_CELL).Value2) & ")", _
            "Depth of the tail: -log10 of the probability", "Correct digits", True
        FormatChart Sheet.ChartObjects(VALUES_CHART).Chart, _
            "log10 of the values, " & CStr(Sheet.Range(SCENARIO_CELL).Value2), _
            "Depth of the tail: -log10 of the probability", "log10 of the value (of the mass, for LogPMF)", False
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
    Err.Raise vbObjectError + 282, "RefreshTails", Failure
End Sub


Private Sub FormatChart(ByVal Chart As Chart, ByVal Title As String, ByVal XTitle As String, _
                        ByVal YTitle As String, ByVal Digits As Boolean)
    Dim Colors As Variant
    Dim Index As Long
    Colors = Array(RGB(112, 173, 71), RGB(31, 56, 100), RGB(237, 125, 49), RGB(128, 128, 128))
    With Chart
        .ChartType = xlXYScatterLines
        .HasTitle = True
        .ChartTitle.Text = Title
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        If Digits Then
            ' Series: the direct path, the alternative path, the limit of the Double range.
            .SeriesCollection(1).Format.Line.ForeColor.RGB = Colors(1)
            .SeriesCollection(2).Format.Line.ForeColor.RGB = Colors(2)
            .SeriesCollection(3).Format.Line.ForeColor.RGB = Colors(3)
        Else
            ' Series: the reference, the direct path, the alternative path, the limit.
            .SeriesCollection(1).Format.Line.ForeColor.RGB = Colors(0)
            .SeriesCollection(2).Format.Line.ForeColor.RGB = Colors(1)
            .SeriesCollection(3).Format.Line.ForeColor.RGB = Colors(2)
            .SeriesCollection(4).Format.Line.ForeColor.RGB = Colors(3)
        End If
        For Index = 1 To .SeriesCollection.Count
            .SeriesCollection(Index).MarkerStyle = xlMarkerStyleNone
        Next Index
        If Digits Then
            .SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
            .SeriesCollection(2).MarkerStyle = xlMarkerStyleDiamond
            .SeriesCollection(1).MarkerSize = 5
            .SeriesCollection(2).MarkerSize = 5
            .SeriesCollection(3).Format.Line.DashStyle = msoLineDash
        Else
            .SeriesCollection(2).MarkerStyle = xlMarkerStyleCircle
            .SeriesCollection(3).MarkerStyle = xlMarkerStyleDiamond
            .SeriesCollection(2).MarkerSize = 5
            .SeriesCollection(3).MarkerSize = 5
            .SeriesCollection(4).Format.Line.DashStyle = msoLineDash
        End If
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = XTitle
        .Axes(xlCategory).MinimumScale = 0#
        .Axes(xlCategory).TickLabels.NumberFormat = "0"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = YTitle
        .Axes(xlValue).TickLabels.NumberFormat = "0"
        If Digits Then
            .Axes(xlValue).MinimumScale = 0#
            .Axes(xlValue).MaximumScale = FULL_DIGITS
            .Axes(xlValue).MajorUnit = 2#
        Else
            .Axes(xlValue).MinimumScaleIsAuto = True
            .Axes(xlValue).MaximumScale = 0#
        End If
    End With
End Sub
