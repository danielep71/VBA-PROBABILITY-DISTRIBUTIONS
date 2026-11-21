Attribute VB_Name = "M_STATS_PROBDIST_RISK"
'==============================================================================
' MODULE: M_STATS_PROBDIST_RISK
' PURPOSE: The Risk sheet (value at risk, expected shortfall, tails and default
' counts): its reset, its two charts and the line that names the Excel
' computing the sheet. The values are worksheet formulas that call the
' production API (Normal and Student-t quantiles, densities and survival
' functions, the Binomial quantile and survival) and spill the table of
' losses. This module computes nothing.
'==============================================================================
Option Explicit

Private Const VALUE_CELL As String = "D6"
Private Const MEAN_RETURN_CELL As String = "D7"
Private Const VOLATILITY_CELL As String = "D8"
Private Const HORIZON_CELL As String = "D9"
Private Const CONFIDENCE_CELL As String = "D10"
Private Const DF_CELL As String = "D11"
Private Const THRESHOLD_CELL As String = "D12"
Private Const LOANS_CELL As String = "D13"
Private Const DEFAULT_PROBABILITY_CELL As String = "D14"
Private Const LOSS_PER_DEFAULT_CELL As String = "D15"
Private Const NORMAL_VAR_CELL As String = "D23"
Private Const T_VAR_CELL As String = "D26"
Private Const BUILD_CELL As String = "D37"
Private Const DEFAULT_VALUE As Double = 1000000
Private Const DEFAULT_MEAN_RETURN As Double = 0.0003
Private Const DEFAULT_VOLATILITY As Double = 0.01
Private Const DEFAULT_HORIZON As Double = 10
Private Const DEFAULT_CONFIDENCE As Double = 0.99
Private Const DEFAULT_DF As Double = 4
Private Const DEFAULT_THRESHOLD As Double = 50000
Private Const DEFAULT_LOANS As Double = 100
Private Const DEFAULT_DEFAULT_PROBABILITY As Double = 0.02
Private Const DEFAULT_LOSS_PER_DEFAULT As Double = 10000
Private Const DENSITY_CHART As String = "RiskDensity"
Private Const TAIL_CHART As String = "RiskTail"

Public Function RiskSheet() As Worksheet
    Dim Candidate As Worksheet
    For Each Candidate In ThisWorkbook.Worksheets
        If Candidate.CodeName = "shRisk" Then
            Set RiskSheet = Candidate
            Exit Function
        End If
    Next Candidate
    Err.Raise vbObjectError + 300, "RiskSheet", "Risk worksheet is missing."
End Function


Public Function RiskIsActive() As Boolean
    ' True when this workbook's active sheet is the Risk sheet: the Ribbon
    ' commands then act on it instead of the Explorer.
    On Error Resume Next
    If ActiveWorkbook Is ThisWorkbook Then
        RiskIsActive = (ActiveSheet.CodeName = "shRisk")
    End If
End Function


Public Function RiskReady() As Boolean
    ' True when the sheet exists with its two charts, so the commands can run.
    On Error GoTo Failed
    RiskReady = (RiskSheet().ChartObjects.Count = 2)
    Exit Function
Failed:
    RiskReady = False
End Function


Public Sub ResetRiskInputs(Optional ByVal Confirm As Boolean = True)
    ' Restores the default example: a position of 1,000,000 with a daily
    ' volatility of 1% over 10 days at 99%, a Student t with 4 degrees of
    ' freedom, a loss threshold of 50,000, and 100 loans that default
    ' independently with probability 2% and a loss of 10,000 each.
    Dim Sheet As Worksheet
    Dim SavedEvents As Boolean
    Dim Failure As String
    Set Sheet = RiskSheet()
    If Confirm Then
        If MsgBox("Reset only the Risk inputs to the default example (1,000,000 over 10 days at 99%, 100 loans)?", _
                  vbOKCancel + vbQuestion, "Risk Applications") <> vbOK Then Exit Sub
    End If
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Sheet.Range(VALUE_CELL).Value2 = DEFAULT_VALUE
    Sheet.Range(MEAN_RETURN_CELL).Value2 = DEFAULT_MEAN_RETURN
    Sheet.Range(VOLATILITY_CELL).Value2 = DEFAULT_VOLATILITY
    Sheet.Range(HORIZON_CELL).Value2 = DEFAULT_HORIZON
    Sheet.Range(CONFIDENCE_CELL).Value2 = DEFAULT_CONFIDENCE
    Sheet.Range(DF_CELL).Value2 = DEFAULT_DF
    Sheet.Range(THRESHOLD_CELL).Value2 = DEFAULT_THRESHOLD
    Sheet.Range(LOANS_CELL).Value2 = DEFAULT_LOANS
    Sheet.Range(DEFAULT_PROBABILITY_CELL).Value2 = DEFAULT_DEFAULT_PROBABILITY
    Sheet.Range(LOSS_PER_DEFAULT_CELL).Value2 = DEFAULT_LOSS_PER_DEFAULT
    Application.EnableEvents = SavedEvents
    RefreshRisk
    Exit Sub
Failed:
    Failure = Err.Description
    Application.EnableEvents = SavedEvents
    Err.Raise vbObjectError + 301, "ResetRiskInputs", Failure
End Sub


Public Sub RiskInputsChanged(ByVal Target As Range)
    ' Called by Workbook_SheetChange. The results and the table follow the
    ' sheet's own formulas; a new input refreshes the two charts, whose titles
    ' state the confidence level and the degrees of freedom.
    Dim Sheet As Worksheet
    On Error GoTo Failed
    Set Sheet = RiskSheet()
    If Not Target.Worksheet Is Sheet Then Exit Sub
    If Application.Intersect(Target, Sheet.Range(VALUE_CELL & ":" & LOSS_PER_DEFAULT_CELL)) Is Nothing Then Exit Sub
    RefreshRisk
    Exit Sub
Failed:
    MsgBox "The Risk sheet could not follow that change: " & Err.Description, vbExclamation, _
           "Risk Applications"
End Sub


Public Sub RiskOpened()
    ' Called by Workbook_Open: names the Excel that now computes the sheet,
    ' when it is not the one that saved it. The charts keep the types and
    ' titles they were saved with, so nothing else is set here: formatting
    ' them at every opening would leave the workbook marked as changed.
    Dim Sheet As Worksheet
    Dim WasSaved As Boolean
    On Error Resume Next
    WasSaved = ThisWorkbook.Saved
    Set Sheet = RiskSheet()
    If CStr(Sheet.Range(BUILD_CELL).Value2) <> RiskExcelBuild() Then
        Sheet.Range(BUILD_CELL).Value2 = RiskExcelBuild()
        ThisWorkbook.Saved = WasSaved
    End If
    On Error GoTo 0
End Sub


Public Function RiskExcelBuild() As String
    ' The Excel that computes the sheet: version, build, bitness and operating
    ' system, as the issue asks every observation to record.
    Dim Bitness As String
#If Win64 Then
    Bitness = "64-bit"
#Else
    Bitness = "32-bit"
#End If
    RiskExcelBuild = "Microsoft Excel " & Application.Version & ", build " & CStr(Application.Build) & _
                     ", " & Bitness & ", " & Application.OperatingSystem
End Function


Public Sub RefreshRisk()
    ' Recalculates the sheet, writes the Excel that computed it and sets the
    ' two charts: the two loss densities with a vertical line at each VaR, and
    ' the two exceedance probabilities on a log scale with a line at
    ' 1 - confidence.
    Dim Sheet As Worksheet
    Dim SavedEvents As Boolean
    Dim WasProtected As Boolean
    Dim Failure As String
    Dim Levels As String
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Set Sheet = RiskSheet()
    WasProtected = Sheet.ProtectContents
    Application.EnableEvents = False
    If WasProtected Then Sheet.Unprotect
    If CStr(Sheet.Range(BUILD_CELL).Value2) <> RiskExcelBuild() Then Sheet.Range(BUILD_CELL).Value2 = RiskExcelBuild()
    Sheet.Calculate
    Levels = "confidence " & CellText(Sheet, CONFIDENCE_CELL) & ", nu = " & CellText(Sheet, DF_CELL)
    If Sheet.ChartObjects.Count = 2 Then
        FormatDensityChart Sheet.ChartObjects(DENSITY_CHART).Chart, _
            "Loss density, Normal and Student t at the same variance; " & Levels & _
            "; VaR " & CellText(Sheet, NORMAL_VAR_CELL) & " and " & CellText(Sheet, T_VAR_CELL)
        FormatTailChart Sheet.ChartObjects(TAIL_CHART).Chart, _
            "P(L > x), log scale; " & Levels
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
    Err.Raise vbObjectError + 302, "RefreshRisk", Failure
End Sub


Private Function CellText(ByVal Sheet As Worksheet, ByVal Address As String) As String
    ' The value of a cell for a chart title, rounded for reading; an error
    ' value reads as "n/a".
    Dim Value As Variant
    Value = Sheet.Range(Address).Value2
    If IsError(Value) Then
        CellText = "n/a"
    ElseIf IsNumeric(Value) Then
        CellText = CStr(Application.WorksheetFunction.Round(CDbl(Value), 4))
    Else
        CellText = CStr(Value)
    End If
End Function


Private Sub FormatDensityChart(ByVal Chart As Chart, ByVal Title As String)
    ' Series: the Normal density, the Student-t density, the Normal VaR line,
    ' the Student-t VaR line.
    With Chart
        .ChartType = xlXYScatterLinesNoMarkers
        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(68, 114, 196)
        .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(237, 125, 49)
        .SeriesCollection(3).Format.Line.ForeColor.RGB = RGB(68, 114, 196)
        .SeriesCollection(3).Format.Line.DashStyle = msoLineDash
        .SeriesCollection(4).Format.Line.ForeColor.RGB = RGB(237, 125, 49)
        .SeriesCollection(4).Format.Line.DashStyle = msoLineDash
        FormatAxes Chart, Title, "Density"
        .Axes(xlValue).ScaleType = xlScaleLinear
        .Axes(xlValue).MinimumScale = 0#
        .Axes(xlValue).MaximumScaleIsAuto = True
    End With
End Sub


Private Sub FormatTailChart(ByVal Chart As Chart, ByVal Title As String)
    ' Series: the Normal and Student-t exceedance probabilities, 1 - confidence.
    With Chart
        .ChartType = xlXYScatterLinesNoMarkers
        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(68, 114, 196)
        .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(237, 125, 49)
        .SeriesCollection(3).Format.Line.ForeColor.RGB = RGB(31, 56, 100)
        .SeriesCollection(3).Format.Line.DashStyle = msoLineDash
        FormatAxes Chart, Title, "P(L > x)"
        .Axes(xlValue).MinimumScaleIsAuto = True
        .Axes(xlValue).MaximumScale = 1#
        .Axes(xlValue).ScaleType = xlScaleLogarithmic
    End With
End Sub


Private Sub FormatAxes(ByVal Chart As Chart, ByVal Title As String, ByVal YTitle As String)
    With Chart
        .HasTitle = True
        .ChartTitle.Text = Title
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Loss x"
        .Axes(xlCategory).TickLabels.NumberFormat = "General"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = YTitle
        .Axes(xlValue).TickLabels.NumberFormat = "General"
    End With
End Sub
