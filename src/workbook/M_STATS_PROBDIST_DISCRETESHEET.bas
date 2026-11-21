Attribute VB_Name = "M_STATS_PROBDIST_DISCRETESHEET"
'==============================================================================
' MODULE: M_STATS_PROBDIST_DISCRETESHEET
' PURPOSE: The Discrete sheet (Binomial and Poisson counts): its reset, its two
' charts and the line that names the Excel computing the sheet. The values are
' worksheet formulas that call the production API (mass, cumulative,
' survival, quantile, mean and standard deviation of the model chosen in D6)
' and spill the table of counts. This module computes nothing. Its name keeps
' it apart from the library's M_STATS_PROBDIST_DISCRETE.
'==============================================================================
Option Explicit

Private Const MODEL_CELL As String = "D6"
Private Const TRIALS_CELL As String = "D7"
Private Const PROBABILITY_CELL As String = "D8"
Private Const RATE_CELL As String = "D9"
Private Const EXPOSURE_CELL As String = "D10"
Private Const THRESHOLD_CELL As String = "D11"
Private Const LEVEL_CELL As String = "D12"
Private Const QUANTILE_CELL As String = "D28"
Private Const BUILD_CELL As String = "D34"
Private Const DEFAULT_MODEL As String = "Binomial"
Private Const DEFAULT_TRIALS As Double = 100
Private Const DEFAULT_PROBABILITY As Double = 0.02
Private Const DEFAULT_RATE As Double = 0.02
Private Const DEFAULT_EXPOSURE As Double = 100
Private Const DEFAULT_THRESHOLD As Double = 5
Private Const DEFAULT_LEVEL As Double = 0.99
Private Const MASS_CHART As String = "DiscreteMass"
Private Const CUMULATIVE_CHART As String = "DiscreteCumulative"

Public Function DiscreteSheet() As Worksheet
    Dim Candidate As Worksheet
    For Each Candidate In ThisWorkbook.Worksheets
        If Candidate.CodeName = "shDiscrete" Then
            Set DiscreteSheet = Candidate
            Exit Function
        End If
    Next Candidate
    Err.Raise vbObjectError + 290, "DiscreteSheet", "Discrete worksheet is missing."
End Function


Public Function DiscreteIsActive() As Boolean
    ' True when this workbook's active sheet is the Discrete sheet: the Ribbon
    ' commands then act on it instead of the Explorer.
    On Error Resume Next
    If ActiveWorkbook Is ThisWorkbook Then
        DiscreteIsActive = (ActiveSheet.CodeName = "shDiscrete")
    End If
End Function


Public Function DiscreteReady() As Boolean
    ' True when the sheet exists with its two charts, so the commands can run.
    On Error GoTo Failed
    DiscreteReady = (DiscreteSheet().ChartObjects.Count = 2)
    Exit Function
Failed:
    DiscreteReady = False
End Function


Public Sub ResetDiscreteInputs(Optional ByVal Confirm As Boolean = True)
    ' Restores the default example: 100 independent loans that default with
    ' probability 2% each, as a Binomial count; the Poisson inputs give the same
    ' mean, 0.02 defaults per loan-year over 100 loan-years.
    Dim Sheet As Worksheet
    Dim SavedEvents As Boolean
    Dim Failure As String
    Set Sheet = DiscreteSheet()
    If Confirm Then
        If MsgBox("Reset only the Discrete inputs to the default example (100 loans, default probability 2%)?", _
                  vbOKCancel + vbQuestion, "Discrete Distributions") <> vbOK Then Exit Sub
    End If
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Sheet.Range(MODEL_CELL).Value2 = DEFAULT_MODEL
    Sheet.Range(TRIALS_CELL).Value2 = DEFAULT_TRIALS
    Sheet.Range(PROBABILITY_CELL).Value2 = DEFAULT_PROBABILITY
    Sheet.Range(RATE_CELL).Value2 = DEFAULT_RATE
    Sheet.Range(EXPOSURE_CELL).Value2 = DEFAULT_EXPOSURE
    Sheet.Range(THRESHOLD_CELL).Value2 = DEFAULT_THRESHOLD
    Sheet.Range(LEVEL_CELL).Value2 = DEFAULT_LEVEL
    Application.EnableEvents = SavedEvents
    RefreshDiscrete
    Exit Sub
Failed:
    Failure = Err.Description
    Application.EnableEvents = SavedEvents
    Err.Raise vbObjectError + 291, "ResetDiscreteInputs", Failure
End Sub


Public Sub DiscreteInputsChanged(ByVal Target As Range)
    ' Called by Workbook_SheetChange. The results and the table follow the
    ' sheet's own formulas; a new input refreshes the two charts, whose titles
    ' state the model and its parameters.
    Dim Sheet As Worksheet
    On Error GoTo Failed
    Set Sheet = DiscreteSheet()
    If Not Target.Worksheet Is Sheet Then Exit Sub
    If Application.Intersect(Target, Sheet.Range(MODEL_CELL & ":" & LEVEL_CELL)) Is Nothing Then Exit Sub
    RefreshDiscrete
    Exit Sub
Failed:
    MsgBox "The Discrete sheet could not follow that change: " & Err.Description, vbExclamation, _
           "Discrete Distributions"
End Sub


Public Sub DiscreteOpened()
    ' Called by Workbook_Open: names the Excel that now computes the sheet,
    ' when it is not the one that saved it. The charts keep the types and
    ' titles they were saved with, so nothing else is set here: formatting
    ' them at every opening would leave the workbook marked as changed.
    Dim Sheet As Worksheet
    Dim WasSaved As Boolean
    On Error Resume Next
    WasSaved = ThisWorkbook.Saved
    Set Sheet = DiscreteSheet()
    If CStr(Sheet.Range(BUILD_CELL).Value2) <> DiscreteExcelBuild() Then
        Sheet.Range(BUILD_CELL).Value2 = DiscreteExcelBuild()
        ThisWorkbook.Saved = WasSaved
    End If
    On Error GoTo 0
End Sub


Public Function DiscreteExcelBuild() As String
    ' The Excel that computes the sheet: version, build, bitness and operating
    ' system, as the issue asks every observation to record.
    Dim Bitness As String
#If Win64 Then
    Bitness = "64-bit"
#Else
    Bitness = "32-bit"
#End If
    DiscreteExcelBuild = "Microsoft Excel " & Application.Version & ", build " & CStr(Application.Build) & _
                         ", " & Bitness & ", " & Application.OperatingSystem
End Function


Public Sub RefreshDiscrete()
    ' Recalculates the sheet, writes the Excel that computed it and sets the
    ' two charts: the mass as columns with the counts from the threshold up
    ' drawn in front, and the cumulative function as columns with the
    ' quantile's column drawn in front and the probability q as a line.
    Dim Sheet As Worksheet
    Dim SavedEvents As Boolean
    Dim WasProtected As Boolean
    Dim Failure As String
    Dim Model As String
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Set Sheet = DiscreteSheet()
    WasProtected = Sheet.ProtectContents
    Application.EnableEvents = False
    If WasProtected Then Sheet.Unprotect
    If CStr(Sheet.Range(BUILD_CELL).Value2) <> DiscreteExcelBuild() Then Sheet.Range(BUILD_CELL).Value2 = DiscreteExcelBuild()
    Sheet.Calculate
    Model = ModelText(Sheet)
    If Sheet.ChartObjects.Count = 2 Then
        FormatMassChart Sheet.ChartObjects(MASS_CHART).Chart, _
            "P(X = k), " & Model & "; at least k = " & CellText(Sheet, THRESHOLD_CELL) & " in front"
        FormatCumulativeChart Sheet.ChartObjects(CUMULATIVE_CHART).Chart, _
            "F(k) = P(X <= k), " & Model & "; Q(" & CellText(Sheet, LEVEL_CELL) & ") = " & CellText(Sheet, QUANTILE_CELL)
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
    Err.Raise vbObjectError + 292, "RefreshDiscrete", Failure
End Sub


Private Function CellText(ByVal Sheet As Worksheet, ByVal Address As String) As String
    ' The value of a cell for a chart title; an error value reads as "n/a".
    Dim Value As Variant
    Value = Sheet.Range(Address).Value2
    If IsError(Value) Then
        CellText = "n/a"
    Else
        CellText = CStr(Value)
    End If
End Function


Private Function ModelText(ByVal Sheet As Worksheet) As String
    ' The model and its parameters, as the chart titles name them.
    If CStr(Sheet.Range(MODEL_CELL).Value2) = "Binomial" Then
        ModelText = "Binomial n = " & CellText(Sheet, TRIALS_CELL) & ", p = " & CellText(Sheet, PROBABILITY_CELL)
    ElseIf CStr(Sheet.Range(MODEL_CELL).Value2) = "Poisson" Then
        ModelText = "Poisson rate " & CellText(Sheet, RATE_CELL) & " x exposure " & CellText(Sheet, EXPOSURE_CELL)
    Else
        ModelText = "no model"
    End If
End Function


Private Sub FormatMassChart(ByVal Chart As Chart, ByVal Title As String)
    ' Series: the mass of every count, the mass of the counts from k up.
    With Chart
        .ChartType = xlColumnClustered
        .ChartGroups(1).Overlap = 100
        .ChartGroups(1).GapWidth = 40
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(155, 175, 205)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
        FormatAxes Chart, Title, "Probability"
    End With
End Sub


Private Sub FormatCumulativeChart(ByVal Chart As Chart, ByVal Title As String)
    ' Series: F(k) of every count, F(Q) at the quantile, the probability q.
    With Chart
        ' The overlap is set while every series is a column, so ChartGroups(1) is
        ' the column group; then q becomes a line.
        .ChartType = xlColumnClustered
        .ChartGroups(1).Overlap = 100
        .ChartGroups(1).GapWidth = 40
        .SeriesCollection(3).ChartType = xlLine
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(155, 175, 205)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(112, 173, 71)
        .SeriesCollection(3).Format.Line.ForeColor.RGB = RGB(31, 56, 100)
        .SeriesCollection(3).Format.Line.DashStyle = msoLineDash
        .SeriesCollection(3).MarkerStyle = xlMarkerStyleNone
        FormatAxes Chart, Title, "Cumulative probability"
        .Axes(xlValue).MaximumScale = 1#
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
        .Axes(xlCategory).AxisTitle.Text = "Count k"
        .Axes(xlCategory).TickLabels.NumberFormat = "0"
        .Axes(xlCategory).TickLabelSpacingIsAuto = True
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = YTitle
        .Axes(xlValue).MinimumScale = 0#
        .Axes(xlValue).MaximumScaleIsAuto = True
        .Axes(xlValue).TickLabels.NumberFormat = "General"
    End With
End Sub
