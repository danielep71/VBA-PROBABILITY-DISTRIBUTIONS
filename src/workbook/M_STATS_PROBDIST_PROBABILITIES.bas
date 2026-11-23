Attribute VB_Name = "M_STATS_PROBDIST_PROBABILITIES"
'==============================================================================
' MODULE: M_STATS_PROBDIST_PROBABILITIES
' PURPOSE: The Probabilities sheet's inputs and chart. The numbers on the sheet
' are worksheet formulas that call the production API; this module loads a
' family's defaults, sets the chart types of its kind, and keeps the number
' formats and the parameter messages in step with the inputs. It computes
' nothing. The family table (limits, defaults, messages) is the Explorer's:
' it is read through the Explorer* names.
'==============================================================================
Option Explicit

Private Const FAMILY_CELL As String = "D6"
Private Const FIRST_PARAMETER_ROW As Long = 7
Private Const PARAMETER_COUNT As Long = 4
Private Const LOWER_CELL As String = "D11"
Private Const UPPER_CELL As String = "D12"
Private Const PROBABILITY_CELL As String = "D13"
Private Const DECIMALS_CELL As String = "D14"
Private Const DEFAULT_FAMILY As String = "Normal"
Private Const DEFAULT_PROBABILITY As Double = 0.95
Private Const DEFAULT_DECIMALS As Long = 6
Private Const MAX_DECIMALS As Long = 15
' The points a and b start at these shares of the family's plot window.
Private Const LOWER_SHARE As Double = 0.35
Private Const UPPER_SHARE As Double = 0.65
Private Const FORMATTED_CELLS As String = "D20:D27,G20:I1020"
Private Const X_CELLS As String = "F20:F1020"
Private Const CHART_NAME As String = "ProbabilitiesDensity"

Public Function ProbabilitiesSheet() As Worksheet
    Dim Candidate As Worksheet
    For Each Candidate In ThisWorkbook.Worksheets
        If Candidate.CodeName = "shProbability" Then
            Set ProbabilitiesSheet = Candidate
            Exit Function
        End If
    Next Candidate
    Err.Raise vbObjectError + 257, "ProbabilitiesSheet", "Probabilities worksheet is missing."
End Function


Public Function ProbabilitiesIsActive() As Boolean
    ' True when this workbook's active sheet is the Probabilities sheet: the
    ' Ribbon commands then act on it instead of the Explorer.
    On Error Resume Next
    If ActiveWorkbook Is ThisWorkbook Then
        ProbabilitiesIsActive = (ActiveSheet.CodeName = "shProbability")
    End If
End Function


Public Function ProbabilitiesReady() As Boolean
    ' True when the sheet exists with its chart, so the commands can run.
    On Error GoTo Failed
    ProbabilitiesReady = (ProbabilitiesSheet().ChartObjects.Count = 1)
    Exit Function
Failed:
    ProbabilitiesReady = False
End Function


Public Sub ResetProbabilitiesInputs(Optional ByVal Confirm As Boolean = True)
    Dim Sheet As Worksheet
    Dim SavedEvents As Boolean
    Dim Failure As String
    Set Sheet = ProbabilitiesSheet()
    If Confirm Then
        If MsgBox("Reset only the Probabilities inputs to Normal(0,1)?", vbOKCancel + vbQuestion, _
                  "Probabilities and Quantiles") <> vbOK Then Exit Sub
    End If
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Sheet.Range(FAMILY_CELL).Value2 = DEFAULT_FAMILY
    LoadFamilyDefaults Sheet
    Sheet.Range(PROBABILITY_CELL).Value2 = DEFAULT_PROBABILITY
    Sheet.Range(DECIMALS_CELL).Value2 = DEFAULT_DECIMALS
    ApplyProbabilitiesDecimals Sheet
    Application.EnableEvents = SavedEvents
    RefreshProbabilities
    Exit Sub
Failed:
    Failure = Err.Description
    Application.EnableEvents = SavedEvents
    Err.Raise vbObjectError + 258, "ResetProbabilitiesInputs", Failure
End Sub


Public Sub ProbabilitiesInputsChanged(ByVal Target As Range)
    ' Called by Workbook_SheetChange. Choosing a family in D6 loads its
    ' parameters and the points a and b; a paste over several cells that
    ' includes D6 keeps the pasted values and only brings the chart, the
    ' number formats and the parameter messages to the pasted family; the
    ' decimals cell sets the number formats; any other edit changes nothing
    ' here, because the results, the grid and the chart follow the sheet's own
    ' formulas.
    Dim Sheet As Worksheet
    Dim SavedEvents As Boolean
    Dim Failure As String
    ' Save the state first: the handler below restores it, whatever fails.
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Set Sheet = ProbabilitiesSheet()
    If Not Target.Worksheet Is Sheet Then Exit Sub
    If Not Application.Intersect(Target, Sheet.Range(DECIMALS_CELL)) Is Nothing Then ApplyProbabilitiesDecimals Sheet
    If Application.Intersect(Target, Sheet.Range(FAMILY_CELL)) Is Nothing Then Exit Sub
    Application.EnableEvents = False
    If Target.Cells.CountLarge = 1 Then
        LoadFamilyDefaults Sheet
    Else
        ApplyProbabilitiesDecimals Sheet
        ApplyProbabilitiesMessages Sheet
    End If
    Application.EnableEvents = SavedEvents
    RefreshProbabilities
    Exit Sub
Failed:
    Failure = Err.Description
    Application.EnableEvents = SavedEvents
    MsgBox "The Probabilities sheet could not follow that change: " & Failure, vbExclamation, _
           "Probabilities and Quantiles"
End Sub


Public Sub ProbabilitiesOpened()
    ' Called by Workbook_Open. A file saved while macros were off, or with
    ' another family's chart types, opens with a stale chart, number formats
    ' that do not follow the decimals cell and parameter messages of another
    ' family: set them again, and leave the Saved flag as it was.
    Dim WasSaved As Boolean
    On Error Resume Next
    WasSaved = ThisWorkbook.Saved
    RefreshProbabilities
    ApplyProbabilitiesDecimals ProbabilitiesSheet()
    ApplyProbabilitiesMessages ProbabilitiesSheet()
    ThisWorkbook.Saved = WasSaved
    On Error GoTo 0
End Sub


Public Sub RefreshProbabilities()
    ' Recalculates the sheet and sets the chart for the family's kind: a
    ' discrete family plots its mass as columns, with the two regions drawn in
    ' front of them, a continuous one plots the density as a line and the two
    ' regions as filled areas. The series follow the spilled grid through
    ' workbook names, so no range is set here.
    Dim Sheet As Worksheet
    Dim Item As ChartObject
    Dim Discrete As Boolean
    Dim SavedEvents As Boolean
    Dim WasProtected As Boolean
    Dim Failure As String
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Set Sheet = ProbabilitiesSheet()
    If Sheet.ChartObjects.Count <> 1 Then Exit Sub
    Set Item = Sheet.ChartObjects(CHART_NAME)
    If Item.Chart.SeriesCollection.Count <> 3 Then Exit Sub
    WasProtected = Sheet.ProtectContents
    Application.EnableEvents = False
    If WasProtected Then Sheet.Unprotect
    Sheet.Calculate
    Discrete = IsDiscrete(Sheet)
    With Item.Chart
        If Discrete Then
            .SeriesCollection(1).ChartType = xlColumnClustered
            .SeriesCollection(2).ChartType = xlColumnClustered
            .SeriesCollection(3).ChartType = xlColumnClustered
            .ChartGroups(1).Overlap = 100
            .ChartGroups(1).GapWidth = 40
            .ChartTitle.Text = "Probability mass P(X = k), with P(X <= a) and P(a < X <= b) shaded"
        Else
            .SeriesCollection(1).ChartType = xlLine
            .SeriesCollection(2).ChartType = xlArea
            .SeriesCollection(3).ChartType = xlArea
            .ChartTitle.Text = "Density f(x), with P(X <= a) and P(a < X <= b) shaded"
        End If
        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(31, 56, 100)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(112, 173, 71)
        If Discrete Then .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(155, 175, 205)
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Axes(xlValue).MinimumScale = 0#
        .Axes(xlValue).MaximumScaleIsAuto = True
        .Axes(xlValue).TickLabels.NumberFormat = "General"
        .Axes(xlCategory).TickLabels.NumberFormat = "General"
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        If Discrete Then
            .Axes(xlCategory).TickLabelSpacingIsAuto = True
        Else
            .Axes(xlCategory).TickLabelSpacing = 25
            .Axes(xlCategory).TickMarkSpacing = 25
        End If
    End With
    If WasProtected Then Sheet.Protect
    Application.EnableEvents = SavedEvents
    Exit Sub
Failed:
    Failure = Err.Description
    On Error Resume Next
    If WasProtected Then Sheet.Protect
    Application.EnableEvents = SavedEvents
    On Error GoTo 0
    Err.Raise vbObjectError + 259, "RefreshProbabilities", Failure
End Sub


Public Sub ApplyProbabilitiesDecimals(ByVal Sheet As Worksheet)
    ' Formats the results and the grid with the number of decimals in D14. A
    ' value that is not a whole number from 0 to 15 leaves the formats as they
    ' are. The x column of a discrete family is shown in whole numbers.
    Dim Places As Variant
    Dim WasProtected As Boolean
    Dim Failure As String
    Dim Pattern As String
    Places = Sheet.Range(DECIMALS_CELL).Value2
    If IsError(Places) Then Exit Sub
    If Not IsNumeric(Places) Or IsEmpty(Places) Then Exit Sub
    If CDbl(Places) <> Int(CDbl(Places)) Or CDbl(Places) < 0 Or CDbl(Places) > MAX_DECIMALS Then Exit Sub
    WasProtected = Sheet.ProtectContents
    If WasProtected Then Sheet.Unprotect
    On Error GoTo Failed
    If CLng(Places) = 0 Then Pattern = "0" Else Pattern = "0." & String$(CLng(Places), "0")
    Sheet.Range(FORMATTED_CELLS).NumberFormat = Pattern
    If IsDiscrete(Sheet) Then
        Sheet.Range(X_CELLS).NumberFormat = "0"
    Else
        Sheet.Range(X_CELLS).NumberFormat = Pattern
    End If
    If WasProtected Then Sheet.Protect
    Exit Sub
Failed:
    ' Protect the sheet again as it was found, then pass the error to the caller.
    Failure = Err.Description
    On Error Resume Next
    If WasProtected Then Sheet.Protect
    On Error GoTo 0
    Err.Raise vbObjectError + 262, "ApplyProbabilitiesDecimals", Failure
End Sub


Public Sub ApplyProbabilitiesMessages(ByVal Sheet As Worksheet)
    ' Gives the parameter cells D7:D10 the input and error messages of the
    ' selected family's parameters: name, meaning, unit and limits, as the
    ' Explorer's family table holds them. A name that is not a family leaves
    ' them as they are.
    Dim Row As Long
    Dim Slot As Long
    Dim WasProtected As Boolean
    Dim Failure As String
    Row = FamilyRow(Sheet)
    If Row = 0 Then Exit Sub
    WasProtected = Sheet.ProtectContents
    If WasProtected Then Sheet.Unprotect
    On Error GoTo Failed
    For Slot = 1 To PARAMETER_COUNT
        With Sheet.Cells(FIRST_PARAMETER_ROW + Slot - 1, 4).Validation
            .InputTitle = Left$(CStr(TableRange("ExplorerLabels").Cells(Row, Slot).Value2), 32)
            .InputMessage = Left$(CStr(TableRange("ExplorerPrompts").Cells(Row, Slot).Value2), 255)
            .ErrorTitle = Left$(CStr(TableRange("ExplorerErrorTitles").Cells(Row, Slot).Value2), 32)
            .ErrorMessage = Left$(CStr(TableRange("ExplorerErrors").Cells(Row, Slot).Value2), 225)
            .ShowInput = True
            .ShowError = True
        End With
    Next Slot
    If WasProtected Then Sheet.Protect
    Exit Sub
Failed:
    ' Protect the sheet again as it was found, then pass the error to the caller.
    Failure = Err.Description
    On Error Resume Next
    If WasProtected Then Sheet.Protect
    On Error GoTo 0
    Err.Raise vbObjectError + 263, "ApplyProbabilitiesMessages", Failure
End Sub


Private Function TableRange(ByVal TableName As String) As Range
    ' A range of the Explorer's family table, by its workbook name: the names
    ' refer to another sheet, which a sheet's own Range(Name) does not follow.
    Set TableRange = ThisWorkbook.Names.Item(TableName).RefersToRange
End Function


Private Function FamilyRow(ByVal Sheet As Worksheet) As Long
    ' Row of the selected family in the table, or 0 when D6 is not a family.
    Dim Position As Variant
    Position = Application.Match(Sheet.Range(FAMILY_CELL).Value2, TableRange("ExplorerFamily"), 0)
    If IsError(Position) Then FamilyRow = 0 Else FamilyRow = CLng(Position)
End Function


Private Function IsDiscrete(ByVal Sheet As Worksheet) As Boolean
    Dim Row As Long
    Row = FamilyRow(Sheet)
    If Row > 0 Then IsDiscrete = (CStr(TableRange("ExplorerKind").Cells(Row, 1).Value2) = "Discrete")
End Function


Private Sub LoadFamilyDefaults(ByVal Sheet As Worksheet)
    ' Writes the family's parameters and the points a and b. The caller has
    ' switched events off, so these writes do not re-enter the handler.
    Dim Row As Long
    Dim Slot As Long
    Row = FamilyRow(Sheet)
    If Row = 0 Then Exit Sub
    For Slot = 1 To PARAMETER_COUNT
        Sheet.Cells(FIRST_PARAMETER_ROW + Slot - 1, 4).Value2 = _
            TableRange("ExplorerDefaults").Cells(Row, Slot).Value2
    Next Slot
    SetDefaultPoints Sheet
    ApplyProbabilitiesDecimals Sheet
    ApplyProbabilitiesMessages Sheet
End Sub


Private Sub SetDefaultPoints(ByVal Sheet As Worksheet)
    ' Puts a and b at fixed shares of the plot window the family suggests for
    ' its current parameters, rounded down for a discrete family, so that both
    ' fall inside the chart. A window that cannot be computed leaves them as
    ' they are.
    Dim Lower As Variant
    Dim Upper As Variant
    Dim A As Double
    Dim B As Double
    Sheet.Calculate
    Lower = Sheet.Evaluate("ProbabilitiesPlotLower")
    Upper = Sheet.Evaluate("ProbabilitiesPlotUpper")
    If IsError(Lower) Or IsError(Upper) Then Exit Sub
    If Not IsNumeric(Lower) Or Not IsNumeric(Upper) Then Exit Sub
    A = CDbl(Lower) + LOWER_SHARE * (CDbl(Upper) - CDbl(Lower))
    B = CDbl(Lower) + UPPER_SHARE * (CDbl(Upper) - CDbl(Lower))
    If IsDiscrete(Sheet) Then
        A = Int(A)
        B = Int(B)
    End If
    Sheet.Range(LOWER_CELL).Value2 = A
    Sheet.Range(UPPER_CELL).Value2 = B
End Sub
