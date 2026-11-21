Attribute VB_Name = "M_STATS_PROBDIST_EXPLORER"
'==============================================================================
' MODULE: M_STATS_PROBDIST_EXPLORER
' PURPOSE: The Distribution Explorer's inputs and charts. The numbers on the
' sheet are worksheet formulas that call the production API and spill as
' dynamic arrays; this module loads a family's defaults, refits the plot
' window, and sets the chart types. It computes nothing.
' The family table (defaults, limits, plot windows) is on the sheet, in the
' hidden columns from MA, and is read through the Explorer* names.
'==============================================================================
Option Explicit

Private Const FAMILY_CELL As String = "D6"
Private Const FIRST_PARAMETER_ROW As Long = 7
Private Const PARAMETER_COUNT As Long = 4
Private Const DEFAULT_FAMILY As String = "Normal"
Private Const DEFAULT_POINTS As Long = 201
Private Const DEFAULT_PROBABILITY As Double = 0.95
Private Const DEFAULT_DECIMALS As Long = 6
' Inputs D6:D12 (distribution, parameters, x, p); plot window D13:D15;
' decimals D16, applied by ApplyDecimals to the results and the grid.
Private Const LOWER_CELL As String = "D13"
Private Const UPPER_CELL As String = "D14"
Private Const POINTS_CELL As String = "D15"
Private Const DECIMALS_CELL As String = "D16"
Private Const MAX_DECIMALS As Long = 15
Private Const FORMATTED_CELLS As String = "D21:D25,G20:I1020"
Private Const X_CELLS As String = "F20:F1020"

Public Function ExplorerSheet() As Worksheet
    Dim Candidate As Worksheet
    For Each Candidate In ThisWorkbook.Worksheets
        If Candidate.CodeName = "shExplorer" Then
            Set ExplorerSheet = Candidate
            Exit Function
        End If
    Next Candidate
    Err.Raise vbObjectError + 254, "ExplorerSheet", "Explorer worksheet is missing."
End Function


Public Sub ExplorerAction(ByVal Control As IRibbonControl)
    On Error GoTo Failed
    Select Case Control.Id
        Case "pdResetInputs"
            If M_STATS_PROBDIST_PROBABILITIES.ProbabilitiesIsActive() Then
                M_STATS_PROBDIST_PROBABILITIES.ResetProbabilitiesInputs
            ElseIf M_STATS_PROBDIST_PARAMETERS.ParameterEffectsIsActive() Then
                M_STATS_PROBDIST_PARAMETERS.ResetParameterEffectsInputs
            ElseIf M_STATS_PROBDIST_COMPARISON.ComparisonIsActive() Then
                M_STATS_PROBDIST_COMPARISON.ResetComparisonInputs
            ElseIf M_STATS_PROBDIST_TAILS.TailsIsActive() Then
                M_STATS_PROBDIST_TAILS.ResetTailsInputs
            Else
                ResetExplorerInputs
            End If
        Case "pdRecalculate"
            RecalculateWorkbook
            RefreshExplorer
            If M_STATS_PROBDIST_PROBABILITIES.ProbabilitiesReady() Then M_STATS_PROBDIST_PROBABILITIES.RefreshProbabilities
            If M_STATS_PROBDIST_PARAMETERS.ParameterEffectsReady() Then M_STATS_PROBDIST_PARAMETERS.RefreshParameterEffects
            If M_STATS_PROBDIST_COMPARISON.ComparisonReady() Then M_STATS_PROBDIST_COMPARISON.RefreshComparison
            If M_STATS_PROBDIST_TAILS.TailsReady() Then M_STATS_PROBDIST_TAILS.RefreshTails
    End Select
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Distribution Explorer"
End Sub


Public Sub ExplorerEnabled(ByVal Control As IRibbonControl, ByRef ReturnedVal As Variant)
    Dim Target As Worksheet
    On Error GoTo Failed
    If M_STATS_PROBDIST_PROBABILITIES.ProbabilitiesIsActive() Then
        ReturnedVal = M_STATS_PROBDIST_PROBABILITIES.ProbabilitiesReady()
        Exit Sub
    End If
    If M_STATS_PROBDIST_PARAMETERS.ParameterEffectsIsActive() Then
        ReturnedVal = M_STATS_PROBDIST_PARAMETERS.ParameterEffectsReady()
        Exit Sub
    End If
    If M_STATS_PROBDIST_COMPARISON.ComparisonIsActive() Then
        ReturnedVal = M_STATS_PROBDIST_COMPARISON.ComparisonReady()
        Exit Sub
    End If
    If M_STATS_PROBDIST_TAILS.TailsIsActive() Then
        ReturnedVal = M_STATS_PROBDIST_TAILS.TailsReady()
        Exit Sub
    End If
    Set Target = ExplorerSheet()
    ReturnedVal = (Target.ChartObjects.Count = 2)
    Exit Sub
Failed:
    ReturnedVal = False
End Sub


Public Sub ResetExplorerInputs(Optional ByVal Confirm As Boolean = True)
    Dim Sheet As Worksheet
    Dim SavedEvents As Boolean
    Dim Failure As String
    Set Sheet = ExplorerSheet()
    If Confirm Then
        If MsgBox("Reset only Explorer inputs to Normal(0,1)?", vbOKCancel + vbQuestion, _
                  "Distribution Explorer") <> vbOK Then Exit Sub
    End If
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Sheet.Range(FAMILY_CELL).Value2 = DEFAULT_FAMILY
    LoadFamilyDefaults Sheet
    Sheet.Range("D12").Value2 = DEFAULT_PROBABILITY
    Sheet.Range(DECIMALS_CELL).Value2 = DEFAULT_DECIMALS
    ApplyDecimals Sheet
    Application.EnableEvents = SavedEvents
    RefreshExplorer
    Exit Sub
Failed:
    Failure = Err.Description
    Application.EnableEvents = SavedEvents
    Err.Raise vbObjectError + 255, "ResetExplorerInputs", Failure
End Sub


Public Sub ExplorerInputsChanged(ByVal Target As Range)
    ' Called by Workbook_SheetChange. Choosing a family loads its parameters,
    ' probe point and plot window; editing a parameter refits the plot window;
    ' any other edit changes nothing here, because the grid and the charts
    ' follow the sheet's own formulas. A failure is reported once, in words,
    ' instead of leaving the user in the VBA debugger.
    Dim Sheet As Worksheet
    Dim SavedEvents As Boolean
    Dim Failure As String
    On Error GoTo Failed
    Set Sheet = ExplorerSheet()
    If Not Target.Worksheet Is Sheet Then Exit Sub
    If Not Application.Intersect(Target, Sheet.Range(DECIMALS_CELL)) Is Nothing Then ApplyDecimals Sheet
    If Application.Intersect(Target, Sheet.Range("D6:D10")) Is Nothing Then Exit Sub
    SavedEvents = Application.EnableEvents
    Application.EnableEvents = False
    If Target.Cells.CountLarge = 1 And Target.Address(False, False) = FAMILY_CELL Then
        LoadFamilyDefaults Sheet
    Else
        If Not Application.Intersect(Target, Sheet.Range("D7:D10")) Is Nothing Then RefitPlotWindow Sheet
    End If
    Application.EnableEvents = SavedEvents
    RefreshExplorer
    Exit Sub
Failed:
    Failure = Err.Description
    Application.EnableEvents = SavedEvents
    MsgBox "The Explorer could not follow that change: " & Failure, vbExclamation, _
           "Distribution Explorer"
End Sub


Public Sub ExplorerOpened()
    ' Called by Workbook_Open. A file saved while macros were off, or with
    ' another family's chart types, opens with stale charts, number formats
    ' that do not follow the decimals cell and parameter messages of another
    ' family: recalculate, set them again, and leave the Saved flag as it was.
    Dim WasSaved As Boolean
    On Error Resume Next
    WasSaved = ThisWorkbook.Saved
    RefreshExplorer
    ApplyDecimals ExplorerSheet()
    ApplyParameterMessages ExplorerSheet()
    ThisWorkbook.Saved = WasSaved
    On Error GoTo 0
End Sub


Public Sub RefreshExplorer()
    ' Recalculates the sheet and sets the two charts for the family's kind:
    ' a discrete family plots its mass as columns and its cumulative as
    ' points, a continuous one plots both as lines. The series follow the
    ' spilled grid through workbook names, so no range is set here.
    Dim Sheet As Worksheet
    Dim Item As ChartObject
    Dim Discrete As Boolean
    Dim SavedEvents As Boolean
    Dim WasProtected As Boolean
    Dim Failure As String
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Set Sheet = ExplorerSheet()
    If Sheet.ChartObjects.Count <> 2 Then Exit Sub
    WasProtected = Sheet.ProtectContents
    Application.EnableEvents = False
    If WasProtected Then Sheet.Unprotect
    Sheet.Calculate
    Discrete = IsDiscrete(Sheet)
    For Each Item In Sheet.ChartObjects
        If Discrete Then
            If Item.Name = "ExplorerDensity" Then
                Item.Chart.ChartType = xlColumnClustered
            Else
                Item.Chart.ChartType = xlXYScatter
            End If
        Else
            Item.Chart.ChartType = xlXYScatterLinesNoMarkers
        End If
        If Item.Name = "ExplorerDensity" Then
            Item.Chart.ChartTitle.Text = IIf(Discrete, "Probability mass P(X = k)", "Density f(x)")
            Item.Chart.Axes(xlValue).MinimumScale = 0#
            Item.Chart.Axes(xlValue).MaximumScaleIsAuto = True
        Else
            Item.Chart.ChartTitle.Text = "Cumulative probability P(X <= x)"
            Item.Chart.Axes(xlValue).MinimumScale = 0#
            Item.Chart.Axes(xlValue).MaximumScale = 1#
        End If
        ' General shows 5 as 5 and 0.25 as 0,25 or 0.25 by locale; a format
        ' with optional decimals leaves a bare separator after whole numbers.
        Item.Chart.Axes(xlCategory).TickLabels.NumberFormat = "General"
        Item.Chart.Axes(xlValue).TickLabels.NumberFormat = "General"
        Item.Chart.Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
    Next Item
    If WasProtected Then Sheet.Protect
    Application.EnableEvents = SavedEvents
    Exit Sub
Failed:
    Failure = Err.Description
    On Error Resume Next
    If WasProtected Then Sheet.Protect
    Application.EnableEvents = SavedEvents
    On Error GoTo 0
    Err.Raise vbObjectError + 256, "RefreshExplorer", Failure
End Sub


Public Sub ApplyDecimals(ByVal Sheet As Worksheet)
    ' Formats the results and the grid with the number of decimals in D16. A
    ' value that is not a whole number from 0 to 15 leaves the formats as they
    ' are. The x column of a discrete family is shown in whole numbers.
    Dim Places As Variant
    Dim WasProtected As Boolean
    Dim Pattern As String
    Places = Sheet.Range(DECIMALS_CELL).Value2
    If IsError(Places) Then Exit Sub
    If Not IsNumeric(Places) Or IsEmpty(Places) Then Exit Sub
    If CDbl(Places) <> Int(CDbl(Places)) Or CDbl(Places) < 0 Or CDbl(Places) > MAX_DECIMALS Then Exit Sub
    WasProtected = Sheet.ProtectContents
    If WasProtected Then Sheet.Unprotect
    If CLng(Places) = 0 Then Pattern = "0" Else Pattern = "0." & String$(CLng(Places), "0")
    Sheet.Range(FORMATTED_CELLS).NumberFormat = Pattern
    If IsDiscrete(Sheet) Then
        Sheet.Range(X_CELLS).NumberFormat = "0"
    Else
        Sheet.Range(X_CELLS).NumberFormat = Pattern
    End If
    If WasProtected Then Sheet.Protect
End Sub


Public Sub ApplyParameterMessages(ByVal Sheet As Worksheet)
    ' Gives the parameter cells D7:D10 the input and error messages of the
    ' selected family's parameters: name, meaning, unit and limits, as the
    ' family table holds them. A name that is not a family leaves them as
    ' they are.
    Dim Row As Long
    Dim Slot As Long
    Dim WasProtected As Boolean
    Row = FamilyRow(Sheet)
    If Row = 0 Then Exit Sub
    WasProtected = Sheet.ProtectContents
    If WasProtected Then Sheet.Unprotect
    For Slot = 1 To PARAMETER_COUNT
        With Sheet.Cells(FIRST_PARAMETER_ROW + Slot - 1, 4).Validation
            .InputTitle = Left$(CStr(Sheet.Range("ExplorerLabels").Cells(Row, Slot).Value2), 32)
            .InputMessage = Left$(CStr(Sheet.Range("ExplorerPrompts").Cells(Row, Slot).Value2), 255)
            .ErrorTitle = Left$(CStr(Sheet.Range("ExplorerErrorTitles").Cells(Row, Slot).Value2), 32)
            .ErrorMessage = Left$(CStr(Sheet.Range("ExplorerErrors").Cells(Row, Slot).Value2), 225)
            .ShowInput = True
            .ShowError = True
        End With
    Next Slot
    If WasProtected Then Sheet.Protect
End Sub


Private Function FamilyRow(ByVal Sheet As Worksheet) As Long
    ' Row of the selected family in the table, or 0 when D6 is not a family.
    Dim Position As Variant
    Position = Application.Match(Sheet.Range(FAMILY_CELL).Value2, Sheet.Range("ExplorerFamily"), 0)
    If IsError(Position) Then FamilyRow = 0 Else FamilyRow = CLng(Position)
End Function


Private Function IsDiscrete(ByVal Sheet As Worksheet) As Boolean
    Dim Row As Long
    Row = FamilyRow(Sheet)
    If Row > 0 Then IsDiscrete = (CStr(Sheet.Range("ExplorerKind").Cells(Row, 1).Value2) = "Discrete")
End Function


Private Sub LoadFamilyDefaults(ByVal Sheet As Worksheet)
    ' Writes the family's parameters, probe point and plot window. The caller
    ' has switched events off, so these writes do not re-enter the handler.
    Dim Row As Long
    Dim Slot As Long
    Row = FamilyRow(Sheet)
    If Row = 0 Then Exit Sub
    For Slot = 1 To PARAMETER_COUNT
        Sheet.Cells(FIRST_PARAMETER_ROW + Slot - 1, 4).Value2 = _
            Sheet.Range("ExplorerDefaults").Cells(Row, Slot).Value2
    Next Slot
    Sheet.Range("D11").Value2 = Sheet.Range("ExplorerX0").Cells(Row, 1).Value2
    Sheet.Range(POINTS_CELL).Value2 = DEFAULT_POINTS
    RefitPlotWindow Sheet
    ApplyDecimals Sheet
    ApplyParameterMessages Sheet
End Sub


Private Sub RefitPlotWindow(ByVal Sheet As Worksheet)
    ' Sets the plot window to the family's suggestion for the current
    ' parameters. A suggestion that cannot be computed, because a parameter
    ' is not a number, leaves the window as it is.
    Dim Row As Long
    Dim Lower As Variant
    Dim Upper As Variant
    Row = FamilyRow(Sheet)
    If Row = 0 Then Exit Sub
    Sheet.Calculate
    Lower = Sheet.Range("ExplorerSuggestLower").Cells(Row, 1).Value2
    Upper = Sheet.Range("ExplorerSuggestUpper").Cells(Row, 1).Value2
    If IsNumeric(Lower) And IsNumeric(Upper) Then
        If Not IsError(Lower) And Not IsError(Upper) Then
            If Len(CStr(Lower)) > 0 And Len(CStr(Upper)) > 0 Then
                Sheet.Range(LOWER_CELL).Value2 = CDbl(Lower)
                Sheet.Range(UPPER_CELL).Value2 = CDbl(Upper)
            End If
        End If
    End If
End Sub
