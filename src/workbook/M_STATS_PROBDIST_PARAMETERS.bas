Attribute VB_Name = "M_STATS_PROBDIST_PARAMETERS"
'==============================================================================
' MODULE: M_STATS_PROBDIST_PARAMETERS
' PURPOSE: The Parameter Effects sheet's inputs and charts. The curves and the
' table on the sheet are worksheet formulas that call the production API; this
' module loads a family's defaults and the four values of the parameter to
' vary, sets the chart types of its kind, and keeps the number formats and the
' input messages in step with the inputs. It computes nothing. The family
' table (limits, defaults, messages) is the Explorer's, read through the
' Explorer* names; the four curve values, the parameter varied first and the
' line on what a parameter does are in the sheet's own hidden table, read
' through the ParameterEffects* names.
'==============================================================================
Option Explicit

Private Const FAMILY_CELL As String = "D6"
Private Const FIRST_PARAMETER_ROW As Long = 7
Private Const PARAMETER_COUNT As Long = 4
Private Const VARY_CELL As String = "D11"
Private Const FIRST_CURVE_ROW As Long = 12
Private Const CURVE_COUNT As Long = 4
Private Const DECIMALS_CELL As String = "D16"
Private Const DEFAULT_FAMILY As String = "Normal"
Private Const DEFAULT_DECIMALS As Long = 6
Private Const MAX_DECIMALS As Long = 15
Private Const FORMATTED_CELLS As String = "F22:I25"
Private Const DENSITY_CHART As String = "ParameterEffectsDensity"
Private Const CDF_CHART As String = "ParameterEffectsCdf"

Public Function ParameterEffectsSheet() As Worksheet
    Dim Candidate As Worksheet
    For Each Candidate In ThisWorkbook.Worksheets
        If Candidate.CodeName = "shParameters" Then
            Set ParameterEffectsSheet = Candidate
            Exit Function
        End If
    Next Candidate
    Err.Raise vbObjectError + 267, "ParameterEffectsSheet", "Parameter Effects worksheet is missing."
End Function


Public Function ParameterEffectsIsActive() As Boolean
    ' True when this workbook's active sheet is the Parameter Effects sheet:
    ' the Ribbon commands then act on it instead of the Explorer.
    On Error Resume Next
    If ActiveWorkbook Is ThisWorkbook Then
        ParameterEffectsIsActive = (ActiveSheet.CodeName = "shParameters")
    End If
End Function


Public Function ParameterEffectsReady() As Boolean
    ' True when the sheet exists with its two charts, so the commands can run.
    On Error GoTo Failed
    ParameterEffectsReady = (ParameterEffectsSheet().ChartObjects.Count = 2)
    Exit Function
Failed:
    ParameterEffectsReady = False
End Function


Public Sub ResetParameterEffectsInputs(Optional ByVal Confirm As Boolean = True)
    Dim Sheet As Worksheet
    Dim SavedEvents As Boolean
    Dim Failure As String
    Set Sheet = ParameterEffectsSheet()
    If Confirm Then
        If MsgBox("Reset only the Parameter Effects inputs to Normal(0,1)?", vbOKCancel + vbQuestion, _
                  "Parameter Effects") <> vbOK Then Exit Sub
    End If
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Sheet.Range(FAMILY_CELL).Value2 = DEFAULT_FAMILY
    LoadFamilyDefaults Sheet
    Sheet.Range(DECIMALS_CELL).Value2 = DEFAULT_DECIMALS
    ApplyParameterEffectsDecimals Sheet
    Application.EnableEvents = SavedEvents
    RefreshParameterEffects
    Exit Sub
Failed:
    Failure = Err.Description
    Application.EnableEvents = SavedEvents
    Err.Raise vbObjectError + 268, "ResetParameterEffectsInputs", Failure
End Sub


Public Sub ParameterEffectsInputsChanged(ByVal Target As Range)
    ' Called by Workbook_SheetChange. Choosing a family in D6 loads its base
    ' parameters, the parameter to vary and its four values; a paste over
    ' several cells that includes D6 keeps the pasted values and only brings
    ' the charts, the number formats and the messages to the pasted family (and
    ' chooses its first parameter to vary if D11 is not one of its parameters);
    ' choosing the parameter to vary in D11 loads its four values; the decimals
    ' cell sets the number formats; any other edit changes nothing here,
    ' because the curves and the table follow the sheet's own formulas.
    Dim Sheet As Worksheet
    Dim SavedEvents As Boolean
    Dim Failure As String
    ' Save the state first: the handler below restores it, whatever fails.
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Set Sheet = ParameterEffectsSheet()
    If Not Target.Worksheet Is Sheet Then Exit Sub
    If Not Application.Intersect(Target, Sheet.Range(DECIMALS_CELL)) Is Nothing Then ApplyParameterEffectsDecimals Sheet
    If Not Application.Intersect(Target, Sheet.Range(FAMILY_CELL)) Is Nothing Then
        Application.EnableEvents = False
        If Target.Cells.CountLarge = 1 Then
            LoadFamilyDefaults Sheet
        Else
            If VarySlot(Sheet) = 0 Then LoadVaryDefaults Sheet
            ApplyParameterEffectsDecimals Sheet
            ApplyParameterEffectsMessages Sheet
        End If
        Application.EnableEvents = SavedEvents
        RefreshParameterEffects
    ElseIf Target.Cells.CountLarge = 1 And Target.Address(False, False) = VARY_CELL Then
        Application.EnableEvents = False
        LoadCurveValues Sheet
        ApplyParameterEffectsMessages Sheet
        Application.EnableEvents = SavedEvents
        RefreshParameterEffects
    End If
    Exit Sub
Failed:
    Failure = Err.Description
    Application.EnableEvents = SavedEvents
    MsgBox "The Parameter Effects sheet could not follow that change: " & Failure, vbExclamation, _
           "Parameter Effects"
End Sub


Public Sub ParameterEffectsOpened()
    ' Called by Workbook_Open. A file saved while macros were off, or with
    ' another family's chart types, opens with stale charts, number formats that
    ' do not follow the decimals cell and messages of another family: set them
    ' again, and leave the Saved flag as it was.
    Dim WasSaved As Boolean
    On Error Resume Next
    WasSaved = ThisWorkbook.Saved
    RefreshParameterEffects
    ApplyParameterEffectsDecimals ParameterEffectsSheet()
    ApplyParameterEffectsMessages ParameterEffectsSheet()
    ThisWorkbook.Saved = WasSaved
    On Error GoTo 0
End Sub


Public Sub RefreshParameterEffects()
    ' Recalculates the sheet and sets the two charts for the family's kind: a
    ' discrete family draws its mass as columns and its cumulative function as
    ' lines with markers, a continuous one draws lines. The series follow the
    ' spilled grid through workbook names, so no range is set here.
    Dim Sheet As Worksheet
    Dim Discrete As Boolean
    Dim SavedEvents As Boolean
    Dim WasProtected As Boolean
    Dim Failure As String
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Set Sheet = ParameterEffectsSheet()
    If Sheet.ChartObjects.Count <> 2 Then Exit Sub
    WasProtected = Sheet.ProtectContents
    Application.EnableEvents = False
    If WasProtected Then Sheet.Unprotect
    Sheet.Calculate
    Discrete = IsDiscrete(Sheet)
    FormatChart Sheet.ChartObjects(DENSITY_CHART).Chart, Discrete, True
    FormatChart Sheet.ChartObjects(CDF_CHART).Chart, Discrete, False
    If WasProtected Then Sheet.Protect
    Application.EnableEvents = SavedEvents
    Exit Sub
Failed:
    Failure = Err.Description
    On Error Resume Next
    If WasProtected Then Sheet.Protect
    Application.EnableEvents = SavedEvents
    On Error GoTo 0
    Err.Raise vbObjectError + 269, "RefreshParameterEffects", Failure
End Sub


Private Sub FormatChart(ByVal Target As Chart, ByVal Discrete As Boolean, ByVal IsDensity As Boolean)
    Dim Index As Long
    Dim Colors As Variant
    Dim ChartKind As Long
    Colors = Array(RGB(31, 56, 100), RGB(237, 125, 49), RGB(112, 173, 71), RGB(112, 48, 160))
    If Target.SeriesCollection.Count <> CURVE_COUNT Then Exit Sub
    If IsDensity And Discrete Then
        ChartKind = xlColumnClustered
    ElseIf Discrete Then
        ChartKind = xlLineMarkers
    Else
        ChartKind = xlLine
    End If
    Target.PlotVisibleOnly = False
    For Index = 1 To CURVE_COUNT
        Target.SeriesCollection(Index).ChartType = ChartKind
    Next Index
    For Index = 1 To CURVE_COUNT
        If ChartKind = xlColumnClustered Then
            Target.SeriesCollection(Index).Format.Fill.ForeColor.RGB = Colors(Index - 1)
        Else
            Target.SeriesCollection(Index).Format.Line.ForeColor.RGB = Colors(Index - 1)
            Target.SeriesCollection(Index).Format.Line.Weight = 2.25
        End If
    Next Index
    If ChartKind = xlColumnClustered Then
        Target.ChartGroups(1).Overlap = 0
        Target.ChartGroups(1).GapWidth = 60
    End If
    If IsDensity Then
        If Discrete Then
            Target.ChartTitle.Text = "Probability mass P(X = k) of the four curves"
        Else
            Target.ChartTitle.Text = "Density f(x) of the four curves"
        End If
    ElseIf Discrete Then
        Target.ChartTitle.Text = "Cumulative P(X <= k) of the four curves"
    Else
        Target.ChartTitle.Text = "Cumulative F(x) of the four curves"
    End If
    Target.HasLegend = True
    Target.Legend.Position = xlLegendPositionBottom
    Target.Axes(xlValue).MinimumScale = 0#
    Target.Axes(xlValue).MaximumScaleIsAuto = True
    Target.Axes(xlValue).TickLabels.NumberFormat = "General"
    Target.Axes(xlCategory).TickLabels.NumberFormat = "General"
    Target.Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
    If Discrete Then
        Target.Axes(xlCategory).TickLabelSpacingIsAuto = True
    Else
        Target.Axes(xlCategory).TickLabelSpacing = 25
        Target.Axes(xlCategory).TickMarkSpacing = 25
    End If
End Sub


Public Sub ApplyParameterEffectsDecimals(ByVal Sheet As Worksheet)
    ' Formats the table with the number of decimals in D16. A value that is not a
    ' whole number from 0 to 15 leaves the formats as they are.
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
    If WasProtected Then Sheet.Protect
    Exit Sub
Failed:
    ' Protect the sheet again as it was found, then pass the error to the caller.
    Failure = Err.Description
    On Error Resume Next
    If WasProtected Then Sheet.Protect
    On Error GoTo 0
    Err.Raise vbObjectError + 264, "ApplyParameterEffectsDecimals", Failure
End Sub


Public Sub ApplyParameterEffectsMessages(ByVal Sheet As Worksheet)
    ' Gives the base parameter cells D7:D10 the input and error messages of the
    ' selected family's parameters, and the four curve cells those of the
    ' parameter to vary, as the Explorer's family table holds them. A name that
    ' is not a family leaves them as they are.
    Dim Row As Long
    Dim Slot As Long
    Dim Curve As Long
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
    Slot = VarySlot(Sheet)
    If Slot > 0 Then
        For Curve = 1 To CURVE_COUNT
            With Sheet.Cells(FIRST_CURVE_ROW + Curve - 1, 4).Validation
                .InputTitle = Left$("Curve " & Curve & " value", 32)
                .InputMessage = Left$("Value of " & CStr(TableRange("ExplorerLabels").Cells(Row, Slot).Value2) & _
                                      " for curve " & Curve & ". " & _
                                      CStr(TableRange("ExplorerPrompts").Cells(Row, Slot).Value2), 255)
                .ErrorTitle = Left$("Invalid curve " & Curve & " value", 32)
                .ErrorMessage = Left$(CStr(TableRange("ExplorerErrors").Cells(Row, Slot).Value2), 225)
                .ShowInput = True
                .ShowError = True
            End With
        Next Curve
    End If
    If WasProtected Then Sheet.Protect
    Exit Sub
Failed:
    ' Protect the sheet again as it was found, then pass the error to the caller.
    Failure = Err.Description
    On Error Resume Next
    If WasProtected Then Sheet.Protect
    On Error GoTo 0
    Err.Raise vbObjectError + 265, "ApplyParameterEffectsMessages", Failure
End Sub


Private Function TableRange(ByVal TableName As String) As Range
    ' A range of a table by its workbook name: the names refer to other
    ' ranges than the sheet's own, which a sheet's Range(Name) does not follow.
    Set TableRange = ThisWorkbook.Names.Item(TableName).RefersToRange
End Function


Private Function FamilyRow(ByVal Sheet As Worksheet) As Long
    ' Row of the selected family in the table, or 0 when D6 is not a family.
    Dim Position As Variant
    Position = Application.Match(Sheet.Range(FAMILY_CELL).Value2, TableRange("ExplorerFamily"), 0)
    If IsError(Position) Then FamilyRow = 0 Else FamilyRow = CLng(Position)
End Function


Private Function VarySlot(ByVal Sheet As Worksheet) As Long
    ' Number of the parameter chosen in D11 among the labels of the selected
    ' family, or 0 when D11 is not one of them.
    Dim Row As Long
    Dim Position As Variant
    Row = FamilyRow(Sheet)
    If Row = 0 Then Exit Function
    Position = Application.Match(Sheet.Range(VARY_CELL).Value2, TableRange("ExplorerLabels").Rows(Row), 0)
    If IsError(Position) Then Exit Function
    If CLng(Position) <= CLng(TableRange("ExplorerNP").Cells(Row, 1).Value2) Then VarySlot = CLng(Position)
End Function


Private Function IsDiscrete(ByVal Sheet As Worksheet) As Boolean
    Dim Row As Long
    Row = FamilyRow(Sheet)
    If Row > 0 Then IsDiscrete = (CStr(TableRange("ExplorerKind").Cells(Row, 1).Value2) = "Discrete")
End Function


Private Sub LoadFamilyDefaults(ByVal Sheet As Worksheet)
    ' Writes the family's base parameters, the parameter to vary and its four
    ' values. The caller has switched events off, so these writes do not
    ' re-enter the handler.
    Dim Row As Long
    Dim Slot As Long
    Row = FamilyRow(Sheet)
    If Row = 0 Then Exit Sub
    For Slot = 1 To PARAMETER_COUNT
        Sheet.Cells(FIRST_PARAMETER_ROW + Slot - 1, 4).Value2 = _
            TableRange("ExplorerDefaults").Cells(Row, Slot).Value2
    Next Slot
    LoadVaryDefaults Sheet
    ApplyParameterEffectsDecimals Sheet
    ApplyParameterEffectsMessages Sheet
End Sub


Private Sub LoadVaryDefaults(ByVal Sheet As Worksheet)
    ' Chooses the parameter the family varies first and writes its four values.
    Dim Row As Long
    Dim Slot As Long
    Row = FamilyRow(Sheet)
    If Row = 0 Then Exit Sub
    Slot = CLng(TableRange("ParameterEffectsVary").Cells(Row, 1).Value2)
    Sheet.Range(VARY_CELL).Value2 = TableRange("ExplorerLabels").Cells(Row, Slot).Value2
    LoadCurveValues Sheet
End Sub


Private Sub LoadCurveValues(ByVal Sheet As Worksheet)
    ' Writes the four default values of the parameter chosen in D11.
    Dim Row As Long
    Dim Slot As Long
    Dim Curve As Long
    Row = FamilyRow(Sheet)
    Slot = VarySlot(Sheet)
    If Row = 0 Or Slot = 0 Then Exit Sub
    For Curve = 1 To CURVE_COUNT
        Sheet.Cells(FIRST_CURVE_ROW + Curve - 1, 4).Value2 = _
            TableRange("ParameterEffectsCurves").Cells(Row, (Slot - 1) * CURVE_COUNT + Curve).Value2
    Next Curve
End Sub
