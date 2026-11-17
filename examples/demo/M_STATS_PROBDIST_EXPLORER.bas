Attribute VB_Name = "M_STATS_PROBDIST_EXPLORER"
'==============================================================================
' MODULE: M_STATS_PROBDIST_EXPLORER
' PURPOSE: Workbook-owned Explorer chart refresh and bounded input reset.
' Numerical results remain visible worksheet calls to the production API.
'==============================================================================
Option Explicit

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
            ResetExplorerInputs
        Case "pdRecalculate"
            RecalculateWorkbook
            RefreshExplorer
    End Select
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Distribution Explorer"
End Sub


Public Sub ExplorerEnabled(ByVal Control As IRibbonControl, ByRef ReturnedVal As Variant)
    Dim Target As Worksheet
    On Error GoTo Failed
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
    Sheet.Range("D6").Value2 = "Normal"
    Sheet.Range("D7").Value2 = 0#
    Sheet.Range("D8").Value2 = 1#
    Sheet.Range("D9").Value2 = 1#
    Sheet.Range("D10").Value2 = 4#
    Sheet.Range("D11").Value2 = 0#
    Application.EnableEvents = SavedEvents
    RefreshExplorer
    Exit Sub
Failed:
    Failure = Err.Description
    Application.EnableEvents = SavedEvents
    Err.Raise vbObjectError + 255, "ResetExplorerInputs", Failure
End Sub


Public Sub RefreshExplorer()
    Dim Sheet As Worksheet
    Dim Discrete As Boolean
    Dim SavedEvents As Boolean
    Dim WasProtected As Boolean
    Dim Item As ChartObject
    Dim Failure As String
    Dim DecimalSeparator As String
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Set Sheet = ExplorerSheet()
    If Sheet.ChartObjects.Count <> 2 Then Exit Sub
    WasProtected = Sheet.ProtectContents
    Application.EnableEvents = False
    Sheet.Unprotect
    Sheet.Calculate
    Discrete = (CStr(Sheet.Range("D6").Value2) = "Poisson")
    DecimalSeparator = CStr(Application.International(xlDecimalSeparator))
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
        Item.Chart.SeriesCollection(1).XValues = Sheet.Range("F18:F" & ExplorerLastRow(Sheet))
        If Item.Name = "ExplorerDensity" Then
            Item.Chart.SeriesCollection(1).Values = Sheet.Range("G18:G" & ExplorerLastRow(Sheet))
            Item.Chart.ChartTitle.Text = "Density / probability mass"
        Else
            Item.Chart.SeriesCollection(1).Values = Sheet.Range("H18:H" & ExplorerLastRow(Sheet))
            Item.Chart.ChartTitle.Text = "Cumulative probability P(X <= x)"
        End If
        Item.Chart.Axes(xlCategory).TickLabels.NumberFormatLocal = "0" & DecimalSeparator & "##"
        Item.Chart.Axes(xlValue).TickLabels.NumberFormatLocal = "0" & DecimalSeparator & "###"
        Item.Chart.Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        Item.Chart.Axes(xlValue).MinimumScale = 0#
        If Item.Name = "ExplorerCDF" Then Item.Chart.Axes(xlValue).MaximumScale = 1#
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


Private Function ExplorerLastRow(ByVal Sheet As Worksheet) As Long
    ExplorerLastRow = 218
    If CStr(Sheet.Range("D6").Value2) = "Poisson" Then
        If CStr(Sheet.Range("G5").Value2) = "OK" Then
            ExplorerLastRow = 18 + CLng(Sheet.Range("G11").Value2)
        End If
    End If
End Function
