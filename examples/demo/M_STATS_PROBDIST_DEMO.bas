Attribute VB_Name = "M_STATS_PROBDIST_DEMO"
Option Explicit

' Workbook-owned callbacks. No application state is changed or retained.
Public Sub DemoGoTo(ByVal Control As Object)
    Dim Target As Worksheet
    On Error GoTo Failed
    For Each Target In ThisWorkbook.Worksheets
        If Target.CodeName = CStr(Control.Tag) Then
            ThisWorkbook.Activate
            Target.Activate
            Exit Sub
        End If
    Next Target
    Err.Raise vbObjectError + 253, "DemoGoTo", "The requested demo sheet is missing."
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Probability Distributions"
End Sub


Public Sub DemoRecalculate(ByVal Control As Object)
    Dim Target As Worksheet
    On Error GoTo Failed
    For Each Target In ThisWorkbook.Worksheets
        Target.Calculate
    Next Target
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Probability Distributions"
End Sub
