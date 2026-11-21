Attribute VB_Name = "M_STATS_PROBDIST_DIAGNOSTICS"
'==============================================================================
' MODULE: M_STATS_PROBDIST_DIAGNOSTICS
' PURPOSE: The Diagnostics sheet (errors and diagnostics): its reset, the
' example arguments a new function loads, its chart, the line that names the
' Excel computing the sheet, and DiagnosticsStatus, the worksheet function that
' reads the Status argument of a library call. A worksheet formula cannot read
' Status, so DiagnosticsStatus makes the same call from VBA and returns the
' text the library wrote there, unchanged. The values on the sheet are
' worksheet formulas that call the production API directly; this module
' computes none of them.
'==============================================================================
Option Explicit

Private Const FUNCTION_CELL As String = "D6"
Private Const FIRST_ARGUMENT_CELL As String = "D7"
Private Const SWEEP_FROM_CELL As String = "D11"
Private Const SWEEP_TO_CELL As String = "D12"
Private Const BUILD_CELL As String = "D32"
Private Const CATALOGUE_NAMES As String = "DiagFunctionList"
Private Const ARGUMENT_COUNT As Long = 4
Private Const DEFAULT_FUNCTION As String = "K_STATS_Gamma_InverseCumulative"
Private Const SWEEP_CHART As String = "DiagSweep"
Private Const TEXT_ARGUMENT As String = "Not called: argument {n} is text that is not a number; " & _
    "VBA cannot pass it as a Double, so the function never runs and Excel returns #VALUE!."
Private Const ERROR_ARGUMENT As String = "Not called: argument {n} holds an error value, " & _
    "which Excel cannot pass as a number."
Private Const UNKNOWN_FUNCTION As String = "Not called: the function is not in the list of this sheet."

Public Function DiagnosticsSheet() As Worksheet
    Dim Candidate As Worksheet
    For Each Candidate In ThisWorkbook.Worksheets
        If Candidate.CodeName = "shDiagnostics" Then
            Set DiagnosticsSheet = Candidate
            Exit Function
        End If
    Next Candidate
    Err.Raise vbObjectError + 400, "DiagnosticsSheet", "Diagnostics worksheet is missing."
End Function


Public Function DiagnosticsIsActive() As Boolean
    ' True when this workbook's active sheet is the Diagnostics sheet: the
    ' Ribbon commands then act on it instead of the Explorer.
    On Error Resume Next
    If ActiveWorkbook Is ThisWorkbook Then
        DiagnosticsIsActive = (ActiveSheet.CodeName = "shDiagnostics")
    End If
End Function


Public Function DiagnosticsReady() As Boolean
    ' True when the sheet exists with its chart, so the commands can run.
    On Error GoTo Failed
    DiagnosticsReady = (DiagnosticsSheet().ChartObjects.Count = 1)
    Exit Function
Failed:
    DiagnosticsReady = False
End Function


Public Function DiagnosticsStatus( _
    ByVal FunctionName As String, _
    Optional ByVal Argument1 As Variant, _
    Optional ByVal Argument2 As Variant, _
    Optional ByVal Argument3 As Variant, _
    Optional ByVal Argument4 As Variant) _
    As String
    ' The Status text the library writes for FunctionName called with these
    ' arguments: "" when the call succeeds, the reason when it returns an
    ' error. An empty cell is passed as 0, as Excel passes it. An argument
    ' that VBA cannot pass as a Double (a text that is not a number, an error
    ' value) stops the worksheet call before the function runs; for it this
    ' returns a line that says so, starting "Not called", since the library
    ' wrote nothing.
    Dim Values(1 To ARGUMENT_COUNT) As Double
    Dim Raw As Variant
    Dim Slot As Long
    Dim Count As Long
    Dim Status As String
    Dim Returned As Variant
    On Error GoTo Failed
    Count = DiagnosticsArity(FunctionName)
    If Count = 0 Then
        DiagnosticsStatus = UNKNOWN_FUNCTION
        Exit Function
    End If
    For Slot = 1 To Count
        Select Case Slot
            Case 1: Raw = ArgumentValue(Argument1)
            Case 2: Raw = ArgumentValue(Argument2)
            Case 3: Raw = ArgumentValue(Argument3)
            Case Else: Raw = ArgumentValue(Argument4)
        End Select
        If IsError(Raw) Then
            DiagnosticsStatus = Replace(ERROR_ARGUMENT, "{n}", CStr(Slot))
            Exit Function
        End If
        If VarType(Raw) = vbString Then
            If Not IsNumeric(Raw) Then
                DiagnosticsStatus = Replace(TEXT_ARGUMENT, "{n}", CStr(Slot))
                Exit Function
            End If
        End If
        Values(Slot) = CDbl(Raw)
    Next Slot
    Returned = DiagnosticsCall(FunctionName, Values, Status)
    DiagnosticsStatus = Status
    Exit Function
Failed:
    DiagnosticsStatus = "Not called: " & Err.Description
End Function


Private Function ArgumentValue(ByVal Argument As Variant) As Variant
    ' The value of an argument: the first cell of a range, 0 for a missing
    ' argument or an empty cell.
    If IsMissing(Argument) Then
        ArgumentValue = 0#
    ElseIf IsObject(Argument) Then
        ArgumentValue = Argument.Cells(1, 1).Value2
        If IsEmpty(ArgumentValue) Then ArgumentValue = 0#
    ElseIf IsEmpty(Argument) Then
        ArgumentValue = 0#
    Else
        ArgumentValue = Argument
    End If
End Function


Private Function DiagnosticsArity(ByVal FunctionName As String) As Long
    ' The number of arguments, Status aside, of each function of the sheet's
    ' list; 0 for any other name.
    Select Case FunctionName
        Case "K_STATS_NormalStandard_Cumulative", "K_STATS_NormalStandard_Survival", _
             "K_STATS_NormalStandard_InverseCumulative"
            DiagnosticsArity = 1
        Case "K_STATS_NormalStandard_IntervalProbability", "K_STATS_Lognormal_Mean", _
             "K_STATS_StudentT_Cumulative", "K_STATS_StudentT_InverseCumulative", "K_STATS_ChiSquare_Density", _
             "K_STATS_Exponential_Density", "K_STATS_Exponential_Survival", "K_STATS_Poisson_PMF", _
             "K_STATS_Poisson_Cumulative"
            DiagnosticsArity = 2
        Case "K_STATS_Normal_Density", "K_STATS_Normal_Cumulative", "K_STATS_Normal_InverseCumulative", _
             "K_STATS_Gamma_Density", "K_STATS_Gamma_Cumulative", "K_STATS_Gamma_InverseCumulative", _
             "K_STATS_Beta_Density", "K_STATS_Beta_Cumulative", "K_STATS_Uniform_Cumulative", _
             "K_STATS_Binomial_PMF", "K_STATS_Binomial_Cumulative", "K_STATS_Binomial_InverseCumulative", _
             "K_STATS_DiscreteUniform_PMF", "K_STATS_DiscreteUniform_LogPMF", "K_STATS_DiscreteUniform_Cumulative"
            DiagnosticsArity = 3
        Case "K_STATS_Hypergeometric_PMF"
            DiagnosticsArity = 4
        Case Else
            DiagnosticsArity = 0
    End Select
End Function


Private Function DiagnosticsCall( _
    ByVal FunctionName As String, _
    ByRef Values() As Double, _
    ByRef Status As String) _
    As Variant
    ' The library call itself, with its Status argument.
    Select Case FunctionName
        Case "K_STATS_NormalStandard_Cumulative"
            DiagnosticsCall = K_STATS_NormalStandard_Cumulative(Values(1), Status)
        Case "K_STATS_NormalStandard_Survival"
            DiagnosticsCall = K_STATS_NormalStandard_Survival(Values(1), Status)
        Case "K_STATS_NormalStandard_InverseCumulative"
            DiagnosticsCall = K_STATS_NormalStandard_InverseCumulative(Values(1), Status)
        Case "K_STATS_NormalStandard_IntervalProbability"
            DiagnosticsCall = K_STATS_NormalStandard_IntervalProbability(Values(1), Values(2), Status)
        Case "K_STATS_Normal_Density"
            DiagnosticsCall = K_STATS_Normal_Density(Values(1), Values(2), Values(3), Status)
        Case "K_STATS_Normal_Cumulative"
            DiagnosticsCall = K_STATS_Normal_Cumulative(Values(1), Values(2), Values(3), Status)
        Case "K_STATS_Normal_InverseCumulative"
            DiagnosticsCall = K_STATS_Normal_InverseCumulative(Values(1), Values(2), Values(3), Status)
        Case "K_STATS_Lognormal_Mean"
            DiagnosticsCall = K_STATS_Lognormal_Mean(Values(1), Values(2), Status)
        Case "K_STATS_StudentT_Cumulative"
            DiagnosticsCall = K_STATS_StudentT_Cumulative(Values(1), Values(2), Status)
        Case "K_STATS_StudentT_InverseCumulative"
            DiagnosticsCall = K_STATS_StudentT_InverseCumulative(Values(1), Values(2), Status)
        Case "K_STATS_ChiSquare_Density"
            DiagnosticsCall = K_STATS_ChiSquare_Density(Values(1), Values(2), Status)
        Case "K_STATS_Gamma_Density"
            DiagnosticsCall = K_STATS_Gamma_Density(Values(1), Values(2), Values(3), Status)
        Case "K_STATS_Gamma_Cumulative"
            DiagnosticsCall = K_STATS_Gamma_Cumulative(Values(1), Values(2), Values(3), Status)
        Case "K_STATS_Gamma_InverseCumulative"
            DiagnosticsCall = K_STATS_Gamma_InverseCumulative(Values(1), Values(2), Values(3), Status)
        Case "K_STATS_Beta_Density"
            DiagnosticsCall = K_STATS_Beta_Density(Values(1), Values(2), Values(3), Status)
        Case "K_STATS_Beta_Cumulative"
            DiagnosticsCall = K_STATS_Beta_Cumulative(Values(1), Values(2), Values(3), Status)
        Case "K_STATS_Exponential_Density"
            DiagnosticsCall = K_STATS_Exponential_Density(Values(1), Values(2), Status)
        Case "K_STATS_Exponential_Survival"
            DiagnosticsCall = K_STATS_Exponential_Survival(Values(1), Values(2), Status)
        Case "K_STATS_Uniform_Cumulative"
            DiagnosticsCall = K_STATS_Uniform_Cumulative(Values(1), Values(2), Values(3), Status)
        Case "K_STATS_Binomial_PMF"
            DiagnosticsCall = K_STATS_Binomial_PMF(Values(1), Values(2), Values(3), Status)
        Case "K_STATS_Binomial_Cumulative"
            DiagnosticsCall = K_STATS_Binomial_Cumulative(Values(1), Values(2), Values(3), Status)
        Case "K_STATS_Binomial_InverseCumulative"
            DiagnosticsCall = K_STATS_Binomial_InverseCumulative(Values(1), Values(2), Values(3), Status)
        Case "K_STATS_Poisson_PMF"
            DiagnosticsCall = K_STATS_Poisson_PMF(Values(1), Values(2), Status)
        Case "K_STATS_Poisson_Cumulative"
            DiagnosticsCall = K_STATS_Poisson_Cumulative(Values(1), Values(2), Status)
        Case "K_STATS_Hypergeometric_PMF"
            DiagnosticsCall = K_STATS_Hypergeometric_PMF(Values(1), Values(2), Values(3), Values(4), Status)
        Case "K_STATS_DiscreteUniform_PMF"
            DiagnosticsCall = K_STATS_DiscreteUniform_PMF(Values(1), Values(2), Values(3), Status)
        Case "K_STATS_DiscreteUniform_LogPMF"
            DiagnosticsCall = K_STATS_DiscreteUniform_LogPMF(Values(1), Values(2), Values(3), Status)
        Case "K_STATS_DiscreteUniform_Cumulative"
            DiagnosticsCall = K_STATS_DiscreteUniform_Cumulative(Values(1), Values(2), Values(3), Status)
        Case Else
            Status = UNKNOWN_FUNCTION
            DiagnosticsCall = CVErr(xlErrNA)
    End Select
End Function


Public Sub ResetDiagnosticsInputs(Optional ByVal Confirm As Boolean = True)
    ' Restores the default example: the Gamma quantile at 0.95 with shape 2
    ' and scale 1.5, the probability swept from 0 to 1.
    Dim Sheet As Worksheet
    Dim SavedEvents As Boolean
    Dim Failure As String
    Set Sheet = DiagnosticsSheet()
    If Confirm Then
        If MsgBox("Reset only the Diagnostics inputs to the default example (the Gamma quantile at 0.95)?", _
                  vbOKCancel + vbQuestion, "Errors and Diagnostics") <> vbOK Then Exit Sub
    End If
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Sheet.Range(FUNCTION_CELL).Value2 = DEFAULT_FUNCTION
    LoadDiagnosticsExample Sheet
    Application.EnableEvents = SavedEvents
    RefreshDiagnostics
    Exit Sub
Failed:
    Failure = Err.Description
    Application.EnableEvents = SavedEvents
    Err.Raise vbObjectError + 401, "ResetDiagnosticsInputs", Failure
End Sub


Private Sub LoadDiagnosticsExample(ByVal Sheet As Worksheet)
    ' Writes the example arguments and the sweep of the chosen function, from
    ' the hidden catalogue of the sheet, into the inputs. A function not in
    ' the list leaves the inputs as they are.
    Dim Names As Range
    Dim Found As Range
    Dim Slot As Long
    Dim Value As Variant
    Set Names = Sheet.Parent.Names(CATALOGUE_NAMES).RefersToRange
    Set Found = Names.Find(What:=CStr(Sheet.Range(FUNCTION_CELL).Value2), LookIn:=xlValues, _
                           LookAt:=xlWhole, MatchCase:=True)
    If Found Is Nothing Then Exit Sub
    For Slot = 1 To ARGUMENT_COUNT
        Value = Found.Offset(0, 5 + Slot).Value2
        If IsEmpty(Value) Then
            Sheet.Range(FIRST_ARGUMENT_CELL).Offset(Slot - 1, 0).ClearContents
        Else
            Sheet.Range(FIRST_ARGUMENT_CELL).Offset(Slot - 1, 0).Value2 = Value
        End If
    Next Slot
    Sheet.Range(SWEEP_FROM_CELL).Value2 = Found.Offset(0, 10).Value2
    Sheet.Range(SWEEP_TO_CELL).Value2 = Found.Offset(0, 11).Value2
End Sub


Public Sub DiagnosticsInputsChanged(ByVal Target As Range)
    ' Called by Workbook_SheetChange. A new function loads its example
    ' arguments and sweep; any input refreshes the chart's title.
    Dim Sheet As Worksheet
    Dim SavedEvents As Boolean
    On Error GoTo Failed
    Set Sheet = DiagnosticsSheet()
    If Not Target.Worksheet Is Sheet Then Exit Sub
    If Application.Intersect(Target, Sheet.Range(FUNCTION_CELL & ":" & SWEEP_TO_CELL)) Is Nothing Then Exit Sub
    If Not Application.Intersect(Target, Sheet.Range(FUNCTION_CELL)) Is Nothing Then
        SavedEvents = Application.EnableEvents
        Application.EnableEvents = False
        LoadDiagnosticsExample Sheet
        Application.EnableEvents = SavedEvents
    End If
    RefreshDiagnostics
    Exit Sub
Failed:
    Application.EnableEvents = True
    MsgBox "The Diagnostics sheet could not follow that change: " & Err.Description, vbExclamation, _
           "Errors and Diagnostics"
End Sub


Public Sub DiagnosticsOpened()
    ' Called by Workbook_Open: names the Excel that now computes the sheet,
    ' when it is not the one that saved it. The chart keeps the type and title
    ' it was saved with, so nothing else is set here: formatting it at every
    ' opening would leave the workbook marked as changed.
    Dim Sheet As Worksheet
    Dim WasSaved As Boolean
    On Error Resume Next
    WasSaved = ThisWorkbook.Saved
    Set Sheet = DiagnosticsSheet()
    If CStr(Sheet.Range(BUILD_CELL).Value2) <> DiagnosticsExcelBuild() Then
        Sheet.Range(BUILD_CELL).Value2 = DiagnosticsExcelBuild()
        ThisWorkbook.Saved = WasSaved
    End If
    On Error GoTo 0
End Sub


Public Function DiagnosticsExcelBuild() As String
    ' The Excel that computes the sheet: version, build, bitness and operating
    ' system, as the issue asks every observation to record.
    Dim Bitness As String
#If Win64 Then
    Bitness = "64-bit"
#Else
    Bitness = "32-bit"
#End If
    DiagnosticsExcelBuild = "Microsoft Excel " & Application.Version & ", build " & CStr(Application.Build) & _
                            ", " & Bitness & ", " & Application.OperatingSystem
End Function


Public Sub RefreshDiagnostics()
    ' Recalculates the sheet, writes the Excel that computed it and sets the
    ' chart: the function across the sweep of argument 1, each error a red
    ' mark at 0.
    Dim Sheet As Worksheet
    Dim SavedEvents As Boolean
    Dim WasProtected As Boolean
    Dim Failure As String
    Dim Title As String
    SavedEvents = Application.EnableEvents
    On Error GoTo Failed
    Set Sheet = DiagnosticsSheet()
    WasProtected = Sheet.ProtectContents
    Application.EnableEvents = False
    If WasProtected Then Sheet.Unprotect
    If CStr(Sheet.Range(BUILD_CELL).Value2) <> DiagnosticsExcelBuild() Then
        Sheet.Range(BUILD_CELL).Value2 = DiagnosticsExcelBuild()
    End If
    Sheet.Calculate
    Title = CStr(Sheet.Range(FUNCTION_CELL).Value2) & " across " & ArgumentName(Sheet) & " from " & _
            CellText(Sheet, SWEEP_FROM_CELL) & " to " & CellText(Sheet, SWEEP_TO_CELL) & "; errors in red at 0"
    If Sheet.ChartObjects.Count = 1 Then FormatSweepChart Sheet.ChartObjects(SWEEP_CHART).Chart, Title, _
                                                         ArgumentName(Sheet)
    If WasProtected Then Sheet.Protect
    Application.EnableEvents = SavedEvents
    Exit Sub
Failed:
    Failure = Err.Description
    On Error Resume Next
    If WasProtected Then Sheet.Protect
    Application.EnableEvents = SavedEvents
    On Error GoTo 0
    Err.Raise vbObjectError + 402, "RefreshDiagnostics", Failure
End Sub


Private Function ArgumentName(ByVal Sheet As Worksheet) As String
    ' The name of argument 1, from its label "Argument 1: Name".
    Dim Label As String
    Label = CStr(Sheet.Range("C7").Text)
    If InStr(Label, ": ") > 0 Then
        ArgumentName = Mid$(Label, InStr(Label, ": ") + 2)
    Else
        ArgumentName = "argument 1"
    End If
End Function


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


Private Sub FormatSweepChart(ByVal Chart As Chart, ByVal Title As String, ByVal XTitle As String)
    ' Series: the value returned, a line; the errors, red marks at 0.
    With Chart
        .ChartType = xlXYScatterLinesNoMarkers
        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(68, 114, 196)
        With .SeriesCollection(2)
            .ChartType = xlXYScatter
            .MarkerStyle = xlMarkerStyleX
            .MarkerSize = 7
            .MarkerForegroundColor = RGB(192, 0, 0)
            .MarkerBackgroundColor = RGB(192, 0, 0)
        End With
        .HasTitle = True
        .ChartTitle.Text = Title
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = XTitle
        .Axes(xlCategory).TickLabels.NumberFormat = "General"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Returned"
        .Axes(xlValue).TickLabels.NumberFormat = "General"
        .Axes(xlValue).ScaleType = xlScaleLinear
        .Axes(xlValue).MinimumScaleIsAuto = True
        .Axes(xlValue).MaximumScaleIsAuto = True
    End With
End Sub
