VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
'==============================================================================
' MODULE: ThisWorkbook
' PURPOSE: The workbook events of the Explorer workbook. They only call the
' modules of src/workbook/: the Function Wizard registration and the
' Explorer's, the Probabilities, the Parameter Effects, the Comparison and the Tails sheets' refresh when the workbook opens,
' their input handling when a cell of their sheet changes, and the navigation bar around a save.
' tools/demobuild/explorer_build.py and tools/Sync-PROBDIST-VBA.ps1 write
' this code into the workbook's ThisWorkbook module.
'==============================================================================
Option Explicit

Private Sub Workbook_Open()
    M_STATS_PROBDIST_REGISTER.PROB_Register_PublicUDFs
    M_STATS_PROBDIST_EXPLORER.ExplorerOpened
    M_STATS_PROBDIST_PROBABILITIES.ProbabilitiesOpened
    M_STATS_PROBDIST_PARAMETERS.ParameterEffectsOpened
    M_STATS_PROBDIST_COMPARISON.ComparisonOpened
    M_STATS_PROBDIST_TAILS.TailsOpened
End Sub


Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Sh.CodeName = "shExplorer" Then M_STATS_PROBDIST_EXPLORER.ExplorerInputsChanged Target
    If Sh.CodeName = "shProbability" Then M_STATS_PROBDIST_PROBABILITIES.ProbabilitiesInputsChanged Target
    If Sh.CodeName = "shParameters" Then M_STATS_PROBDIST_PARAMETERS.ParameterEffectsInputsChanged Target
    If Sh.CodeName = "shComparison" Then M_STATS_PROBDIST_COMPARISON.ComparisonInputsChanged Target
    If Sh.CodeName = "shTails" Then M_STATS_PROBDIST_TAILS.TailsInputsChanged Target
End Sub


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Save the navigation bar shown even while it is hidden.
    M_STATS_PROBDIST_DEMO.NavigationBarBeforeSave
End Sub


Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    ' Hide the navigation bar again if it was hidden before the save.
    M_STATS_PROBDIST_DEMO.NavigationBarAfterSave
End Sub
