Attribute VB_Name = "M_NAVIGATION"
Option Explicit

'==============================================================================
' M_NAVIGATION
'------------------------------------------------------------------------------
' PURPOSE
'   Creates and manages a dynamic collapsible worksheet navigation pane
'
' WHY THIS EXISTS
'   Workbooks containing multiple analytical or teaching worksheets are easier
'   to use when navigation is visible consistent and independent of Excel's
'   worksheet-tab bar
'
'   This module creates an application-style sidebar that can be collapsed when
'   additional worksheet space is required
'
' PUBLIC PROCEDURES
'   - InstallNavigation
'   - ToggleNavigation
'   - RefreshActiveSheet
'   - RefreshAllNavigation
'
' PRIVATE PROCEDURES AND FUNCTIONS
'   - InstallNavigationOnSheet
'   - BuildNavigation
'   - ClearNavigationArea
'   - EnsureToggleButton
'   - PositionToggleButton
'   - UpdateToggleAppearance
'   - CollapseNavigation
'   - ExpandNavigation
'   - GetToggleShape
'   - HasNavigationPane
'   - IsNavigationCollapsed
'   - IsNavigationSheet
'   - GetExistingNavigationLastRow
'   - GetNavigationSignature
'   - NavigationIcon
'   - SafeWorksheetName
'   - ReportNavigationError
'
' DESIGN PRINCIPLES
'   - Columns A:B are reserved exclusively for the navigation pane
'   - The toggle is positioned over C1 so it remains visible when A:B are hidden
'   - Navigation is generated from the current list of visible worksheets
'   - Each worksheet permanently highlights its own navigation entry
'   - Normal sheet activation does not rebuild hyperlinks or reapply formatting
'   - Full rebuilds occur only when the visible worksheet structure changes
'   - Shape creation is idempotent and existing toggle shapes are reused
'   - Public entry points preserve and restore Excel application state
'   - Failures on one worksheet do not prevent other worksheets from installing
'
' ERROR POLICY
'   - Public entry points restore application state before returning
'   - Worksheet-specific installation errors are isolated and logged
'   - Diagnostic messages are written to the VBA Immediate Window
'   - No MsgBox is raised
'
' DEPENDENCIES
'   - Microsoft Excel object model
'   - Microsoft Office object model for Shape and Mso constants
'   - ThisWorkbook event procedures call the public procedures in this module
'
' NOTES
'   - The worksheet named _SETUP is excluded from navigation
'   - Only visible worksheets belonging to ThisWorkbook are included
'   - Columns A:B must not contain user data formulas or workbook controls
'   - Cell C1 must remain visible because it anchors the toggle
'   - Navigation expand and collapse are intentionally instantaneous to avoid
'     worksheet redraw flicker
'   - Worksheet renames hides unhides deletions and moves are detected on the
'     next worksheet activation through the cached navigation signature
'
' UPDATED
'   2026-07-12
'==============================================================================

'==============================================================================
' PRIVATE CONSTANTS
'==============================================================================

    'Navigation placement and reserved worksheet area
        Private Const NAV_FIRST_COLUMN              As String = "A"
        Private Const NAV_LAST_COLUMN               As String = "B"
        Private Const NAV_TITLE_ROW                 As Long = 1
        Private Const NAV_HEADER_ROW                As Long = 2
        Private Const NAV_FIRST_ITEM_ROW            As Long = 3
        Private Const NAV_MIN_LAST_ROW              As Long = 11

    'Navigation content
        Private Const NAV_TITLE_TEXT                As String = "PROBABILITY LIBRARY"
        Private Const NAV_HEADER_TEXT               As String = "NAVIGATION"
        Private Const NAV_LINK_TARGET_CELL          As String = "D1"
        Private Const SETUP_SHEET_NAME              As String = "SETUP"

    'Navigation dimensions
        Private Const NAV_ICON_WIDTH                As Double = 5#
        Private Const NAV_LABEL_WIDTH               As Double = 20#

    'Toggle configuration
        Private Const TOGGLE_ANCHOR_CELL            As String = "C1"
        Private Const TOGGLE_SHAPE_NAME             As String = "btnNavigationToggle"
        Private Const TOGGLE_WIDTH                  As Double = 16#
        Private Const TOGGLE_HEIGHT                 As Double = 16#
        Private Const TOGGLE_LEFT_OFFSET            As Double = 1#
        Private Const TOGGLE_TOP_OFFSET             As Double = 7#
        Private Const TOGGLE_FONT_SIZE              As Double = 9#
        Private Const TOGGLE_POSITION_TOLERANCE     As Double = 0.25

    'Interface colours stored as VBA Long values
        Private Const COLOR_NAVY                    As Long = 5190166
        Private Const COLOR_BLUE                    As Long = 11892015
        Private Const COLOR_PALE                    As Long = 16315114
        Private Const COLOR_ACTIVE                  As Long = 16313046
        Private Const COLOR_TEXT                    As Long = 4666404
        Private Const COLOR_WHITE                   As Long = 16777215
        Private Const COLOR_BORDER                  As Long = 15393497

'==============================================================================
' PRIVATE STATE
'==============================================================================
    'Stores the visible worksheet structure used to build the navigation panes
        Private CachedNavigationSignature           As String
    'Prevents overlapping navigation operations caused by repeated clicks
        Private NavigationBusy                      As Boolean


Public Sub InstallNavigation()
'
'==============================================================================
' InstallNavigation
'------------------------------------------------------------------------------
' PURPOSE
'   Installs or refreshes navigation on every eligible worksheet
'
' WHY THIS EXISTS
'   Every user-facing worksheet must contain a navigation pane synchronized with
'   the current workbook structure and a correctly configured toggle button
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Preserves the current screen-updating and event settings
'   - Prevents overlapping navigation operations
'   - Processes each eligible worksheet independently
'   - Continues when one worksheet cannot be processed
'   - Caches the worksheet structure used to build the navigation panes
'
' ERROR POLICY
'   - Worksheet-specific errors are logged and isolated
'   - Procedure-level errors are written to the VBA Immediate Window
'   - The original Excel application state is always restored
'   - No MsgBox is raised
'
' DEPENDENCIES
'   - IsNavigationSheet
'   - InstallNavigationOnSheet
'   - GetNavigationSignature
'   - ReportNavigationError
'
' CALLED FROM
'   - Workbook_Open
'   - RefreshAllNavigation
'   - RefreshActiveSheet when the worksheet structure changes
'   - Manual execution
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Ws                          As Worksheet

    Dim FailedWorksheetCount        As Long
    Dim ErrorNumber                 As Long

    Dim ErrorDescription            As String

    Dim PreviousScreenUpdating      As Boolean
    Dim PreviousEnableEvents        As Boolean

'------------------------------------------------------------------------------
' VALIDATE EXECUTION STATE
'------------------------------------------------------------------------------
    'Ignore overlapping installation or toggle requests
        If NavigationBusy Then Exit Sub

'------------------------------------------------------------------------------
' PRESERVE APPLICATION STATE
'------------------------------------------------------------------------------
    'Store the current application settings before changing them
        PreviousScreenUpdating = Application.ScreenUpdating
        PreviousEnableEvents = Application.EnableEvents

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route procedure-level runtime errors through the common error handler
        On Error GoTo Err_Handler

    'Mark the navigation engine as busy
        NavigationBusy = True

'------------------------------------------------------------------------------
' OPTIMIZE EXECUTION
'------------------------------------------------------------------------------
    'Suppress screen redraws while rebuilding worksheet navigation
        Application.ScreenUpdating = False

    'Prevent workbook events from firing during installation
        Application.EnableEvents = False

'------------------------------------------------------------------------------
' INSTALL NAVIGATION
'------------------------------------------------------------------------------
    For Each Ws In ThisWorkbook.Worksheets

        'Process only worksheets eligible for the navigation interface
            If IsNavigationSheet(Ws) Then

                'Install each worksheet independently so one failure does not
                'prevent subsequent worksheets from being processed
                    If Not InstallNavigationOnSheet(Ws) Then
                        FailedWorksheetCount = FailedWorksheetCount + 1
                    End If

            End If

    Next Ws

'------------------------------------------------------------------------------
' CACHE WORKBOOK STRUCTURE
'------------------------------------------------------------------------------
    'Store the worksheet structure used to build the navigation panes
        CachedNavigationSignature = GetNavigationSignature()

'------------------------------------------------------------------------------
' REPORT PARTIAL INSTALLATION
'------------------------------------------------------------------------------
    'Write a summary when one or more worksheets could not be processed
        If FailedWorksheetCount > 0 Then

            Debug.Print _
                Format$(Now, "yyyy-mm-dd hh:nn:ss") & _
                " | M_NAVIGATION.InstallNavigation" & _
                " | Failed worksheets: " & _
                CStr(FailedWorksheetCount)

        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
Clean_Exit:
    'Restore the event state captured on entry
        Application.EnableEvents = PreviousEnableEvents

    'Restore the screen-updating state captured on entry
        Application.ScreenUpdating = PreviousScreenUpdating

    'Release the navigation re-entry guard
        NavigationBusy = False

    'Exit after successful execution or cleanup
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Capture the original runtime error before calling another procedure
        ErrorNumber = Err.Number
        ErrorDescription = Err.Description

    'Write the procedure-level failure to the Immediate Window
        ReportNavigationError _
            ProcedureName:="InstallNavigation", _
            ErrorNumber:=ErrorNumber, _
            ErrorDescription:=ErrorDescription

    'Restore the original application state before returning
        Resume Clean_Exit

End Sub
Public Sub ToggleNavigation()
'
'==============================================================================
' ToggleNavigation
'------------------------------------------------------------------------------
' PURPOSE
'   Expands or collapses the navigation pane on the active worksheet
'
' WHY THIS EXISTS
'   Users require a persistent control that can release worksheet space without
'   removing the navigation interface from the workbook
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Validates that the active sheet belongs to ThisWorkbook
'   - Hides or restores the reserved navigation columns in one operation
'   - Repositions the toggle after the worksheet layout changes
'   - Synchronizes the arrow with the resulting navigation state
'
' ERROR POLICY
'   - Restores the original application state after success or failure
'   - Writes unexpected runtime errors to the VBA Immediate Window
'   - No MsgBox is raised
'
' DEPENDENCIES
'   - IsNavigationSheet
'   - IsNavigationCollapsed
'   - CollapseNavigation
'   - ExpandNavigation
'   - PositionToggleButton
'   - UpdateToggleAppearance
'   - SafeWorksheetName
'   - ReportNavigationError
'
' CALLED FROM
'   - Worksheet toggle shapes through OnAction
'   - Manual execution
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Ws                          As Worksheet

    Dim ErrorNumber                 As Long

    Dim ErrorDescription            As String
    Dim WorksheetName               As String

    Dim PreviousScreenUpdating      As Boolean
    Dim PreviousEnableEvents        As Boolean

'------------------------------------------------------------------------------
' VALIDATE EXECUTION STATE
'------------------------------------------------------------------------------
    'Ignore overlapping installation or toggle requests
        If NavigationBusy Then Exit Sub

    'Exit when the active object is not a worksheet
        If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub

    'Capture the active worksheet
        Set Ws = ActiveSheet

    'Exit when the active worksheet is not eligible for navigation
        If Not IsNavigationSheet(Ws) Then Exit Sub

'------------------------------------------------------------------------------
' PRESERVE APPLICATION STATE
'------------------------------------------------------------------------------
    'Store the current application settings before changing them
        PreviousScreenUpdating = Application.ScreenUpdating
        PreviousEnableEvents = Application.EnableEvents

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors through the common error handler
        On Error GoTo Err_Handler

    'Mark the navigation engine as busy
        NavigationBusy = True

'------------------------------------------------------------------------------
' OPTIMIZE EXECUTION
'------------------------------------------------------------------------------
    'Prevent intermediate worksheet redraws
        Application.ScreenUpdating = False

    'Prevent workbook events during the layout change
        Application.EnableEvents = False

'------------------------------------------------------------------------------
' TOGGLE NAVIGATION
'------------------------------------------------------------------------------
    If IsNavigationCollapsed(Ws) Then

        'Restore the navigation pane
            ExpandNavigation Ws

    Else

        'Hide the navigation pane
            CollapseNavigation Ws

    End If

'------------------------------------------------------------------------------
' REFRESH TOGGLE
'------------------------------------------------------------------------------
    'Reposition the toggle after the final worksheet layout is established
        PositionToggleButton Ws

    'Synchronize the displayed arrow with the navigation state
        UpdateToggleAppearance Ws

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
Clean_Exit:
    'Restore the event state captured on entry
        Application.EnableEvents = PreviousEnableEvents

    'Restore screen updating after all visual changes are complete
        Application.ScreenUpdating = PreviousScreenUpdating

    'Release the navigation re-entry guard
        NavigationBusy = False

    'Exit after successful execution or cleanup
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Capture the original runtime error before diagnostic processing
        ErrorNumber = Err.Number
        ErrorDescription = Err.Description
        WorksheetName = SafeWorksheetName(Ws)

    'Write the unexpected toggle failure to the Immediate Window
        ReportNavigationError _
            ProcedureName:="ToggleNavigation", _
            ErrorNumber:=ErrorNumber, _
            ErrorDescription:=ErrorDescription, _
            WorksheetName:=WorksheetName

    'Restore the original application state before returning
        Resume Clean_Exit

End Sub
Public Sub RefreshActiveSheet(ByVal TargetSheet As Worksheet)
'
'==============================================================================
' RefreshActiveSheet
'------------------------------------------------------------------------------
' PURPOSE
'   Refreshes the navigation state required when a worksheet becomes active
'
' WHY THIS EXISTS
'   Normal worksheet activation should not rebuild hyperlinks formatting or
'   borders because each worksheet already highlights its own navigation entry
'
'   A complete rebuild is required only when the visible worksheet structure has
'   changed since the previous installation
'
' INPUTS
'   TargetSheet
'     Worksheet that has just become active
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Validates the target worksheet
'   - Compares the current worksheet signature with the cached signature
'   - Reinstalls navigation only when the worksheet structure changed
'   - Recreates a missing pane or toggle only on the active worksheet
'   - Avoids cell-format updates during normal worksheet activation
'
' ERROR POLICY
'   - Restores the original application state after success or failure
'   - Writes unexpected runtime errors to the VBA Immediate Window
'   - No MsgBox is raised
'
' DEPENDENCIES
'   - IsNavigationSheet
'   - GetNavigationSignature
'   - InstallNavigation
'   - HasNavigationPane
'   - InstallNavigationOnSheet
'   - GetToggleShape
'   - EnsureToggleButton
'   - PositionToggleButton
'   - UpdateToggleAppearance
'   - SafeWorksheetName
'   - ReportNavigationError
'
' CALLED FROM
'   - Workbook_SheetActivate
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ToggleShape                 As Shape

    Dim ErrorNumber                 As Long

    Dim CurrentNavigationSignature  As String
    Dim ErrorDescription            As String
    Dim WorksheetName               As String

    Dim PreviousScreenUpdating      As Boolean
    Dim PreviousEnableEvents        As Boolean

'------------------------------------------------------------------------------
' VALIDATE EXECUTION STATE
'------------------------------------------------------------------------------
    'Ignore refresh requests raised by an active navigation operation
        If NavigationBusy Then Exit Sub

    'Exit when the target worksheet is not eligible for navigation
        If Not IsNavigationSheet(TargetSheet) Then Exit Sub

'------------------------------------------------------------------------------
' PRESERVE APPLICATION STATE
'------------------------------------------------------------------------------
    'Store the current application settings before any refresh work
        PreviousScreenUpdating = Application.ScreenUpdating
        PreviousEnableEvents = Application.EnableEvents

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors through the common error handler
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' DETECT WORKBOOK-STRUCTURE CHANGES
'------------------------------------------------------------------------------
    'Generate the current visible-worksheet signature
        CurrentNavigationSignature = GetNavigationSignature()

    'Rebuild every navigation pane only when the worksheet structure changed
        If StrComp( _
            CurrentNavigationSignature, _
            CachedNavigationSignature, _
            vbBinaryCompare) <> 0 Then

            InstallNavigation
            GoTo Clean_Exit

        End If

'------------------------------------------------------------------------------
' LOCK REFRESH
'------------------------------------------------------------------------------
    'Mark the navigation engine as busy for the local visual refresh
        NavigationBusy = True

'------------------------------------------------------------------------------
' OPTIMIZE EXECUTION
'------------------------------------------------------------------------------
    'Prevent visual refresh work from producing sheet-activation flicker
        Application.ScreenUpdating = False

    'Prevent workbook events from firing during the refresh
        Application.EnableEvents = False

'------------------------------------------------------------------------------
' VERIFY NAVIGATION PANE
'------------------------------------------------------------------------------
    'Rebuild only the active worksheet when its generated pane is missing
        If Not HasNavigationPane(TargetSheet) Then

            If Not InstallNavigationOnSheet(TargetSheet) Then
                GoTo Clean_Exit
            End If

        End If

'------------------------------------------------------------------------------
' VERIFY TOGGLE BUTTON
'------------------------------------------------------------------------------
    'Retrieve the existing toggle without raising an expected lookup error
        Set ToggleShape = GetToggleShape(TargetSheet)

    If ToggleShape Is Nothing Then

        'Recreate the toggle if it was manually deleted
            EnsureToggleButton TargetSheet

    Else

        'Correct the position only when it differs from the configured anchor
            PositionToggleButton TargetSheet

        'Synchronize the arrow with the current navigation state
            UpdateToggleAppearance TargetSheet

    End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
Clean_Exit:
    'Restore the event state captured on entry
        Application.EnableEvents = PreviousEnableEvents

    'Restore the screen-updating state captured on entry
        Application.ScreenUpdating = PreviousScreenUpdating

    'Release the navigation re-entry guard
        NavigationBusy = False

    'Exit after successful execution or cleanup
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Capture the original runtime error before diagnostic processing
        ErrorNumber = Err.Number
        ErrorDescription = Err.Description
        WorksheetName = SafeWorksheetName(TargetSheet)

    'Write the unexpected refresh failure to the Immediate Window
        ReportNavigationError _
            ProcedureName:="RefreshActiveSheet", _
            ErrorNumber:=ErrorNumber, _
            ErrorDescription:=ErrorDescription, _
            WorksheetName:=WorksheetName

    'Restore the original application state before returning
        Resume Clean_Exit

End Sub
Public Sub RefreshAllNavigation()
'
'==============================================================================
' RefreshAllNavigation
'------------------------------------------------------------------------------
' PURPOSE
'   Rebuilds navigation on all eligible worksheets
'
' WHY THIS EXISTS
'   Workbook-structure changes require every sidebar to be synchronized with the
'   current visible worksheet list
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Delegates the complete refresh to InstallNavigation
'
' ERROR POLICY
'   Delegates application-state restoration and error handling to
'   InstallNavigation
'
' DEPENDENCIES
'   - InstallNavigation
'
' CALLED FROM
'   - Workbook_NewSheet
'   - Manual execution
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' REFRESH NAVIGATION
'------------------------------------------------------------------------------
    'Run the centralized installation and refresh routine
        InstallNavigation

End Sub


Private Function InstallNavigationOnSheet( _
    ByVal Ws As Worksheet) _
    As Boolean
'
'==============================================================================
' InstallNavigationOnSheet
'------------------------------------------------------------------------------
' PURPOSE
'   Installs or refreshes navigation on one eligible worksheet
'
' WHY THIS EXISTS
'   A failure on one worksheet must not prevent navigation from being installed
'   on the remaining eligible worksheets
'
' INPUTS
'   Ws
'     Worksheet receiving the navigation interface
'
' RETURNS
'   Boolean
'     True  => navigation was installed successfully
'     False => the worksheet was ineligible or installation failed
'
' BEHAVIOR
'   - Validates worksheet eligibility
'   - Rebuilds the worksheet navigation pane
'   - Creates or refreshes the worksheet toggle
'   - Contains failures within the current worksheet
'
' ERROR POLICY
'   - Unexpected runtime errors are written to the VBA Immediate Window
'   - No MsgBox is raised
'
' DEPENDENCIES
'   - IsNavigationSheet
'   - BuildNavigation
'   - EnsureToggleButton
'   - SafeWorksheetName
'   - ReportNavigationError
'
' CALLED FROM
'   - InstallNavigation
'   - RefreshActiveSheet
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrorNumber                 As Long

    Dim ErrorDescription            As String
    Dim WorksheetName               As String

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Return False unless the complete installation succeeds
        InstallNavigationOnSheet = False

    'Route unexpected runtime errors to the local error handler
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' VALIDATE WORKSHEET
'------------------------------------------------------------------------------
    'Reject worksheets that are not eligible for navigation
        If Not IsNavigationSheet(Ws) Then Exit Function

'------------------------------------------------------------------------------
' INSTALL NAVIGATION
'------------------------------------------------------------------------------
    'Rebuild the worksheet navigation pane
        BuildNavigation Ws

    'Create or refresh the worksheet toggle button
        EnsureToggleButton Ws

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Confirm that the complete worksheet installation succeeded
        InstallNavigationOnSheet = True

    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Capture the original runtime error before diagnostic processing
        ErrorNumber = Err.Number
        ErrorDescription = Err.Description
        WorksheetName = SafeWorksheetName(Ws)

    'Write the worksheet-specific failure to the Immediate Window
        ReportNavigationError _
            ProcedureName:="InstallNavigationOnSheet", _
            ErrorNumber:=ErrorNumber, _
            ErrorDescription:=ErrorDescription, _
            WorksheetName:=WorksheetName

End Function
Private Sub BuildNavigation(ByVal Ws As Worksheet)
'
'==============================================================================
' BuildNavigation
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the current visible worksheet list into the reserved navigation area
'
' WHY THIS EXISTS
'   Each eligible worksheet requires a locally rendered navigation pane whose
'   hyperlinks and highlighted entry reflect the current workbook structure
'
' INPUTS
'   Ws
'     Worksheet receiving the navigation pane
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Preserves the expanded or collapsed state
'   - Clears only the generated navigation area
'   - Creates one internal hyperlink for each eligible worksheet
'   - Applies default formatting in bulk
'   - Highlights the entry representing Ws
'   - Restores the configured navigation widths when expanded
'
' ERROR POLICY
'   Runtime errors propagate to InstallNavigationOnSheet
'
' DEPENDENCIES
'   - IsNavigationCollapsed
'   - ClearNavigationArea
'   - IsNavigationSheet
'   - NavigationIcon
'
' CALLED FROM
'   - InstallNavigationOnSheet
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SourceSheet                 As Worksheet

    Dim NavigationRow               As Long
    Dim ActualLastRow               As Long
    Dim FormattedLastRow            As Long
    Dim ActiveNavigationRow         As Long

    Dim TargetAddress               As String

    Dim NavigationWasCollapsed      As Boolean

'------------------------------------------------------------------------------
' PRESERVE NAVIGATION STATE
'------------------------------------------------------------------------------
    'Capture the current pane state before rebuilding the reserved area
        NavigationWasCollapsed = IsNavigationCollapsed(Ws)

'------------------------------------------------------------------------------
' CLEAR EXISTING NAVIGATION
'------------------------------------------------------------------------------
    'Remove values hyperlinks formats and merges generated by this module
        ClearNavigationArea Ws

'------------------------------------------------------------------------------
' BUILD TITLE
'------------------------------------------------------------------------------
    With Ws
        'Merge the title cells and write the configured title
            .Range( _
                NAV_FIRST_COLUMN & CStr(NAV_TITLE_ROW) & ":" & _
                NAV_LAST_COLUMN & CStr(NAV_TITLE_ROW)).Merge
            .Range(NAV_FIRST_COLUMN & CStr(NAV_TITLE_ROW)).Value2 = _
                NAV_TITLE_TEXT
        'Apply the title formatting
            With .Range( _
                NAV_FIRST_COLUMN & CStr(NAV_TITLE_ROW) & ":" & _
                NAV_LAST_COLUMN & CStr(NAV_TITLE_ROW))
                .Interior.Color = vbBlack
                .Font.Bold = True
                .Font.Color = VBA.RGB(255, 192, 0)
                .Font.Size = 12
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With

'------------------------------------------------------------------------------
' BUILD SECTION HEADER
'------------------------------------------------------------------------------
        'Merge the section-header cells and write the configured label
            .Range( _
                NAV_FIRST_COLUMN & CStr(NAV_HEADER_ROW) & ":" & _
                NAV_LAST_COLUMN & CStr(NAV_HEADER_ROW)).Merge
            .Range(NAV_FIRST_COLUMN & CStr(NAV_HEADER_ROW)).Value2 = _
                NAV_HEADER_TEXT

        'Apply the section-header formatting
            With .Range( _
                NAV_FIRST_COLUMN & CStr(NAV_HEADER_ROW) & ":" & _
                NAV_LAST_COLUMN & CStr(NAV_HEADER_ROW))
                .Interior.Color = VBA.RGB(15, 31, 54)
                .Font.Bold = True
                .Font.Color = COLOR_WHITE
                .Font.Size = 10
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .IndentLevel = 1
            End With

'------------------------------------------------------------------------------
' BUILD WORKSHEET LINKS
'------------------------------------------------------------------------------
        'Start the worksheet list immediately below the section header
            NavigationRow = NAV_FIRST_ITEM_ROW

        For Each SourceSheet In ThisWorkbook.Worksheets

            'Create one navigation entry for each eligible worksheet
                If IsNavigationSheet(SourceSheet) Then

                    'Write the compact worksheet icon
                        .Cells(NavigationRow, .Columns(NAV_FIRST_COLUMN).Column).Value2 = _
                            NavigationIcon(SourceSheet.Name)

                    'Build an escaped internal hyperlink destination
                        TargetAddress = "'" & _
                                        Replace(SourceSheet.Name, "'", "''") & _
                                        "'!" & _
                                        NAV_LINK_TARGET_CELL

                    'Create the internal worksheet hyperlink and display name
                        .Hyperlinks.Add _
                            Anchor:=.Cells( _
                                NavigationRow, _
                                .Columns(NAV_LAST_COLUMN).Column), _
                            Address:=vbNullString, _
                            SubAddress:=TargetAddress, _
                            TextToDisplay:=SourceSheet.Name

                    'Record the row representing the worksheet being built
                        If StrComp( _
                            SourceSheet.Name, _
                            Ws.Name, _
                            vbBinaryCompare) = 0 Then

                            ActiveNavigationRow = NavigationRow

                        End If

                    'Advance to the next navigation row
                        NavigationRow = NavigationRow + 1

                End If

        Next SourceSheet

'------------------------------------------------------------------------------
' FORMAT WORKSHEET LINKS
'------------------------------------------------------------------------------
        'Identify the last row containing a generated worksheet link
            ActualLastRow = NavigationRow - 1

        If ActualLastRow >= NAV_FIRST_ITEM_ROW Then

            'Apply default formatting to all navigation entries in one operation
                With .Range( _
                    .Cells( _
                        NAV_FIRST_ITEM_ROW, _
                        .Columns(NAV_FIRST_COLUMN).Column), _
                    .Cells( _
                        ActualLastRow, _
                        .Columns(NAV_LAST_COLUMN).Column))

                    .Interior.Color = COLOR_PALE
                    .Font.Bold = False
                    .Font.Color = COLOR_TEXT
                    .Font.Size = 10
                    .VerticalAlignment = xlCenter

                End With

            'Apply icon-column formatting in one operation
                With .Range( _
                    .Cells( _
                        NAV_FIRST_ITEM_ROW, _
                        .Columns(NAV_FIRST_COLUMN).Column), _
                    .Cells( _
                        ActualLastRow, _
                        .Columns(NAV_FIRST_COLUMN).Column))

                    .Font.Bold = True
                    .Font.Color = COLOR_BLUE
                    .HorizontalAlignment = xlCenter

                End With

            'Apply label-column formatting in one operation
                With .Range( _
                    .Cells( _
                        NAV_FIRST_ITEM_ROW, _
                        .Columns(NAV_LAST_COLUMN).Column), _
                    .Cells( _
                        ActualLastRow, _
                        .Columns(NAV_LAST_COLUMN).Column))

                    .HorizontalAlignment = xlLeft
                    .Font.Underline = xlUnderlineStyleNone

                End With

            'Highlight the navigation entry representing this worksheet
                If ActiveNavigationRow >= NAV_FIRST_ITEM_ROW Then

                    With .Range( _
                        .Cells( _
                            ActiveNavigationRow, _
                            .Columns(NAV_FIRST_COLUMN).Column), _
                        .Cells( _
                            ActiveNavigationRow, _
                            .Columns(NAV_LAST_COLUMN).Column))

                        .Interior.Color = COLOR_ACTIVE
                        .Font.Bold = True
                        .Font.Color = COLOR_NAVY

                    End With

                End If

        End If

'------------------------------------------------------------------------------
' FINALIZE NAVIGATION LAYOUT
'------------------------------------------------------------------------------
        'Retain a minimum formatted navigation depth
            FormattedLastRow = Application.Max( _
                NAV_MIN_LAST_ROW, _
                ActualLastRow)

        'Apply light borders to the complete navigation area
            With .Range( _
                NAV_FIRST_COLUMN & CStr(NAV_TITLE_ROW) & ":" & _
                NAV_LAST_COLUMN & CStr(FormattedLastRow)).Borders

                .LineStyle = xlContinuous
                .Color = COLOR_BORDER
                .Weight = xlHairline

            End With

        'Restore expanded widths only when the pane was previously visible
            If NavigationWasCollapsed Then

                .Columns( _
                    NAV_FIRST_COLUMN & ":" & _
                    NAV_LAST_COLUMN).Hidden = True

            Else

                .Columns(NAV_FIRST_COLUMN).Hidden = False
                .Columns(NAV_LAST_COLUMN).Hidden = False

                .Columns(NAV_FIRST_COLUMN).ColumnWidth = NAV_ICON_WIDTH
                .Columns(NAV_LAST_COLUMN).ColumnWidth = NAV_LABEL_WIDTH

            End If

    End With

End Sub


Private Sub ClearNavigationArea(ByVal Ws As Worksheet)
'
'==============================================================================
' ClearNavigationArea
'------------------------------------------------------------------------------
' PURPOSE
'   Removes the navigation content and formatting generated by this module
'
' WHY THIS EXISTS
'   A rebuild must remove stale links borders formats and merged headers without
'   clearing an arbitrary fixed range such as A1:B200
'
' INPUTS
'   Ws
'     Worksheet containing the generated navigation pane
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Identifies the existing generated navigation depth
'   - Unmerges the generated title and section-header rows
'   - Deletes existing hyperlinks from the generated area
'   - Clears generated contents formats and borders
'
' ERROR POLICY
'   Runtime errors propagate to InstallNavigationOnSheet
'
' DEPENDENCIES
'   - GetExistingNavigationLastRow
'
' CALLED FROM
'   - BuildNavigation
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ClearLastRow                As Long
    Dim NavigationRange             As Range

'------------------------------------------------------------------------------
' IDENTIFY NAVIGATION AREA
'------------------------------------------------------------------------------
    'Clear at least the minimum formatted navigation depth
        ClearLastRow = Application.Max( _
            NAV_MIN_LAST_ROW, _
            GetExistingNavigationLastRow(Ws))

    'Capture the complete generated navigation area
        Set NavigationRange = Ws.Range( _
            NAV_FIRST_COLUMN & CStr(NAV_TITLE_ROW) & ":" & _
            NAV_LAST_COLUMN & CStr(ClearLastRow))

'------------------------------------------------------------------------------
' REMOVE GENERATED MERGES
'------------------------------------------------------------------------------
    'Unmerge only the two header rows controlled by this module
        Ws.Range( _
            NAV_FIRST_COLUMN & CStr(NAV_TITLE_ROW) & ":" & _
            NAV_LAST_COLUMN & CStr(NAV_HEADER_ROW)).UnMerge

'------------------------------------------------------------------------------
' REMOVE GENERATED HYPERLINKS
'------------------------------------------------------------------------------
    'Delete hyperlinks without failing when the range contains none
        On Error Resume Next
        NavigationRange.Hyperlinks.Delete
        On Error GoTo 0

'------------------------------------------------------------------------------
' CLEAR GENERATED CONTENT AND FORMATTING
'------------------------------------------------------------------------------
    'Remove generated values and formulas
        NavigationRange.ClearContents

    'Remove generated cell formatting
        NavigationRange.ClearFormats

    'Remove generated borders explicitly
        NavigationRange.Borders.LineStyle = xlNone

End Sub


Private Sub EnsureToggleButton(ByVal Ws As Worksheet)
'
'==============================================================================
' EnsureToggleButton
'------------------------------------------------------------------------------
' PURPOSE
'   Creates or refreshes the rounded navigation toggle on a worksheet
'
' WHY THIS EXISTS
'   Every eligible worksheet requires a persistent control that remains
'   accessible whether the navigation pane is expanded or collapsed
'
' INPUTS
'   Ws
'     Worksheet receiving or refreshing the toggle button
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Reuses the existing named shape when available
'   - Creates a rounded rectangle when the shape does not exist
'   - Applies configured size colour typography placement and macro action
'   - Positions the button and synchronizes its arrow
'
' ERROR POLICY
'   Runtime errors propagate to InstallNavigationOnSheet
'
' DEPENDENCIES
'   - GetToggleShape
'   - PositionToggleButton
'   - UpdateToggleAppearance
'
' CALLED FROM
'   - InstallNavigationOnSheet
'   - RefreshActiveSheet when the toggle is missing
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ToggleShape                 As Shape
    Dim EscapedWorkbookName         As String

'------------------------------------------------------------------------------
' FIND OR CREATE TOGGLE
'------------------------------------------------------------------------------
    'Retrieve the existing named toggle shape when available
        Set ToggleShape = GetToggleShape(Ws)

    'Create the rounded toggle button when it does not already exist
        If ToggleShape Is Nothing Then

            Set ToggleShape = Ws.Shapes.AddShape( _
                Type:=msoShapeRoundedRectangle, _
                Left:=Ws.Range(TOGGLE_ANCHOR_CELL).Left + TOGGLE_LEFT_OFFSET, _
                Top:=Ws.Range(TOGGLE_ANCHOR_CELL).Top + TOGGLE_TOP_OFFSET, _
                Width:=TOGGLE_WIDTH, _
                Height:=TOGGLE_HEIGHT)

            ToggleShape.Name = TOGGLE_SHAPE_NAME

        End If

'------------------------------------------------------------------------------
' BUILD MACRO ACTION
'------------------------------------------------------------------------------
    'Escape apostrophes in the workbook name used by the shape action
        EscapedWorkbookName = Replace(ThisWorkbook.Name, "'", "''")

'------------------------------------------------------------------------------
' FORMAT TOGGLE
'------------------------------------------------------------------------------
    With ToggleShape

        'Keep the configured dimensions independent from worksheet zoom
            .LockAspectRatio = msoFalse
            .Width = TOGGLE_WIDTH
            .Height = TOGGLE_HEIGHT

        'Move with the anchor cell without resizing with the cell
            .Placement = xlMove

        'Assign the fully qualified navigation macro
            .OnAction = "'" & _
                        EscapedWorkbookName & _
                        "'!M_NAVIGATION.ToggleNavigation"

        'Apply fill and outline formatting
            .Fill.ForeColor.RGB = COLOR_NAVY
            .Line.ForeColor.RGB = COLOR_NAVY
            .Line.Weight = 1

'------------------------------------------------------------------------------
' FORMAT TOGGLE
'------------------------------------------------------------------------------
    With ToggleShape

        'Keep the configured dimensions independent from worksheet zoom
            .LockAspectRatio = msoFalse
            .Width = TOGGLE_WIDTH
            .Height = TOGGLE_HEIGHT

        'Move with the anchor cell without resizing with the cell
            .Placement = xlMove

        'Assign the fully qualified navigation macro
            .OnAction = "'" & _
                        EscapedWorkbookName & _
                        "'!M_NAVIGATION.ToggleNavigation"

        'Apply fill and outline formatting
            .Fill.ForeColor.RGB = COLOR_NAVY
            .Line.ForeColor.RGB = COLOR_NAVY
            .Line.Weight = 1

        'Configure text-frame sizing and internal margins
            .TextFrame2.AutoSize = msoAutoSizeNone
            .TextFrame2.MarginLeft = 0
            .TextFrame2.MarginRight = 0
            .TextFrame2.MarginTop = 0
            .TextFrame2.MarginBottom = 0
            .TextFrame2.VerticalAnchor = msoAnchorMiddle

        'Configure arrow typography
            With .TextFrame2.TextRange

                .Font.Name = "Aptos"
                .Font.Size = TOGGLE_FONT_SIZE
                .Font.Bold = msoTrue
                .Font.Fill.ForeColor.RGB = COLOR_WHITE
                .ParagraphFormat.Alignment = msoAlignCenter

            End With

        'Provide an accessibility description
            .AlternativeText = _
                "Expand or collapse worksheet navigation"

    End With




        'Configure text-frame sizing and internal margins
            .TextFrame2.AutoSize = msoAutoSizeNone
            .TextFrame2.MarginLeft = 0
            .TextFrame2.MarginRight = 0
            .TextFrame2.MarginTop = 0
            .TextFrame2.MarginBottom = 0
            .TextFrame2.VerticalAnchor = msoAnchorMiddle

        'Configure arrow typography
            With .TextFrame2.TextRange

                .Font.Name = "Aptos"
                .Font.Size = TOGGLE_FONT_SIZE
                .Font.Bold = msoTrue
                .Font.Fill.ForeColor.RGB = COLOR_WHITE
                .ParagraphFormat.Alignment = msoAlignCenter

            End With

        'Provide an accessibility description
            .AlternativeText = "Expand or collapse worksheet navigation"

    End With

'------------------------------------------------------------------------------
' FINALIZE TOGGLE
'------------------------------------------------------------------------------
    'Position the toggle and bring it above worksheet content
        PositionToggleButton Ws, True

    'Synchronize the arrow with the current navigation state
        UpdateToggleAppearance Ws

End Sub


Private Sub PositionToggleButton( _
    ByVal Ws As Worksheet, _
    Optional ByVal BringToFront As Boolean = False)
'
'==============================================================================
' PositionToggleButton
'------------------------------------------------------------------------------
' PURPOSE
'   Aligns the toggle with its configured worksheet anchor
'
' WHY THIS EXISTS
'   Hiding or restoring navigation columns changes the physical position of C1
'   and Excel does not always redraw a shape at the expected coordinate
'
' INPUTS
'   Ws
'     Worksheet containing the toggle
'
'   BringToFront
'     Optional flag that raises the toggle above other worksheet objects
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Retrieves the existing toggle
'   - Calculates the expected left and top coordinates
'   - Writes coordinates only when they differ beyond a small tolerance
'   - Brings the toggle to the front only when explicitly requested
'
' ERROR POLICY
'   Missing-shape lookup errors are handled by GetToggleShape
'   Other runtime errors propagate to the calling procedure
'
' DEPENDENCIES
'   - GetToggleShape
'
' CALLED FROM
'   - ToggleNavigation
'   - RefreshActiveSheet
'   - EnsureToggleButton
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ToggleShape                 As Shape
    Dim AnchorRange                 As Range

    Dim ExpectedLeft                As Double
    Dim ExpectedTop                 As Double

'------------------------------------------------------------------------------
' FIND TOGGLE
'------------------------------------------------------------------------------
    'Retrieve the existing named toggle shape
        Set ToggleShape = GetToggleShape(Ws)

    'Exit when the toggle does not exist
        If ToggleShape Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' CALCULATE POSITION
'------------------------------------------------------------------------------
    'Capture the configured anchor cell
        Set AnchorRange = Ws.Range(TOGGLE_ANCHOR_CELL)

    'Calculate the required shape coordinates
        ExpectedLeft = AnchorRange.Left + TOGGLE_LEFT_OFFSET
        ExpectedTop = AnchorRange.Top + TOGGLE_TOP_OFFSET

'------------------------------------------------------------------------------
' APPLY POSITION
'------------------------------------------------------------------------------
    With ToggleShape

        'Correct the horizontal position only when required
            If Abs(.Left - ExpectedLeft) > TOGGLE_POSITION_TOLERANCE Then
                .Left = ExpectedLeft
            End If

        'Correct the vertical position only when required
            If Abs(.Top - ExpectedTop) > TOGGLE_POSITION_TOLERANCE Then
                .Top = ExpectedTop
            End If

        'Raise the toggle only during installation or recreation
            If BringToFront Then
                .ZOrder msoBringToFront
            End If

    End With

End Sub


Private Sub UpdateToggleAppearance(ByVal Ws As Worksheet)
'
'==============================================================================
' UpdateToggleAppearance
'------------------------------------------------------------------------------
' PURPOSE
'   Synchronizes the toggle arrow with the navigation-pane state
'
' WHY THIS EXISTS
'   The button must communicate the action available to the user by displaying a
'   right arrow for expansion and a left arrow for collapse
'
' INPUTS
'   Ws
'     Worksheet containing the toggle
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Retrieves the named toggle
'   - Determines the required arrow from the pane state
'   - Updates the shape text only when the arrow changed
'
' ERROR POLICY
'   Missing-shape lookup errors are handled by GetToggleShape
'   Other runtime errors propagate to the calling procedure
'
' DEPENDENCIES
'   - GetToggleShape
'   - IsNavigationCollapsed
'
' CALLED FROM
'   - ToggleNavigation
'   - RefreshActiveSheet
'   - EnsureToggleButton
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ToggleShape                 As Shape

    Dim RequiredArrow               As String

'------------------------------------------------------------------------------
' FIND TOGGLE
'------------------------------------------------------------------------------
    'Retrieve the existing named toggle shape
        Set ToggleShape = GetToggleShape(Ws)

    'Exit when the toggle does not exist
        If ToggleShape Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' DETERMINE REQUIRED ARROW
'------------------------------------------------------------------------------
    'Show the arrow representing the action currently available
        If IsNavigationCollapsed(Ws) Then
            RequiredArrow = ChrW(&H25B6)
        Else
            RequiredArrow = ChrW(&H25C0)
        End If

'------------------------------------------------------------------------------
' UPDATE ARROW
'------------------------------------------------------------------------------
    'Avoid rewriting the shape text when the current arrow is already correct
        If StrComp( _
            ToggleShape.TextFrame2.TextRange.Text, _
            RequiredArrow, _
            vbBinaryCompare) <> 0 Then

            ToggleShape.TextFrame2.TextRange.Text = RequiredArrow

        End If

End Sub


Private Sub CollapseNavigation(ByVal Ws As Worksheet)
'
'==============================================================================
' CollapseNavigation
'------------------------------------------------------------------------------
' PURPOSE
'   Collapses the worksheet navigation pane without animated redraws
'
' WHY THIS EXISTS
'   Repeated column-width changes force Excel to repaint the complete worksheet
'   and cause visible flickering on complex worksheets
'
' INPUTS
'   Ws
'     Worksheet containing the navigation pane
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Hides both reserved navigation columns in one operation
'
' ERROR POLICY
'   Runtime errors propagate to ToggleNavigation
'
' DEPENDENCIES
'   - NAV_FIRST_COLUMN
'   - NAV_LAST_COLUMN
'
' CALLED FROM
'   - ToggleNavigation
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' COLLAPSE NAVIGATION
'------------------------------------------------------------------------------
    'Hide the complete navigation pane in one worksheet-layout operation
        Ws.Columns( _
            NAV_FIRST_COLUMN & ":" & _
            NAV_LAST_COLUMN).Hidden = True

End Sub


Private Sub ExpandNavigation(ByVal Ws As Worksheet)
'
'==============================================================================
' ExpandNavigation
'------------------------------------------------------------------------------
' PURPOSE
'   Expands the worksheet navigation pane without animated redraws
'
' WHY THIS EXISTS
'   Restoring both columns and their configured widths in one operation avoids
'   the repaint loop caused by column-width animation
'
' INPUTS
'   Ws
'     Worksheet containing the navigation pane
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Restores both navigation columns
'   - Restores the configured icon and label widths
'
' ERROR POLICY
'   Runtime errors propagate to ToggleNavigation
'
' DEPENDENCIES
'   - NAV_FIRST_COLUMN
'   - NAV_LAST_COLUMN
'   - NAV_ICON_WIDTH
'   - NAV_LABEL_WIDTH
'
' CALLED FROM
'   - ToggleNavigation
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' EXPAND NAVIGATION
'------------------------------------------------------------------------------
    'Restore both navigation columns
        Ws.Columns( _
            NAV_FIRST_COLUMN & ":" & _
            NAV_LAST_COLUMN).Hidden = False

    'Restore the configured navigation widths
        Ws.Columns(NAV_FIRST_COLUMN).ColumnWidth = NAV_ICON_WIDTH
        Ws.Columns(NAV_LAST_COLUMN).ColumnWidth = NAV_LABEL_WIDTH

End Sub


Private Function GetToggleShape(ByVal Ws As Worksheet) As Shape
'
'==============================================================================
' GetToggleShape
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the named navigation toggle from a worksheet when it exists
'
' WHY THIS EXISTS
'   Shape lookup raises an expected runtime error when a toggle was deleted or
'   has not yet been installed
'
' INPUTS
'   Ws
'     Worksheet containing the expected toggle
'
' RETURNS
'   Shape
'     Existing toggle shape
'     Nothing when the shape is unavailable
'
' BEHAVIOR
'   Performs a localized shape lookup with expected-error suppression
'
' ERROR POLICY
'   Missing-shape lookup errors return Nothing
'   No MsgBox is raised
'
' DEPENDENCIES
'   - TOGGLE_SHAPE_NAME
'
' CALLED FROM
'   - RefreshActiveSheet
'   - EnsureToggleButton
'   - PositionToggleButton
'   - UpdateToggleAppearance
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' FIND TOGGLE
'------------------------------------------------------------------------------
    'Return Nothing unless the named shape is found
        Set GetToggleShape = Nothing

    'Suppress only the expected missing-shape lookup error
        On Error Resume Next
        Set GetToggleShape = Ws.Shapes(TOGGLE_SHAPE_NAME)
        On Error GoTo 0

End Function


Private Function HasNavigationPane(ByVal Ws As Worksheet) As Boolean
'
'==============================================================================
' HasNavigationPane
'------------------------------------------------------------------------------
' PURPOSE
'   Determines whether a worksheet contains the expected navigation headers
'
' WHY THIS EXISTS
'   A user may delete or overwrite the generated pane without changing the
'   workbook worksheet structure
'
' INPUTS
'   Ws
'     Worksheet to inspect
'
' RETURNS
'   Boolean
'     True when both generated headers are present
'     False otherwise
'
' BEHAVIOR
'   Performs a case-sensitive comparison against the configured header text
'
' ERROR POLICY
'   Unexpected runtime errors return False
'
' DEPENDENCIES
'   - NAV_TITLE_TEXT
'   - NAV_HEADER_TEXT
'
' CALLED FROM
'   - RefreshActiveSheet
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Return False unless both generated headers are present
        HasNavigationPane = False

    'Treat unexpected inspection errors as a missing pane
        On Error GoTo Clean_Exit

'------------------------------------------------------------------------------
' VALIDATE GENERATED HEADERS
'------------------------------------------------------------------------------
    'Validate the generated title and section-header text
        HasNavigationPane = _
            (StrComp( _
                CStr(Ws.Range( _
                    NAV_FIRST_COLUMN & CStr(NAV_TITLE_ROW)).Value2), _
                NAV_TITLE_TEXT, _
                vbBinaryCompare) = 0) And _
            (StrComp( _
                CStr(Ws.Range( _
                    NAV_FIRST_COLUMN & CStr(NAV_HEADER_ROW)).Value2), _
                NAV_HEADER_TEXT, _
                vbBinaryCompare) = 0)

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
Clean_Exit:

End Function


Private Function IsNavigationCollapsed(ByVal Ws As Worksheet) As Boolean
'
'==============================================================================
' IsNavigationCollapsed
'------------------------------------------------------------------------------
' PURPOSE
'   Determines whether both reserved navigation columns are hidden
'
' WHY THIS EXISTS
'   Reading the Hidden property from a multi-column range may return Null when
'   the columns have inconsistent visibility states
'
' INPUTS
'   Ws
'     Worksheet containing the navigation pane
'
' RETURNS
'   Boolean
'     True when both navigation columns are hidden
'     False otherwise
'
' BEHAVIOR
'   Tests each reserved navigation column independently
'
' ERROR POLICY
'   Runtime errors propagate to the calling procedure
'
' DEPENDENCIES
'   - NAV_FIRST_COLUMN
'   - NAV_LAST_COLUMN
'
' CALLED FROM
'   - ToggleNavigation
'   - BuildNavigation
'   - UpdateToggleAppearance
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' RETURN NAVIGATION STATE
'------------------------------------------------------------------------------
    'Return True only when both navigation columns are hidden
        IsNavigationCollapsed = _
            Ws.Columns(NAV_FIRST_COLUMN).Hidden And _
            Ws.Columns(NAV_LAST_COLUMN).Hidden

End Function


Private Function IsNavigationSheet(ByVal Ws As Worksheet) As Boolean
'
'==============================================================================
' IsNavigationSheet
'------------------------------------------------------------------------------
' PURPOSE
'   Determines whether a worksheet is eligible for the navigation interface
'
' WHY THIS EXISTS
'   Public procedures may be called while another workbook is active so the
'   worksheet must be confirmed as belonging to ThisWorkbook before modification
'
' INPUTS
'   Ws
'     Worksheet to evaluate
'
' RETURNS
'   Boolean
'     True for visible eligible worksheets belonging to ThisWorkbook
'     False otherwise
'
' BEHAVIOR
'   - Rejects an uninitialized worksheet reference
'   - Rejects worksheets belonging to another workbook
'   - Rejects hidden and very-hidden worksheets
'   - Excludes the configured setup worksheet
'
' ERROR POLICY
'   Unexpected runtime errors return False
'
' DEPENDENCIES
'   - SETUP_SHEET_NAME
'
' CALLED FROM
'   - InstallNavigation
'   - ToggleNavigation
'   - RefreshActiveSheet
'   - InstallNavigationOnSheet
'   - BuildNavigation
'   - GetNavigationSignature
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Return False unless every eligibility test is passed
        IsNavigationSheet = False

    'Treat unexpected validation errors as an ineligible worksheet
        On Error GoTo Clean_Exit

'------------------------------------------------------------------------------
' VALIDATE WORKSHEET
'------------------------------------------------------------------------------
    'Reject an uninitialized worksheet reference
        If Ws Is Nothing Then GoTo Clean_Exit

    'Reject worksheets that do not belong to the workbook containing this code
        If Not (Ws.Parent Is ThisWorkbook) Then GoTo Clean_Exit

    'Reject hidden and very-hidden worksheets
        If Ws.Visible <> xlSheetVisible Then GoTo Clean_Exit

    'Reject the configured setup worksheet
        If StrComp( _
            Ws.Name, _
            SETUP_SHEET_NAME, _
            vbTextCompare) = 0 Then

            GoTo Clean_Exit

        End If

'------------------------------------------------------------------------------
' RETURN ELIGIBLE WORKSHEET
'------------------------------------------------------------------------------
    'All validation conditions have been satisfied
        IsNavigationSheet = True

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
Clean_Exit:

End Function


Private Function GetExistingNavigationLastRow( _
    ByVal Ws As Worksheet) _
    As Long
'
'==============================================================================
' GetExistingNavigationLastRow
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the final used row in the reserved navigation columns
'
' WHY THIS EXISTS
'   Rebuilding navigation must clear stale entries when the number of eligible
'   worksheets decreases
'
' INPUTS
'   Ws
'     Worksheet containing the navigation pane
'
' RETURNS
'   Long
'     Final used row across both reserved navigation columns
'     Never less than NAV_HEADER_ROW
'
' BEHAVIOR
'   Uses End(xlUp) from the bottom of each reserved column
'
' ERROR POLICY
'   Runtime errors propagate to ClearNavigationArea
'
' DEPENDENCIES
'   - NAV_FIRST_COLUMN
'   - NAV_LAST_COLUMN
'   - NAV_HEADER_ROW
'
' CALLED FROM
'   - ClearNavigationArea
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LastIconRow                 As Long
    Dim LastLabelRow                As Long

'------------------------------------------------------------------------------
' IDENTIFY LAST USED ROWS
'------------------------------------------------------------------------------
    'Identify the final used row in the icon column
        LastIconRow = Ws.Cells( _
            Ws.Rows.Count, _
            Ws.Columns(NAV_FIRST_COLUMN).Column).End(xlUp).Row

    'Identify the final used row in the label column
        LastLabelRow = Ws.Cells( _
            Ws.Rows.Count, _
            Ws.Columns(NAV_LAST_COLUMN).Column).End(xlUp).Row

'------------------------------------------------------------------------------
' RETURN FINAL NAVIGATION ROW
'------------------------------------------------------------------------------
    'Return the greatest row while preserving the header boundary
        GetExistingNavigationLastRow = Application.Max( _
            NAV_HEADER_ROW, _
            LastIconRow, _
            LastLabelRow)

End Function


Private Function GetNavigationSignature() As String
'
'==============================================================================
' GetNavigationSignature
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a deterministic signature representing the worksheets currently
'   included in the navigation interface
'
' WHY THIS EXISTS
'   Navigation should be rebuilt only when workbook structure changes rather
'   than every time the user activates another worksheet
'
' INPUTS
'   None
'
' RETURNS
'   String
'     Concatenated signature containing the names and order of all eligible
'     worksheets
'
' BEHAVIOR
'   - Includes only worksheets accepted by IsNavigationSheet
'   - Prefixes each worksheet name with its length to avoid ambiguous joins
'   - Changes after insertion deletion rename hide unhide or reordering
'
' ERROR POLICY
'   Runtime errors propagate to the calling public procedure
'
' DEPENDENCIES
'   - IsNavigationSheet
'
' CALLED FROM
'   - InstallNavigation
'   - RefreshActiveSheet
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Ws                          As Worksheet
    Dim NavigationSignature         As String

'------------------------------------------------------------------------------
' BUILD WORKSHEET SIGNATURE
'------------------------------------------------------------------------------
    'Initialize the signature buffer
        NavigationSignature = vbNullString

    For Each Ws In ThisWorkbook.Worksheets

        'Include only worksheets eligible for navigation
            If IsNavigationSheet(Ws) Then

                'Prefix the name with its length to prevent ambiguous joins
                    NavigationSignature = NavigationSignature & _
                                          CStr(Len(Ws.Name)) & ":" & _
                                          Ws.Name & "|"

            End If

    Next Ws

'------------------------------------------------------------------------------
' RETURN SIGNATURE
'------------------------------------------------------------------------------
    'Return the completed worksheet signature
        GetNavigationSignature = NavigationSignature

End Function


Private Function NavigationIcon(ByVal SheetName As String) As String
'
'==============================================================================
' NavigationIcon
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a compact icon or abbreviation for a worksheet name
'
' WHY THIS EXISTS
'   The first navigation column provides a concise visual cue without consuming
'   the width required by the full worksheet name
'
' INPUTS
'   SheetName
'     Worksheet name for which an icon or abbreviation is required
'
' RETURNS
'   String
'     Configured symbol for known worksheets
'     Uppercase first two characters for other worksheets
'
' BEHAVIOR
'   Performs a case-insensitive name lookup and returns a deterministic fallback
'
' ERROR POLICY
'   Runtime errors propagate to BuildNavigation
'
' DEPENDENCIES
'   None
'
' CALLED FROM
'   - BuildNavigation
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' MAP WORKSHEET NAME
'------------------------------------------------------------------------------
    'Return the configured icon or abbreviation for the worksheet
        Select Case LCase$(SheetName)

            Case "home"
                NavigationIcon = ChrW(&H2302)

            Case "normal"
                NavigationIcon = "N"

            Case "lognormal"
                NavigationIcon = "LN"

            Case "gamma"
                NavigationIcon = ChrW(&H393)

            Case "beta"
                NavigationIcon = ChrW(&H3B2)

            Case "poisson"
                NavigationIcon = "P"

            Case "binomial"
                NavigationIcon = "B"

            Case "examples"
                NavigationIcon = ChrW(&H2605)

            Case "documentation"
                NavigationIcon = "?"

            Case Else
                NavigationIcon = UCase$(Left$(SheetName, 2))

        End Select

End Function


Private Function SafeWorksheetName(ByVal Ws As Worksheet) As String
'
'==============================================================================
' SafeWorksheetName
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a worksheet name without raising a secondary diagnostic error
'
' WHY THIS EXISTS
'   Error handlers may execute with an uninitialized or invalid worksheet
'   reference
'
' INPUTS
'   Ws
'     Worksheet whose name is required
'
' RETURNS
'   String
'     Worksheet name when available
'     vbNullString otherwise
'
' BEHAVIOR
'   Performs a localized expected-error-safe name lookup
'
' ERROR POLICY
'   Unexpected lookup errors return vbNullString
'
' DEPENDENCIES
'   None
'
' CALLED FROM
'   - ToggleNavigation
'   - RefreshActiveSheet
'   - InstallNavigationOnSheet
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' RETURN WORKSHEET NAME
'------------------------------------------------------------------------------
    'Return an empty string unless the worksheet name is available
        SafeWorksheetName = vbNullString

    'Suppress only errors raised while reading the diagnostic name
        On Error Resume Next
        SafeWorksheetName = Ws.Name
        On Error GoTo 0

End Function


Private Sub ReportNavigationError( _
    ByVal ProcedureName As String, _
    ByVal ErrorNumber As Long, _
    ByVal ErrorDescription As String, _
    Optional ByVal WorksheetName As String = vbNullString)
'
'==============================================================================
' ReportNavigationError
'------------------------------------------------------------------------------
' PURPOSE
'   Writes a structured navigation error message to the VBA Immediate Window
'
' WHY THIS EXISTS
'   Navigation errors should not interrupt workbook use with message boxes but
'   must remain observable for testing support and troubleshooting
'
' INPUTS
'   ProcedureName
'     Name of the procedure in which the error occurred
'
'   ErrorNumber
'     VBA runtime error number
'
'   ErrorDescription
'     VBA runtime error description
'
'   WorksheetName
'     Optional name of the worksheet associated with the failure
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Builds a timestamped diagnostic message
'   - Includes the module procedure worksheet error number and description
'   - Writes the result to the VBA Immediate Window
'
' ERROR POLICY
'   Diagnostic reporting errors are suppressed
'
' DEPENDENCIES
'   None
'
' CALLED FROM
'   - InstallNavigation
'   - ToggleNavigation
'   - RefreshActiveSheet
'   - InstallNavigationOnSheet
'
' UPDATED
'   2026-07-12
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim DiagnosticMessage           As String

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Diagnostic reporting must never raise a secondary runtime error
        On Error Resume Next

'------------------------------------------------------------------------------
' BUILD DIAGNOSTIC MESSAGE
'------------------------------------------------------------------------------
    'Start with the timestamp module and procedure name
        DiagnosticMessage = _
            Format$(Now, "yyyy-mm-dd hh:nn:ss") & _
            " | M_NAVIGATION." & _
            ProcedureName

    'Append the worksheet name when one was supplied
        If Len(WorksheetName) > 0 Then

            DiagnosticMessage = DiagnosticMessage & _
                                " | Worksheet: " & _
                                WorksheetName

        End If

    'Append the runtime error number and description
        DiagnosticMessage = DiagnosticMessage & _
                            " | Error " & _
                            CStr(ErrorNumber) & _
                            ": " & _
                            ErrorDescription

'------------------------------------------------------------------------------
' WRITE DIAGNOSTIC
'------------------------------------------------------------------------------
    'Write the structured message to the VBA Immediate Window
        Debug.Print DiagnosticMessage

End Sub


