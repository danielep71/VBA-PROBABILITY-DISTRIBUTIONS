param([Parameter(Mandatory=$true)][string]$PackageDirectory)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$report = @{status='FAIL'; sha256=(Get-FileHash $path -Algorithm SHA256).Hash}
$excel = $null
$book = $null
$dialogJob = $null
$window = $null

function Find-Control([string]$Name) {
    $condition = [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty,$Name)
    $matches = $window.FindAll([Windows.Automation.TreeScope]::Descendants,$condition)
    foreach ($item in $matches) {
        if (-not $item.Current.IsOffscreen -and $item.Current.IsEnabled) { return $item }
    }
    throw "Visible enabled control not found in the owned Excel window: $Name"
}

function Invoke-Control($Element) {
    $pattern = $Element.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern)
    $pattern.Invoke()
}

# A control that is shown, enabled or not: used to read the enabled state of a button.
function Find-Shown([string]$Name) {
    $condition = [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty,$Name)
    foreach ($item in $window.FindAll([Windows.Automation.TreeScope]::Descendants,$condition)) {
        if (-not $item.Current.IsOffscreen) { return $item }
    }
    throw "Visible control not found in the owned Excel window: $Name"
}

# Open the Go to menu and choose the entry of a sheet. Office may host the drop-down in a separate popup HWND
# in the same owned process; worksheet tabs with the same label are excluded by control type.
function Invoke-GoTo([string]$Tab) {
    $split = Find-Control 'Go to'
    $expand = $split.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern)
    $expand.Expand()
    $owned = [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ProcessIdProperty,[int]$excelProcessId)
    $roots = [Windows.Automation.AutomationElement]::RootElement.FindAll([Windows.Automation.TreeScope]::Children,$owned)
    $named = [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty,$Tab)
    foreach ($root in $roots) {
        foreach ($item in $root.FindAll([Windows.Automation.TreeScope]::Descendants,$named)) {
            if (-not $item.Current.IsOffscreen -and $item.Current.IsEnabled -and
                $item.Current.ControlType -in @([Windows.Automation.ControlType]::MenuItem,[Windows.Automation.ControlType]::Button)) {
                Invoke-Control $item
                return
            }
        }
    }
    throw "Ribbon menu target not found: $Tab"
}

# A worker that waits for the owned dialog with this title and prompt and presses the first of the button IDs
# it has (1 OK, 2 Cancel). It never sends keystrokes to the foreground desktop.
function Start-DialogWorker([string]$Title, [string]$Prompt, [int[]]$ButtonIds) {
    return Start-Job -ArgumentList $excelProcessId,([long]$excel.Hwnd),$Title,$Prompt,$ButtonIds -ScriptBlock {
        param($OwnedProcessId,$OwnedWindowHandle,$Title,$Prompt,$ButtonIds)
        $ErrorActionPreference = 'Stop'
        Add-Type -AssemblyName UIAutomationClient
        Add-Type -AssemblyName UIAutomationTypes
        Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class ExplorerDialogNative {
    [DllImport("user32.dll")]
    public static extern IntPtr GetDlgItem(IntPtr parent, int id);
    [DllImport("user32.dll", SetLastError=true)]
    public static extern IntPtr SendMessageTimeout(IntPtr window, uint message,
        UIntPtr wParam, IntPtr lParam, uint flags, uint timeout, out UIntPtr result);
}
'@
        $ownedRoot = [Windows.Automation.AutomationElement]::FromHandle([IntPtr]$OwnedWindowHandle)
        $named = [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty,$Title)
        $deadline = [DateTime]::UtcNow.AddSeconds(20)
        do {
            foreach ($candidate in $ownedRoot.FindAll([Windows.Automation.TreeScope]::Descendants,$named)) {
                if ($candidate.Current.ProcessId -ne $OwnedProcessId -or
                    $candidate.Current.ControlType -ne [Windows.Automation.ControlType]::Window) { continue }
                $text = [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty,$Prompt)
                if ($null -eq $candidate.FindFirst([Windows.Automation.TreeScope]::Descendants,$text)) { continue }
                foreach ($buttonId in $ButtonIds) {
                    $id = [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::AutomationIdProperty,[string]$buttonId)
                    $button = $candidate.FindFirst([Windows.Automation.TreeScope]::Descendants,$id)
                    $invoke = $null
                    if ($null -ne $button -and $button.TryGetCurrentPattern([Windows.Automation.InvokePattern]::Pattern,[ref]$invoke)) {
                        $invoke.Invoke()
                        return "Clicked button $buttonId"
                    }
                    # This Excel build exposes the standard dialog buttons as Pane controls. Address only the
                    # matched dialog's child.
                    $handle = [ExplorerDialogNative]::GetDlgItem([IntPtr]$candidate.Current.NativeWindowHandle,[int]$buttonId)
                    if ($handle -eq [IntPtr]::Zero) { continue }
                    $result = [UIntPtr]::Zero
                    $sent = [ExplorerDialogNative]::SendMessageTimeout($handle,0xF5,[UIntPtr]::Zero,[IntPtr]::Zero,2,5000,[ref]$result)
                    if ($sent -eq [IntPtr]::Zero) { throw 'Matched dialog button did not respond' }
                    return "Clicked button $buttonId"
                }
                throw 'Expected dialog button missing'
            }
            Start-Sleep -Milliseconds 100
        } while ([DateTime]::UtcNow -lt $deadline)
        throw "The owned dialog was not observed: $Title / $Prompt"
    }
}

# Wait for a dialog worker and return what it did.
function Complete-DialogWorker($Job) {
    $completed = Wait-Job $Job -Timeout 25
    if ($null -eq $completed -or $Job.State -ne 'Completed') {
        $report.dialog_worker_errors = @($Job.ChildJobs[0].Error | ForEach-Object { $_.Exception.Message })
        throw 'Dialog worker failed or timed out'
    }
    return (Receive-Job $Job -ErrorAction Stop)
}

function Read-Inputs($Sheet) {
    return @(6..11 | ForEach-Object { $Sheet.Range('D' + $_).Value2 })
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($path,0,$true)
    # Excel may still be loading or calculating the sheets after the workbook opens: a COM call is rejected until it is ready.
    for ($wait = 0; $wait -lt 240; $wait++) { try { if ($excel.Ready) { break } } catch { }; Start-Sleep -Milliseconds 500 }
    $sheet = $book.Worksheets.Item('Explorer')
    $book.Activate()
    $sheet.Activate()
    $excel.WindowState = -4137
    $window = [Windows.Automation.AutomationElement]::FromHandle([IntPtr]$excel.Hwnd)
    $excelProcessId = $window.Current.ProcessId
    $report.excel_version = $excel.Version
    $report.excel_build = $excel.Build
    $report.scope = 'Native Ribbon controls via Windows UI Automation, in an owned Excel process; no global keyboard or mouse input.'
    $tab = Find-Control 'Probability Distributions'
    $pattern = $null
    if ($tab.TryGetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern,[ref]$pattern)) {
        $pattern.Select()
    } else {
        Invoke-Control $tab
    }
    Invoke-Control (Find-Control 'Recalculate')
    if ([Math]::Abs([double]$sheet.Range('D22').Value2-0.5) -gt 1e-12) { throw 'Ribbon recalculate result differs' }
    $report.ribbon_recalculate = 'PASS'

    $toggle = Find-Control 'Navigation bar'
    $togglePattern = $toggle.GetCurrentPattern([Windows.Automation.TogglePattern]::Pattern)
    $togglePattern.Toggle()
    foreach ($item in $book.Worksheets) {
        if (-not $item.Columns.Item(1).Hidden) { throw 'Ribbon toggle did not hide every bar' }
    }
    $togglePattern.Toggle()
    foreach ($item in $book.Worksheets) {
        if ($item.Columns.Item(1).Hidden) { throw 'Ribbon toggle did not restore every bar' }
    }
    $report.ribbon_navigation_toggle = 'PASS'

    Invoke-Control (Find-Control 'Go to')
    if ($excel.ActiveSheet.CodeName -ne 'shHome') { throw 'Ribbon Go to did not open Home' }
    $report.ribbon_home = 'PASS'
    $sheetMap = Get-Content (Join-Path $PSScriptRoot '../examples/demo/sheets.json') -Raw | ConvertFrom-Json
    $visited = @()
    foreach ($section in $sheetMap.sections) {
        foreach ($entry in $section.sheets) {
            Invoke-GoTo ([string]$entry.tab)
            if ($excel.ActiveWorkbook.FullName -ne $book.FullName -or $excel.ActiveSheet.CodeName -ne $entry.code) {
                throw "Ribbon menu reached the wrong target: $($entry.tab)"
            }
            $visited += $entry.code
        }
    }
    $report.ribbon_menu_targets = $visited

    # Reset inputs acts on the active sheet: disabled on Home and README, which have no inputs, enabled on a
    # sheet with inputs. Activating a sheet makes the Ribbon read the state again (Workbook_SheetActivate).
    $states = [ordered]@{}
    foreach ($pair in @(@('Home', $false), @('README', $false), @('Probabilities', $true), @('Explorer', $true))) {
        $book.Worksheets.Item($pair[0]).Activate()
        $enabled = $false
        for ($wait = 0; $wait -lt 20; $wait++) {
            $enabled = (Find-Shown 'Reset inputs').Current.IsEnabled
            if ($enabled -eq $pair[1]) { break }
            Start-Sleep -Milliseconds 250
        }
        if ($enabled -ne $pair[1]) { throw "Reset inputs is enabled=$enabled on $($pair[0]), expected $($pair[1])" }
        $states[$pair[0]] = $enabled
    }
    $report.reset_enabled_by_sheet = $states
    $sheet.Activate()

    # Resolve the real confirmation dialog by owned process, exact prompt and
    # standard button ID. Never send keystrokes to the foreground desktop.
    foreach ($choice in @('2','1')) {
        $excel.EnableEvents = $false
        $sheet.Range('D6').Value2 = 'Poisson'
        $sheet.Range('D7').Value2 = 7.0
        $excel.EnableEvents = $true
        $sheet.Range('D11').Value2 = 2.0
        $before = (Read-Inputs $sheet) | ConvertTo-Json -Compress
        $dialogJob = Start-DialogWorker 'Distribution Explorer' 'Reset only Explorer inputs to Normal(0,1)?' @([int]$choice)
        Invoke-Control (Find-Control 'Reset inputs')
        $report['dialog_button_' + $choice] = Complete-DialogWorker $dialogJob
        Remove-Job $dialogJob
        $dialogJob = $null
        $after = (Read-Inputs $sheet) | ConvertTo-Json -Compress
        if ($choice -eq '2') {
            if ($before -ne $after) { throw 'Cancel changed Explorer inputs' }
            $report.reset_cancel = 'PASS'
        } else {
            $actual = Read-Inputs $sheet
            $expected = @('Normal',0.0,1.0,0.0,0.0,0.0)
            for ($index=0; $index -lt $expected.Count; $index++) {
                if ($actual[$index] -ne $expected[$index]) { throw "Confirmed reset did not restore input $index" }
            }
            $report.reset_confirm = 'PASS'
        }
        if (-not $excel.EnableEvents) { throw 'Ribbon reset leaked events state' }
    }

    # A renamed tab and a missing sheet, in this read-only copy (nothing is saved). Go to opens a sheet by its
    # CodeName, so the renamed tab is still reached; an entry whose sheet is gone shows a message and changes
    # nothing.
    $book.Unprotect()
    $excel.DisplayAlerts = $false
    $book.Worksheets.Item('Tails').Name = 'Tails renamed'
    Invoke-GoTo 'Tails'
    if ($excel.ActiveWorkbook.FullName -ne $book.FullName -or $excel.ActiveSheet.CodeName -ne 'shTails') {
        throw 'Go to did not reach the renamed Tails sheet'
    }
    $report.goto_renamed_tab = 'PASS'
    $book.Worksheets.Item('Risk').Delete()
    $sheet.Activate()
    $dialogJob = Start-DialogWorker 'Probability Distributions' 'This workbook has no sheet for the Ribbon entry pdGoRisk.' @(2, 1)
    Invoke-GoTo 'Risk'
    $report.goto_missing_sheet_dialog = Complete-DialogWorker $dialogJob
    Remove-Job $dialogJob
    $dialogJob = $null
    if ($excel.ActiveWorkbook.FullName -ne $book.FullName -or $excel.ActiveSheet.CodeName -ne 'shExplorer') {
        throw 'Go to a missing sheet changed the active sheet'
    }
    $report.goto_missing_sheet = 'PASS'
    $report.status = 'PASS'
} catch {
    $report.error = $_.Exception.Message
    $report.position = $_.InvocationInfo.PositionMessage
    if ($null -ne $window) {
        # Bounded diagnostic list of controls in this owned window only.
        $nodes = $window.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.Condition]::TrueCondition)
        $report.controls = @($nodes | Where-Object { $_.Current.Name -and $_.Current.ControlType.ProgrammaticName -ne 'ControlType.DataItem' } |
            Select-Object -First 120 | ForEach-Object { @{name=$_.Current.Name;type=$_.Current.ControlType.ProgrammaticName;id=$_.Current.AutomationId;offscreen=$_.Current.IsOffscreen} })
    }
    throw
} finally {
    if ($null -ne $dialogJob) { Stop-Job $dialogJob; Remove-Job $dialogJob }
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    $report | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 (Join-Path $directory 'explorer-ui-report.json')
    $report | ConvertTo-Json -Depth 5 | Write-Output
}
