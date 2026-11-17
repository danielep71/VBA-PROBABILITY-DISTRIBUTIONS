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

function Read-Inputs($Sheet) {
    return @(6..11 | ForEach-Object { $Sheet.Range('D' + $_).Value2 })
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($path,0,$true)
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
    if ([Math]::Abs([double]$sheet.Range('G8').Value2-0.5) -gt 1e-12) { throw 'Ribbon recalculate result differs' }
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
    $sheet.Activate()

    # Resolve the real confirmation dialog by owned process, exact prompt and
    # standard button ID. Never send keystrokes to the foreground desktop.
    foreach ($choice in @('2','1')) {
        $sheet.Range('D8').Value2 = 2.0
        $before = (Read-Inputs $sheet) | ConvertTo-Json -Compress
        $dialogJob = Start-Job -ArgumentList $excelProcessId,$choice -ScriptBlock {
            param($OwnedProcessId,$ButtonId)
            $ErrorActionPreference = 'Stop'
            Add-Type -AssemblyName UIAutomationClient
            Add-Type -AssemblyName UIAutomationTypes
            $owned = [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ProcessIdProperty,[int]$OwnedProcessId)
            $deadline = [DateTime]::UtcNow.AddSeconds(20)
            do {
                $windows = [Windows.Automation.AutomationElement]::RootElement.FindAll([Windows.Automation.TreeScope]::Children,$owned)
                foreach ($candidate in $windows) {
                    if ($candidate.Current.Name -ne 'Distribution Explorer') { continue }
                    $prompt = [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty,'Reset only Explorer inputs to Normal(0,1)?')
                    if ($null -eq $candidate.FindFirst([Windows.Automation.TreeScope]::Descendants,$prompt)) { continue }
                    $id = [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::AutomationIdProperty,[string]$ButtonId)
                    $button = $candidate.FindFirst([Windows.Automation.TreeScope]::Descendants,$id)
                    if ($null -eq $button) { throw 'Expected confirmation button missing' }
                    $button.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
                    return "Clicked confirmation button $ButtonId"
                }
                Start-Sleep -Milliseconds 100
            } while ([DateTime]::UtcNow -lt $deadline)
            throw 'The owned Explorer confirmation dialog was not observed'
        }
        Invoke-Control (Find-Control 'Reset Explorer inputs')
        $completed = Wait-Job $dialogJob -Timeout 25
        if ($null -eq $completed -or $dialogJob.State -ne 'Completed') { throw 'Confirmation worker failed or timed out' }
        $report['dialog_button_' + $choice] = Receive-Job $dialogJob -ErrorAction Stop
        Remove-Job $dialogJob
        $dialogJob = $null
        $after = (Read-Inputs $sheet) | ConvertTo-Json -Compress
        if ($choice -eq '2') {
            if ($before -ne $after) { throw 'Cancel changed Explorer inputs' }
            $report.reset_cancel = 'PASS'
        } else {
            if ($sheet.Range('D6').Value2 -ne 'Normal' -or $sheet.Range('D8').Value2 -ne 1.0) { throw 'Confirmed reset did not restore defaults' }
            $report.reset_confirm = 'PASS'
        }
        if (-not $excel.EnableEvents -or -not $sheet.ProtectContents) { throw 'Ribbon reset leaked state or protection' }
    }
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
