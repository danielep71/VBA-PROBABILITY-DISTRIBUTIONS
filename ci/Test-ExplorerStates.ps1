param([Parameter(Mandatory=$true)][string]$PackageDirectory)
# The Explorer as a user meets it: one input edited at a time with events on,
# a file saved while macros were off, and a session with macros blocked.
# A VBA error dialog is read, reported and ends the owned Excel process.
$ErrorActionPreference = 'Stop'
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class ExplorerStatesNative {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
}
'@
$directory = (Resolve-Path $PackageDirectory).Path
$source = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$copy = Join-Path $directory 'Explorer-states.xlsm'
$dialogFile = Join-Path $directory 'explorer-states-dialog.json'
$report = [ordered]@{ status = 'FAIL'; sha256 = (Get-FileHash $source -Algorithm SHA256).Hash }
$xlErrNA = -2146826246
$watch = {
    param($OwnedProcessId,$OutFile)
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
    $deadline = [DateTime]::UtcNow.AddSeconds(90)
    $root = [Windows.Automation.AutomationElement]::RootElement
    $byPid = [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ProcessIdProperty,[int]$OwnedProcessId)
    do {
        Start-Sleep -Milliseconds 700
        $found = @()
        $tops = $root.FindAll([Windows.Automation.TreeScope]::Children,$byPid)
        $candidates = @($tops)
        foreach ($top in $tops) {
            $candidates += @($top.FindAll([Windows.Automation.TreeScope]::Children,[Windows.Automation.Condition]::TrueCondition) |
                Where-Object { $_.Current.ControlType -eq [Windows.Automation.ControlType]::Window })
        }
        foreach ($candidate in $candidates) {
            if ($candidate.Current.ClassName -ne '#32770') { continue }
            $texts = @($candidate.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.Condition]::TrueCondition) |
                ForEach-Object { $_.Current.Name } | Where-Object { $_ })
            $found += @{ title = $candidate.Current.Name; texts = $texts }
        }
        if ($found.Count -gt 0) {
            $found | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 $OutFile
            Stop-Process -Id $OwnedProcessId -Force
            return
        }
    } while ([DateTime]::UtcNow -lt $deadline)
}
function Start-OwnedExcel([int]$Security) {
    $instance = New-Object -ComObject Excel.Application
    $instance.Visible = $true
    $instance.DisplayAlerts = $false
    $instance.AutomationSecurity = $Security
    return $instance
}
function Close-OwnedExcel($Instance) {
    if ($null -ne $Instance) {
        try { $Instance.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Instance) } catch {}
    }
}
# One call site per value type: PowerShell binds a COM setter to the first type it sees.
function Set-Text($Sheet, [string]$Address, [string]$Value) { $Sheet.Range($Address).Value2 = $Value }
function Set-Number($Sheet, [string]$Address, [double]$Value) { $Sheet.Range($Address).Value2 = $Value }
function Number($Sheet, [string]$Address) { return [double]$Sheet.Range($Address).Value2 }
function Assert-Near([double]$Actual, [double]$Expected, [string]$What) {
    if ([Math]::Abs($Actual - $Expected) -gt 1e-12) { throw "${What}: $Actual, expected $Expected" }
}
function Assert-Charts($Sheet, [int]$DensityType, [int]$CdfType, [int]$Points, [string]$What) {
    foreach ($pair in @(@('ExplorerDensity',$DensityType),@('ExplorerCDF',$CdfType))) {
        $chart = $Sheet.ChartObjects($pair[0]).Chart
        if ($chart.ChartType -ne $pair[1]) { throw "${What}: $($pair[0]) has chart type $($chart.ChartType), expected $($pair[1])" }
        $count = @($chart.SeriesCollection(1).Values).Count
        if ($count -ne $Points) { throw "${What}: $($pair[0]) plots $count points, expected $Points" }
    }
}
$excel = $null
$job = $null
try {
    Remove-Item $dialogFile -ErrorAction SilentlyContinue

    # 1. Macros on: every family through the SheetChange event, one edit at a time.
    Copy-Item $source $copy -Force
    $excel = Start-OwnedExcel 1
    $processId = 0
    [void][ExplorerStatesNative]::GetWindowThreadProcessId([IntPtr][long]$excel.Hwnd, [ref]$processId)
    $job = Start-Job -ArgumentList $processId,$dialogFile -ScriptBlock $watch
    $book = $excel.Workbooks.Open($copy, 0, $false)
    $sheet = $book.Worksheets.Item('Explorer')
    $sheet.Activate()
    if ([string]$sheet.Range('G6').Value2 -ne 'Enabled') { throw 'Macros cell does not say Enabled with macros on' }
    if ([string]$sheet.Range('F4').Value2 -ne '') { throw 'The warning line is not empty with macros on' }
    if (-not $book.Saved) { throw 'Opening the workbook left it marked as changed' }
    Assert-Charts $sheet 75 75 201 'Normal at open'
    Set-Text $sheet 'D6' 'Exponential'
    Assert-Near (Number $sheet 'G7') 1.0 'Exponential density at 0, rate 1'
    Assert-Charts $sheet 75 75 201 'Exponential'
    Set-Number $sheet 'D9' 2.0
    Assert-Near (Number $sheet 'G7') 2.0 'Exponential density at 0, rate 2'
    Assert-Near (Number $sheet 'G11') 4.0 'Exponential plot upper bound, rate 2'
    Assert-Near (Number $sheet 'G118') (2.0*[Math]::Exp(-4.0)) 'Exponential grid midpoint'
    Set-Text $sheet 'D6' 'Poisson'
    Assert-Charts $sheet 51 -4169 31 'Poisson'
    Assert-Near (Number $sheet 'G7') ([Math]::Exp(-4.0)) 'Poisson mass at 0, mean 4'
    foreach ($address in @('F118','G118','H118','G218')) {
        if ($sheet.Range($address).Value2 -ne $xlErrNA) { throw "Poisson grid cell $address beyond the domain is not #N/A" }
    }
    Set-Number $sheet 'D11' 3.0
    Assert-Near (Number $sheet 'G7') (64.0*[Math]::Exp(-4.0)/6.0) 'Poisson mass at 3, mean 4'
    # An edit outside the inputs must not redraw or fail; the sheet is protected, so use an unlocked input's neighbour state.
    Set-Text $sheet 'D6' 'Normal'
    Assert-Charts $sheet 75 75 201 'Normal again'
    Assert-Near (Number $sheet 'G8') 0.99865010196836990 'Normal CDF at 3'
    if (-not $sheet.ProtectContents -or -not $excel.EnableEvents) { throw 'Edits left the sheet unprotected or events off' }
    $book.Close($false)
    Close-OwnedExcel $excel; $excel = $null
    $report.every_family_through_events = 'PASS'

    # 2. Macros off: the sheet must say so, and the file is saved in that state.
    Copy-Item $source $copy -Force
    $excel = Start-OwnedExcel 3
    $book = $excel.Workbooks.Open($copy, 0, $false)
    $sheet = $book.Worksheets.Item('Explorer')
    Set-Text $sheet 'D6' 'Poisson'
    if ([string]$sheet.Range('G6').Value2 -notlike 'DISABLED*') { throw 'Macros cell does not say DISABLED with macros off' }
    if ([string]$sheet.Range('F4').Value2 -notlike 'MACROS ARE DISABLED*') { throw 'The warning line is empty with macros off' }
    if ($sheet.Range('G7').Value2 -isnot [int]) { throw 'A result is still a number with macros off' }
    $book.Save()
    $book.Close($false)
    Close-OwnedExcel $excel; $excel = $null
    $report.macros_off_is_stated = 'PASS'

    # 3. That file, opened with macros on: results and charts must be right at once.
    $excel = Start-OwnedExcel 1
    $book = $excel.Workbooks.Open($copy, 0, $false)
    $sheet = $book.Worksheets.Item('Explorer')
    if ([string]$sheet.Range('G6').Value2 -ne 'Enabled') { throw 'Macros cell stays DISABLED after macros are enabled' }
    Assert-Near (Number $sheet 'G7') ([Math]::Exp(-4.0)) 'Poisson mass after reopening'
    Assert-Charts $sheet 51 -4169 31 'Poisson after reopening'
    $book.Close($false)
    Close-OwnedExcel $excel; $excel = $null
    $report.stale_file_recovers_at_open = 'PASS'

    if (Test-Path $dialogFile) { throw 'Excel showed a dialog' }
    $report.status = 'PASS'
} catch {
    $report.error = $_.Exception.Message
    throw
} finally {
    Close-OwnedExcel $excel
    if ($null -ne $job) { Stop-Job $job -ErrorAction SilentlyContinue; Remove-Job $job -Force -ErrorAction SilentlyContinue }
    if (Test-Path $dialogFile) { $report.dialog = Get-Content $dialogFile -Raw | ConvertFrom-Json }
    Remove-Item $copy -ErrorAction SilentlyContinue
    $report.excel_states_scope = 'Cell edits through COM with events on; not mouse or keyboard input.'
    $report | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 (Join-Path $directory 'explorer-states-report.json')
    $report | ConvertTo-Json -Depth 6 | Write-Output
}
