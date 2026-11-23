param([Parameter(Mandatory=$true)][string]$PackageDirectory)
# A demo handler that fails leaves Excel as it found it (#482). With a second
# workbook open, two errors are induced in the Explorer's input handler, in a
# copy opened read-only that is never saved, each after a routine has
# unprotected the sheet:
#   1. a new family is chosen once a named range the parameter messages read is
#      deleted: the error comes after the handler has saved the events state;
#   2. the number of decimals is changed once the family table's name is
#      deleted: the error comes before the old handler saved the events state.
# No password is used: Unprotect without one would prompt for it.
# Each time the handler must show one message (answered here through Windows UI
# Automation, in the owned Excel process only), leave events, screen updating
# and calculation as they were, leave the sheet protected as it was, and leave
# the other workbook untouched.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$report = [ordered]@{ status = 'FAIL'; sha256 = (Get-FileHash $path -Algorithm SHA256).Hash }
$excel = $null
$book = $null
$other = $null
$dialogJob = $null

# A worker that waits for the owned message box with this title, whose text starts with the prompt, presses OK
# and returns the text. It never sends keystrokes to the foreground desktop.
function Start-DialogWorker([string]$Title, [string]$Prompt) {
    return Start-Job -ArgumentList $excelProcessId, ([long]$excel.Hwnd), $Title, $Prompt -ScriptBlock {
        param($OwnedProcessId, $OwnedWindowHandle, $Title, $Prompt)
        $ErrorActionPreference = 'Stop'
        Add-Type -AssemblyName UIAutomationClient
        Add-Type -AssemblyName UIAutomationTypes
        Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class HandlerDialogNative {
    [DllImport("user32.dll")]
    public static extern IntPtr GetDlgItem(IntPtr parent, int id);
    [DllImport("user32.dll", SetLastError=true)]
    public static extern IntPtr SendMessageTimeout(IntPtr window, uint message,
        UIntPtr wParam, IntPtr lParam, uint flags, uint timeout, out UIntPtr result);
}
'@
        $ownedRoot = [Windows.Automation.AutomationElement]::FromHandle([IntPtr]$OwnedWindowHandle)
        $named = [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty, $Title)
        $deadline = [DateTime]::UtcNow.AddSeconds(30)
        do {
            foreach ($candidate in $ownedRoot.FindAll([Windows.Automation.TreeScope]::Descendants, $named)) {
                if ($candidate.Current.ProcessId -ne $OwnedProcessId -or
                    $candidate.Current.ControlType -ne [Windows.Automation.ControlType]::Window) { continue }
                $text = $null
                foreach ($item in $candidate.FindAll([Windows.Automation.TreeScope]::Descendants, [Windows.Automation.Condition]::TrueCondition)) {
                    if ([string]$item.Current.Name -and ([string]$item.Current.Name).StartsWith($Prompt)) { $text = [string]$item.Current.Name; break }
                }
                if ($null -eq $text) { continue }
                foreach ($buttonId in @(2, 1)) {
                    $handle = [HandlerDialogNative]::GetDlgItem([IntPtr]$candidate.Current.NativeWindowHandle, $buttonId)
                    if ($handle -eq [IntPtr]::Zero) { continue }
                    $result = [UIntPtr]::Zero
                    $sent = [HandlerDialogNative]::SendMessageTimeout($handle, 0xF5, [UIntPtr]::Zero, [IntPtr]::Zero, 2, 5000, [ref]$result)
                    if ($sent -eq [IntPtr]::Zero) { throw 'The dialog button did not respond' }
                    return $text
                }
                throw 'The dialog has no OK button'
            }
            Start-Sleep -Milliseconds 100
        } while ([DateTime]::UtcNow -lt $deadline)
        throw "The owned dialog was not observed: $Title / $Prompt"
    }
}

function Complete-DialogWorker($Job) {
    $completed = Wait-Job $Job -Timeout 40
    if ($null -eq $completed -or $Job.State -ne 'Completed') {
        $report.dialog_worker_errors = @($Job.ChildJobs[0].Error | ForEach-Object { $_.Exception.Message })
        throw 'The dialog worker failed or timed out'
    }
    $text = Receive-Job $Job -ErrorAction Stop
    Remove-Job $Job
    return $text
}

# Excel's state and the other workbook, as a text that two moments can be compared by.
function Read-State {
    return ('events=' + $excel.EnableEvents + ' screen=' + $excel.ScreenUpdating + ' calculation=' + $excel.Calculation +
            ' other=' + [string]$other.Worksheets.Item(1).Range('A1').Value2 + '/' + $other.Saved)
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($path, 0, $true)
    for ($wait = 0; $wait -lt 240; $wait++) { try { if ($excel.Ready) { break } } catch { }; Start-Sleep -Milliseconds 500 }
    $window = [Windows.Automation.AutomationElement]::FromHandle([IntPtr]$excel.Hwnd)
    $excelProcessId = $window.Current.ProcessId
    $report.excel_version = $excel.Version
    $report.excel_build = $excel.Build
    $other = $excel.Workbooks.Add()
    $other.Worksheets.Item(1).Range('A1').Value2 = 'isolation sentinel'
    $other.Saved = $true
    $sheet = $book.Worksheets.Item('Explorer')
    $book.Activate()
    $sheet.Activate()
    $excel.EnableEvents = $true
    $before = Read-State

    # 1. A new family, once the input messages cannot be read: the error comes after the unprotect.
    $sheet.Protect()
    $book.Unprotect()
    $book.Names.Item('ExplorerLabels').Delete()
    $dialogJob = Start-DialogWorker 'Distribution Explorer' 'The Explorer could not follow that change:'
    $sheet.Range('D6').Value2 = 'Gamma'
    $report.error_in_messages_message = Complete-DialogWorker $dialogJob
    $dialogJob = $null
    if ((Read-State) -ne $before) { throw "The state changed after an error in the input messages: $(Read-State), was $before" }
    if (-not $sheet.ProtectContents) { throw 'The sheet was left unprotected after an error in the input messages' }
    $report.error_in_messages = 'PASS'

    # 2. New decimals, once the family table cannot be read: the decimals routine fails after its unprotect,
    #    and the handler fails before the point where it used to save the events state.
    $book.Names.Item('ExplorerFamily').Delete()
    $dialogJob = Start-DialogWorker 'Distribution Explorer' 'The Explorer could not follow that change:'
    $sheet.Range('D16').Value2 = 3.0
    $report.error_in_decimals_message = Complete-DialogWorker $dialogJob
    $dialogJob = $null
    if ((Read-State) -ne $before) { throw "The state changed after an error in the decimals: $(Read-State), was $before" }
    if (-not $sheet.ProtectContents) { throw 'The sheet was left unprotected after an error in the decimals' }
    $report.error_in_decimals = 'PASS'

    $report.state = $before
    $report.scope = 'Errors induced in the Explorer input handler, with a second workbook open, in a read-only copy never saved; message boxes answered through Windows UI Automation in the owned Excel process.'
    $report.status = 'PASS'
} catch {
    $report.error = $_.Exception.Message
    $report.position = $_.InvocationInfo.PositionMessage
    throw
} finally {
    if ($null -ne $dialogJob) { Stop-Job $dialogJob; Remove-Job $dialogJob }
    if ($null -ne $other) { $other.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($other) }
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    $report | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 (Join-Path $directory 'handler-errors-report.json')
    $report | ConvertTo-Json -Depth 5 | Write-Output
}
