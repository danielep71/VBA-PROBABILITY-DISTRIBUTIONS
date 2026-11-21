param([Parameter(Mandatory=$true)][string]$PackageDirectory)
# Applies diagnostics-plan.json to the Diagnostics sheet (errors and
# diagnostics) of the package in desktop Excel, after the sheets before it:
# the sheet takes the formats and widths of the Explorer (the same roles:
# inputs, status, results, notes, headings), then the names, the hidden
# catalogue of functions, the try-it results, the case table, the validation,
# the conditional formats and the chart. The plan is made by
# tools/demobuild/diagnostics_plan.py.
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$plan = Get-Content (Join-Path $directory 'diagnostics-plan.json') -Raw | ConvertFrom-Json
$excel = $null
$book = $null
Remove-Item (Join-Path $env:TEMP 'diagnostics-watchdog.txt') -ErrorAction SilentlyContinue
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class EvolveNative {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
}
'@

# Names.Add reads its RefersTo with English function names but the list
# separator and the decimal separator of the user's regional settings (a
# semicolon and a comma in many locales): both are converted outside the quotes.
function Localize-Formula([string]$Formula, [string]$Separator, [string]$Decimal) {
    $builder = New-Object System.Text.StringBuilder
    $quoted = $false
    $characters = $Formula.ToCharArray()
    for ($index = 0; $index -lt $characters.Length; $index++) {
        $character = $characters[$index]
        if ($character -eq '"') { $quoted = -not $quoted }
        if ($quoted) { [void]$builder.Append($character); continue }
        $isPoint = $character -eq '.' -and $index -gt 0 -and $index -lt $characters.Length - 1 -and
            [char]::IsDigit($characters[$index - 1]) -and [char]::IsDigit($characters[$index + 1])
        if ($character -eq ',') { [void]$builder.Append($Separator) }
        elseif ($isPoint) { [void]$builder.Append($Decimal) }
        else { [void]$builder.Append($character) }
    }
    return $builder.ToString()
}

# Excel rejects a COM call while it is busy (RPC_E_CALL_REJECTED): wait and retry.
function Invoke-Excel([scriptblock]$Action) {
    for ($attempt = 1; ; $attempt++) {
        try { return & $Action }
        catch {
            $text = [string]$_.Exception + ' ' + [string]$_.Exception.InnerException
            $busy = $_.Exception.HResult -eq -2147418111 -or $text -match '0x80010001|RPC_E_CALL_REJECTED'
            if (-not $busy -or $attempt -ge 600) { throw }
            if ($attempt % 20 -eq 1) { Write-Host "  Excel is busy; waiting (attempt $attempt)" }
            Start-Sleep -Milliseconds 500
        }
    }
}

function Wait-Excel {
    for ($wait = 0; $wait -lt 600; $wait++) {
        if (Invoke-Excel { $excel.Ready }) { return }
        Start-Sleep -Milliseconds 500
    }
    throw 'Excel stayed busy for five minutes'
}

# A macro that stops on a hidden dialog would hold the step for ever: a watchdog
# lists the windows of the owned Excel process and ends it after two minutes, so
# the step fails with the text of the dialog instead of timing out.
function Start-Watchdog([int]$ProcessId, [int]$Seconds) {
    return Start-Job -ArgumentList $ProcessId, $Seconds -ScriptBlock {
        param($ProcessId, $Seconds)
        Start-Sleep -Seconds $Seconds
        $text = @("watchdog: Excel (pid $ProcessId) is still busy after $Seconds s")
        try {
            Add-Type -AssemblyName UIAutomationClient
            Add-Type -AssemblyName UIAutomationTypes
            $root = [Windows.Automation.AutomationElement]::RootElement
            $byPid = [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ProcessIdProperty, $ProcessId)
            foreach ($top in $root.FindAll([Windows.Automation.TreeScope]::Children, $byPid)) {
                $names = @($top.FindAll([Windows.Automation.TreeScope]::Descendants, [Windows.Automation.Condition]::TrueCondition) |
                    ForEach-Object { $_.Current.Name } | Where-Object { $_ } | Select-Object -First 30)
                $text += ("  window: [" + $top.Current.ClassName + "] " + $top.Current.Name + " :: " + ($names -join ' | '))
            }
        } catch { $text += ('  could not list the windows: ' + $_.Exception.Message) }
        $text | Out-File -FilePath (Join-Path $env:TEMP 'diagnostics-watchdog.txt') -Encoding UTF8
        Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
    }
}

function Step([string]$Text) {
    if ($null -ne $excel) { $excel.DisplayAlerts = $false }
    Write-Output ('[' + (Get-Date -Format 'HH:mm:ss') + '] ' + $Text)
}

# Copies the formats (and merges) of one Explorer range onto a range of the
# Diagnostics sheet.
function Copy-Format($Source, $Target) {
    [void]$Source.Copy()
    [void]$Target.PasteSpecial(-4122)
}

# A plan value is a number or a text: each is written as itself, so a number of the catalogue stays a number.
function Set-Value($Range, $Value) {
    if ($Value -is [string]) { $Range.Value2 = [string]$Value } else { $Range.Value2 = [double]$Value }
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $excel.ScreenUpdating = $false
    $book = $excel.Workbooks.Open($path, 0, $false)
    try { $book.Unprotect() } catch { }
    $explorer = $book.Worksheets.Item('Explorer')
    $sheet = $book.Worksheets.Item($plan.sheet)
    try { $sheet.Unprotect() } catch { }
    $excel.Calculation = -4135
    $dec = [string]$excel.International(3)
    $at = $plan.layout
    $header = [int]$at.header_row
    $first = [int]$at.first_row
    $last = [int]$at.last_row
    $rows = $at.result_rows
    $columns = $at.table
    $prefix = "'" + $book.Name + "'!"

    Step 'clear the sheet and take the Explorer formats'
    foreach ($old in @($sheet.ChartObjects())) { $old.Delete() }
    $sheet.Columns.Hidden = $false
    $sheet.Range("A1:P$last").UnMerge()
    $sheet.Range("A1:AZ$last").ClearContents() | Out-Null
    $sheet.Range("B1:AZ$last").ClearFormats() | Out-Null
    $sheet.Range("A1:AZ$last").Validation.Delete() | Out-Null
    $sheet.Cells.FormatConditions.Delete() | Out-Null
    Copy-Format $explorer.Range('A1:P20') $sheet.Range('A1')
    try { $excel.CutCopyMode = $false } catch { }
    # The inputs are D6:D12 here: the rows the Explorer keeps for its other inputs are spacers.
    foreach ($address in $at.spacers) {
        $target = $sheet.Range($address)
        Copy-Format $explorer.Range('C17').Resize($target.Rows.Count, 2) $target
    }
    # Rows 19 and 20 of the Explorer carry its grid headings and first grid row in F:I.
    $sheet.Range('F19:I20').ClearFormats()
    # The results: a label in C, the value in D. The returned value has its formula as text over F:I; every
    # other result is a text over D:I.
    foreach ($property in $rows.PSObject.Properties) {
        $row = [int]$property.Value
        Copy-Format $explorer.Range('C21:D21') $sheet.Range("C${row}:D$row")
        Copy-Format $explorer.Range('D21') $sheet.Range("E${row}:I$row")
        if ($property.Name -eq 'returned') {
            $sheet.Range("E$row").ClearFormats()
            $sheet.Range($at.formula_merge.Replace('{row}', [string]$row)).Merge()
        } else {
            $sheet.Range($at.text_merge.Replace('{row}', [string]$row)).Merge()
        }
    }
    $notes = $sheet.Range("D$($rows.call):I$($rows.provenance)")
    $notes.Font.Size = 9
    $notes.HorizontalAlignment = -4131
    $notes.VerticalAlignment = -4108
    $notes.WrapText = $true
    $sheet.Range("D$($rows.returned)").HorizontalAlignment = -4152
    $sheet.Range("D$($rows.returned)").Font.Size = 11
    # The case table: a heading like the Explorer's section titles, the column headings like its grid's.
    Copy-Format $explorer.Range('C5:D5') $sheet.Range("C$($header - 1)")
    foreach ($column in @('C', 'D', 'E', 'F', 'G', 'H', 'I', 'K', 'L', 'M', 'N', 'O')) {
        Copy-Format $explorer.Range('F19') $sheet.Range("$column$header")
        Copy-Format $explorer.Range('D21') $sheet.Range("${column}${first}:$column$last")
    }
    foreach ($row in $header..$last) {
        foreach ($merge in $at.table_merges) { $sheet.Range($merge.Replace('{row}', [string]$row)).Merge() }
    }
    $widths = 1..16 | ForEach-Object { $explorer.Columns.Item($_).ColumnWidth }
    for ($column = 1; $column -le 16; $column++) { $sheet.Columns.Item($column).ColumnWidth = $widths[$column - 1] }
    $sheet.Range("$($at.hidden_from):XFD").EntireColumn.Hidden = $true
    $sheet.Rows.Item("1:$last").RowHeight = 20.1
    foreach ($row in 1..20) { $sheet.Rows.Item($row).RowHeight = $explorer.Rows.Item($row).RowHeight }
    foreach ($pair in @(@('call', 26), @('status', 30), @('meaning', 30), @('action', 30), @('domain', 30),
                        @('converted', 30), @('sweep', 30), @('not_shown', 66), @('provenance', 40))) {
        $sheet.Rows.Item([int]$rows.($pair[0])).RowHeight = [double]$pair[1]
    }
    $sheet.Rows.Item($header).RowHeight = 30
    $sheet.Range("C${header}:O$header").WrapText = $true
    $table = $sheet.Range("C${first}:O$last")
    $table.Font.Size = 9
    $table.WrapText = $true
    $table.VerticalAlignment = -4108
    $sheet.Rows.Item("${first}:$last").RowHeight = 52
    # General keeps every digit a cell can show and needs no decimal separator. Through COM the format is
    # read in the language of Excel ("General", "Standard", ...): it is taken from a cell that has none.
    $general = [string]$sheet.Range('AZ1').NumberFormatLocal
    $sheet.Range($at.inputs).NumberFormatLocal = $general
    $sheet.Range("D$($rows.returned)").NumberFormatLocal = $general
    $sheet.Range("$($columns.returned)${first}:$($columns.expected)$last").NumberFormatLocal = $general
    $sheet.Range("$($columns.title)${first}:$($columns.call)$last").HorizontalAlignment = -4131
    $sheet.Range("$($columns.returned)${first}:$($columns.expected)$last").HorizontalAlignment = -4152
    $sheet.Range("$($columns.check)${first}:$($columns.check)$last").HorizontalAlignment = -4108
    $sheet.Range("$($columns.status)${first}:O$last").HorizontalAlignment = -4131
    $sheet.Tab.Color = $explorer.Tab.Color
    try { $excel.CutCopyMode = $false } catch { }
    $sheet.Range('C5').Value2 = [string]$explorer.Range('C5').Value2

    # The copied formats bring the Explorer's conditional rules, whose formulas read the Explorer's
    # inputs: remove every formula rule here; the sheet's own are added below. A rule on a cell's
    # value (the green OK) is kept.
    $rules = $sheet.Cells.FormatConditions
    for ($index = $rules.Count; $index -ge 1; $index--) {
        if ($rules.Item($index).Type -eq 2) { $rules.Item($index).Delete() }
    }

    Step 'names'
    $failed = @()
    foreach ($item in $plan.names) {
        try { [void]$book.Names.Add($item.name, (Localize-Formula $item.refers ([string]$excel.International(5)) $dec)) }
        catch { $failed += ($item.name + ' = ' + $item.refers + ' -> ' + $_.Exception.Message) }
    }
    if ($failed.Count -gt 0) { throw ("Names rejected:`n" + ($failed -join "`n")) }

    Step 'texts, the catalogue, defaults and validation'
    foreach ($text in $plan.texts) { Set-Value $sheet.Range($text.address) $text.value }
    $sheet.Range($at.function).Value2 = [string]$plan.defaults.function
    for ($slot = 0; $slot -lt $at.arguments.Count; $slot++) {
        $target = $sheet.Range([string]$at.arguments[$slot])
        $value = $plan.defaults.args[$slot]
        if ($null -eq $value) { $target.ClearContents() | Out-Null } else { Set-Value $target $value }
    }
    Set-Value $sheet.Range([string]$at.sweep[0]) $plan.defaults.sweep[0]
    Set-Value $sheet.Range([string]$at.sweep[1]) $plan.defaults.sweep[1]
    $choice = $sheet.Range($at.function)
    $choice.Validation.Delete()
    $choice.Validation.Add(3, 1, 1, $plan.validation.function_list)
    foreach ($entry in $plan.validation.custom.PSObject.Properties) {
        $cell = $sheet.Range($entry.Name)
        $cell.Validation.Delete()
        $cell.Validation.Add(7, 1, 1, [string]$entry.Value)
    }
    # The arguments take any value: their rule only shows the input message.
    foreach ($address in $plan.validation.input_only) {
        $cell = $sheet.Range([string]$address)
        $cell.Validation.Delete()
        $cell.Validation.Add(0, 1, 1)
    }
    foreach ($entry in $plan.messages.PSObject.Properties) {
        $rule = $sheet.Range($entry.Name).Validation
        $rule.InputTitle = [string]$entry.Value[0]
        $rule.InputMessage = [string]$entry.Value[1]
        $rule.ShowInput = $true
        if ([string]$entry.Value[2]) {
            $rule.ErrorTitle = [string]$entry.Value[2]
            $rule.ErrorMessage = [string]$entry.Value[3]
            $rule.ShowError = $true
        }
    }

    Step 'the formulas: the try-it results, the sweep, the case table'
    foreach ($item in $plan.formulas) {
        $target = $sheet.Range($item.address)
        Invoke-Excel { $excel.DisplayAlerts = $false; $target.Formula2 = [string]$item.formula }
    }
    Wait-Excel

    Step 'conditional formats of the status and of the checks'
    $status = $sheet.Range($at.status)
    $bad = $status.FormatConditions.Add(2, 0, $plan.conditional.status_bad)
    $bad.Font.Color = 0x0000C0
    $bad.Interior.Color = 0xCEC7FF
    $checks = $sheet.Range("$($columns.check)${first}:$($columns.check)$last")
    foreach ($rule in $plan.conditional.check_rules) {
        $cell = $checks.FormatConditions.Add(1, 3, '="' + [string]$rule.text + '"')
        $cell.Font.Color = [int]$rule.font
        $cell.Interior.Color = [int]$rule.fill
    }
    $inside = $table.Borders
    foreach ($edge in @(7, 8, 9, 10, 11, 12)) { $inside.Item($edge).LineStyle = 1; $inside.Item($edge).Color = 0xD9D9D9 }

    Step 'inputs unlocked, outputs locked'
    $sheet.Range($at.inputs).Locked = $false
    foreach ($address in $at.outputs) { $sheet.Range($address).Locked = $true }

    Step 'the chart'
    $top = $explorer.ChartObjects('ExplorerDensity')
    $bottom = $explorer.ChartObjects('ExplorerCDF')
    foreach ($spec in $plan.charts) {
        $item = $sheet.ChartObjects().Add($top.Left, $top.Top, $top.Width, $bottom.Top + $bottom.Height - $top.Top)
        $item.Name = $spec.name
        $chart = $item.Chart
        $chart.ChartType = 75
        # The helpers of the chart sit in hidden columns: with the default, Excel drops such a series.
        $chart.PlotVisibleOnly = $false
        while ($chart.SeriesCollection().Count -gt 0) { $chart.SeriesCollection(1).Delete() }
        $order = 1
        foreach ($series in $spec.series) {
            $created = $chart.SeriesCollection().NewSeries()
            $created.Formula = '=SERIES("' + $series.name + '",' + $prefix + $series.x + ',' + $prefix + $series.y + ',' + $order + ')'
            $order++
        }
        $chart.HasTitle = $true
        $chart.ChartTitle.Text = [string]$spec.title
        $chart.HasLegend = $true
        $item.Placement = 2
    }

    $excel.Calculation = -4105
    Invoke-Excel { $excel.Calculate() }
    Invoke-Excel { $excel.CalculateFull() }
    Wait-Excel
    $excel.ScreenUpdating = $true
    Step 'defaults and chart through the Diagnostics module'
    $processId = 0
    [void][EvolveNative]::GetWindowThreadProcessId([IntPtr][long]$excel.Hwnd, [ref]$processId)
    $watchdog = Start-Watchdog $processId 120
    [void]$excel.Run($prefix + 'ResetDiagnosticsInputs', $false)
    [void]$excel.Run($prefix + 'RefreshDiagnostics')
    Stop-Job $watchdog -ErrorAction SilentlyContinue
    Remove-Job $watchdog -Force -ErrorAction SilentlyContinue
    $sheet.Activate()
    $excel.ActiveWindow.DisplayGridlines = $false
    $sheet.Range('D6').Select()
    $excel.EnableEvents = $true
    [void]$excel.Run($prefix + 'RefreshRibbon')
    Step 'save'
    $explorer.Activate()
    $excel.CommandBars.ExecuteMso('FileSave')
    $sheet.ExportAsFixedFormat(0, (Join-Path $directory 'Diagnostics-preview.pdf'))
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    Write-Output ('Diagnostics evolved: ' + (Get-FileHash $path -Algorithm SHA256).Hash)
} catch {
    $note = Join-Path $env:TEMP 'diagnostics-watchdog.txt'
    if (Test-Path $note) { Get-Content $note | Write-Output }
    Write-Output $_.InvocationInfo.PositionMessage
    Write-Output $_.ScriptStackTrace
    throw
} finally {
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
