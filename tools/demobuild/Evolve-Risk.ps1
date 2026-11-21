param([Parameter(Mandatory=$true)][string]$PackageDirectory)
# Applies risk-plan.json to the Risk sheet (value at risk, expected
# shortfall, tails and default counts) of the package in desktop Excel, after
# the sheets before it: the sheet takes the formats and widths of the Explorer
# (the same roles: inputs, status, results, notes, headings), then the names,
# the results with the formula of each written beside it, the spilled table
# of losses, the validation, the conditional format of the status and the two
# charts. The plan is made by tools/demobuild/risk_plan.py.
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$plan = Get-Content (Join-Path $directory 'risk-plan.json') -Raw | ConvertFrom-Json
$excel = $null
$book = $null
Remove-Item (Join-Path $env:TEMP 'risk-watchdog.txt') -ErrorAction SilentlyContinue
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
        $text | Out-File -FilePath (Join-Path $env:TEMP 'risk-watchdog.txt') -Encoding UTF8
        Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
    }
}

function Step([string]$Text) {
    if ($null -ne $excel) { $excel.DisplayAlerts = $false }
    Write-Output ('[' + (Get-Date -Format 'HH:mm:ss') + '] ' + $Text)
}

# Copies the formats (and merges) of one Explorer range onto a range of the
# Risk sheet.
function Copy-Format($Source, $Target) {
    [void]$Source.Copy()
    [void]$Target.PasteSpecial(-4122)
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
    $prefix = "'" + $book.Name + "'!"
    $textRows = @('reading', 'assumptions', 'excel_build', 'provenance')

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
    # The inputs are D6:D15 here: the rows the Explorer keeps for its other inputs are spacers.
    foreach ($address in $at.spacers) {
        $target = $sheet.Range($address)
        Copy-Format $explorer.Range('C17').Resize($target.Rows.Count, 2) $target
    }
    # Rows 19 and 20 of the Explorer carry its grid headings and first grid row in F:I.
    $sheet.Range('F19:I20').ClearFormats()
    # The results: a label in C, the value in D, the formula of the value as text over F:I. The reading, the
    # assumptions, the build and the provenance are text over D:I.
    foreach ($property in $rows.PSObject.Properties) {
        $row = [int]$property.Value
        Copy-Format $explorer.Range('C21:D21') $sheet.Range("C${row}:D$row")
        Copy-Format $explorer.Range('D21') $sheet.Range("F${row}:I$row")
        if ($textRows -contains $property.Name) {
            Copy-Format $explorer.Range('D21') $sheet.Range("E${row}:E$row")
            if ($property.Name -ne 'assumptions') { $sheet.Range($at.text_merge.Replace('{row}', [string]$row)).Merge() }
        } else {
            $sheet.Range($at.formula_merge.Replace('{row}', [string]$row)).Merge()
        }
    }
    $assumptions = [int]$rows.assumptions
    Copy-Format $explorer.Range('C21:D21') $sheet.Range("C$($assumptions + 1):D$($assumptions + 1)")
    Copy-Format $explorer.Range('D21') $sheet.Range("E$($assumptions + 1):I$($assumptions + 1)")
    $sheet.Range("C${assumptions}:C$($assumptions + 1)").Merge()
    $sheet.Range($at.assumptions_merge).Merge()
    $sheet.Range("C$assumptions").VerticalAlignment = -4108
    $texts = $sheet.Range("F$($rows.mean_loss):I$($rows.default_exceed)")
    $texts.Font.Size = 9
    $texts.HorizontalAlignment = -4131
    $texts.VerticalAlignment = -4108
    $texts.WrapText = $true
    $notes = $sheet.Range("D$($rows.reading):I$($rows.provenance)")
    $notes.Font.Size = 9
    $notes.HorizontalAlignment = -4131
    $notes.VerticalAlignment = -4108
    $notes.WrapText = $true
    # The table: a heading like the Explorer's section titles, the column headings like its grid's.
    Copy-Format $explorer.Range('C5:D5') $sheet.Range("C$($header - 1)")
    foreach ($column in @('C', 'D', 'F', 'G', 'H', 'I')) {
        Copy-Format $explorer.Range('F19') $sheet.Range("$column$header")
        Copy-Format $explorer.Range('D21') $sheet.Range("${column}${first}:$column$last")
    }
    $widths = 1..16 | ForEach-Object { $explorer.Columns.Item($_).ColumnWidth }
    for ($column = 1; $column -le 16; $column++) { $sheet.Columns.Item($column).ColumnWidth = $widths[$column - 1] }
    $sheet.Range("$($at.hidden_from):XFD").EntireColumn.Hidden = $true
    $sheet.Rows.Item("1:$last").RowHeight = 20.1
    foreach ($row in 1..20) { $sheet.Rows.Item($row).RowHeight = $explorer.Rows.Item($row).RowHeight }
    foreach ($pair in @(@('normal_es', 30), @('t_es', 40), @('t_exceed', 26), @('reading', 52), @('assumptions', 52), @('provenance', 26))) {
        $sheet.Rows.Item([int]$rows.($pair[0])).RowHeight = [double]$pair[1]
    }
    $sheet.Rows.Item($assumptions + 1).RowHeight = 52
    $sheet.Rows.Item($header).RowHeight = 30
    $sheet.Range("C${header}:I$header").WrapText = $true
    $table = $sheet.Range("C${first}:I$last")
    $table.Font.Size = 9
    # General keeps every digit a cell can show and needs no decimal separator. Through COM the format is
    # read in the language of Excel ("General", "Standard", ...): it is taken from a cell that has none.
    $general = [string]$sheet.Range('AZ1').NumberFormatLocal
    $sheet.Range($at.probabilities).NumberFormatLocal = $general
    $sheet.Range($at.losses).NumberFormatLocal = $general
    $sheet.Range($at.results).NumberFormatLocal = $general
    $sheet.Range($at.inputs).NumberFormatLocal = $general
    $sheet.Range("D$($rows.default_quantile)").NumberFormatLocal = '0'
    $table.HorizontalAlignment = -4152
    $sheet.Tab.Color = $explorer.Tab.Color
    try { $excel.CutCopyMode = $false } catch { }
    $sheet.Range('C5').Value2 = [string]$explorer.Range('C5').Value2

    # The copied formats bring the Explorer's conditional rules, whose formulas read the Explorer's
    # inputs: remove every formula rule here; the sheet's own is added below. A rule on a cell's
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

    Step 'texts, defaults and validation'
    foreach ($text in $plan.texts) { $sheet.Range($text.address).Value2 = [string]$text.value }
    foreach ($entry in $at.input_cells.PSObject.Properties) {
        $sheet.Range([string]$entry.Value).Value2 = [double]$plan.defaults.($entry.Name)
    }
    foreach ($entry in $plan.validation.custom.PSObject.Properties) {
        $cell = $sheet.Range($entry.Name)
        $cell.Validation.Delete()
        $cell.Validation.Add(7, 1, 1, [string]$entry.Value)
    }
    foreach ($entry in $plan.messages.PSObject.Properties) {
        $rule = $sheet.Range($entry.Name).Validation
        $rule.InputTitle = [string]$entry.Value[0]
        $rule.InputMessage = [string]$entry.Value[1]
        $rule.ErrorTitle = [string]$entry.Value[2]
        $rule.ErrorMessage = [string]$entry.Value[3]
        $rule.ShowInput = $true
        $rule.ShowError = $true
    }

    Step 'the formulas: the results, their formulas as text, the table and the helpers'
    foreach ($item in $plan.formulas) {
        $target = $sheet.Range($item.address)
        Invoke-Excel { $excel.DisplayAlerts = $false; $target.Formula2 = [string]$item.formula }
    }
    Wait-Excel

    Step 'conditional format of the status'
    $status = $sheet.Range($at.status)
    $bad = $status.FormatConditions.Add(2, 0, $plan.conditional.status_bad)
    $bad.Font.Color = 0x0000C0
    $bad.Interior.Color = 0xCEC7FF
    $inside = $table.Borders
    foreach ($edge in @(7, 8, 9, 10, 11, 12)) { $inside.Item($edge).LineStyle = 1; $inside.Item($edge).Color = 0xD9D9D9 }

    Step 'inputs unlocked, outputs locked'
    $sheet.Range($at.inputs).Locked = $false
    foreach ($address in $at.outputs) { $sheet.Range($address).Locked = $true }

    Step 'the two charts'
    $places = @{ 'RiskDensity' = $explorer.ChartObjects('ExplorerDensity'); 'RiskTail' = $explorer.ChartObjects('ExplorerCDF') }
    foreach ($spec in $plan.charts) {
        $donor = $places[$spec.name]
        $item = $sheet.ChartObjects().Add($donor.Left, $donor.Top, $donor.Width, $donor.Height)
        $item.Name = $spec.name
        $chart = $item.Chart
        $chart.ChartType = 75
        # The helpers of the charts sit in hidden columns: with the default, Excel drops such a series.
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
    Step 'defaults and charts through the Risk module'
    $processId = 0
    [void][EvolveNative]::GetWindowThreadProcessId([IntPtr][long]$excel.Hwnd, [ref]$processId)
    $watchdog = Start-Watchdog $processId 120
    [void]$excel.Run($prefix + 'ResetRiskInputs', $false)
    [void]$excel.Run($prefix + 'RefreshRisk')
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
    $sheet.ExportAsFixedFormat(0, (Join-Path $directory 'Risk-preview.pdf'))
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    Write-Output ('Risk evolved: ' + (Get-FileHash $path -Algorithm SHA256).Hash)
} catch {
    $note = Join-Path $env:TEMP 'risk-watchdog.txt'
    if (Test-Path $note) { Get-Content $note | Write-Output }
    Write-Output $_.InvocationInfo.PositionMessage
    Write-Output $_.ScriptStackTrace
    throw
} finally {
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
