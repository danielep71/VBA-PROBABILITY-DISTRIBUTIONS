param([Parameter(Mandatory=$true)][string]$PackageDirectory)
# Applies parameters-plan.json to the Parameters sheet of the package in
# desktop Excel, after Evolve-Explorer.ps1 has built the Explorer: the sheet
# takes the formats, widths and merges of the Explorer cell by cell (the same
# roles: inputs, status, results, notes), then the names, the spilled
# formulas, the validation, the conditional formats and the two charts. The
# plan is made by tools/demobuild/parameters_plan.py.
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$plan = Get-Content (Join-Path $directory 'parameters-plan.json') -Raw | ConvertFrom-Json
$excel = $null
$book = $null
Remove-Item (Join-Path $env:TEMP 'parameters-watchdog.txt') -ErrorAction SilentlyContinue
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
        $text | Out-File -FilePath (Join-Path $env:TEMP 'parameters-watchdog.txt') -Encoding UTF8
        Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
    }
}

function Step([string]$Text) {
    if ($null -ne $excel) { $excel.DisplayAlerts = $false }
    Write-Output ('[' + (Get-Date -Format 'HH:mm:ss') + '] ' + $Text)
}

# Copies the formats (and merges) of one Explorer range onto a range of the
# Parameters sheet.
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
    $last = [int]$plan.grid.last_row
    $dec = [string]$excel.International(3)
    $at = $plan.layout
    $headerRow = [int]$at.header_row
    $firstCurve = [int]$at.first_curve_row
    $lastCurve = $firstCurve + 3
    $prefix = "'" + $book.Name + "'!"

    Step 'clear the sheet and take the Explorer formats'
    $sheet.Columns.Hidden = $false
    $sheet.Range('A1:P40').UnMerge()
    $sheet.Range('A1:P40').ClearContents() | Out-Null
    $sheet.Range("Q1:BA$last").ClearContents() | Out-Null
    $sheet.Range('A1:P40').FormatConditions.Delete()
    Copy-Format $explorer.Range('A1:P20') $sheet.Range('A1')
    try { $excel.CutCopyMode = $false } catch { }
    # The inputs D6:D16, the status D19 and the macros D20 sit where the Explorer has them.
    Copy-Format $explorer.Range('F19:G19') $sheet.Range("C${headerRow}:D$headerRow")
    Copy-Format $explorer.Range('F19:I19') $sheet.Range("F${headerRow}:I$headerRow")
    foreach ($row in $firstCurve..$lastCurve) { Copy-Format $explorer.Range('C21:D21') $sheet.Range("C${row}:D$row") }
    Copy-Format $explorer.Range('D21') $sheet.Range("F${firstCurve}:I$lastCurve")
    Copy-Format $explorer.Range('C21:D21') $sheet.Range('C26:D26')
    Copy-Format $explorer.Range('C26:D27') $sheet.Range('C27:D28')
    $widths = 1..16 | ForEach-Object { $explorer.Columns.Item($_).ColumnWidth }
    for ($column = 1; $column -le 16; $column++) { $sheet.Columns.Item($column).ColumnWidth = $widths[$column - 1] }
    $sheet.Range('Q:XFD').EntireColumn.Hidden = $true
    $sheet.Rows.Item("1:$last").RowHeight = 20.1
    foreach ($row in 1..20) { $sheet.Rows.Item($row).RowHeight = $explorer.Rows.Item($row).RowHeight }
    $sheet.Tab.Color = $explorer.Tab.Color
    try { $excel.CutCopyMode = $false } catch { }
    $effect = $sheet.Range($at.effect)
    $effect.Merge()
    $effect.WrapText = $true
    $effect.VerticalAlignment = -4160
    $effect.Font.Size = 9
    $sheet.Range('C5').Value2 = [string]$explorer.Range('C5').Value2

    Step 'names'
    $failed = @()
    foreach ($item in $plan.names) {
        try { [void]$book.Names.Add($item.name, (Localize-Formula $item.refers ([string]$excel.International(5)) $dec)) }
        catch { $failed += ($item.name + ' = ' + $item.refers + ' -> ' + $_.Exception.Message) }
    }
    if ($failed.Count -gt 0) { throw ("Names rejected:`n" + ($failed -join "`n")) }

    Step 'the family table, texts and validation'
    foreach ($cell in $plan.table.cells) {
        $target = $sheet.Range($cell.address)
        if ($cell.value -is [string]) { $target.Value2 = [string]$cell.value }
        else { $target.Value2 = [double]$cell.value }
    }
    foreach ($text in $plan.texts) { $sheet.Range($text.address).Value2 = [string]$text.value }
    $sheet.Range($at.family).Value2 = 'Normal'
    $sheet.Range($at.decimals).Value2 = [double]$plan.decimals.default
    $family = $sheet.Range($at.family)
    $family.Validation.Delete()
    $family.Validation.Add(3, 1, 1, $plan.validation.family_list)
    $family.Validation.ShowError = $true
    for ($row = 7; $row -le 10; $row++) {
        $cell = $sheet.Range("D$row")
        $cell.Validation.Delete()
        $cell.Validation.Add(7, 1, 1, $plan.validation.parameter.Replace('{slot}', [string]($row - 6)))
        $cell.Validation.ErrorTitle = 'Parameter out of range'
        $cell.Validation.ErrorMessage = 'The value is outside the limits of this parameter. The status line names the first one to check.'
        $cell.Validation.ShowError = $true
    }
    $vary = $sheet.Range($at.vary)
    $vary.Validation.Delete()
    $vary.Validation.Add(3, 1, 1, $plan.validation.vary_list)
    $vary.Validation.ShowError = $true
    for ($k = 1; $k -le 4; $k++) {
        $cell = $sheet.Range($at.curve_values[$k - 1])
        $cell.Validation.Delete()
        $cell.Validation.Add(7, 1, 1, $plan.validation.curve.Replace('{k}', [string]$k))
        $cell.Validation.ShowError = $true
    }
    $decimals = $sheet.Range($at.decimals)
    $decimals.Validation.Delete()
    $decimals.Validation.Add(3, 1, 1, $plan.validation.decimals_list)
    $decimals.Validation.ShowError = $true
    foreach ($entry in $plan.messages.PSObject.Properties) {
        $rule = $sheet.Range($entry.Name).Validation
        $rule.InputTitle = [string]$entry.Value[0]
        $rule.InputMessage = [string]$entry.Value[1]
        $rule.ErrorTitle = [string]$entry.Value[2]
        $rule.ErrorMessage = [string]$entry.Value[3]
        $rule.ShowInput = $true
        $rule.ShowError = $true
    }

    Step 'the formulas of the curves and the table'
    $sheet.Range("F${firstCurve}:I$lastCurve").NumberFormatLocal = '0' + $dec + '000000'
    $sheet.Range("F${firstCurve}:I$lastCurve").HorizontalAlignment = -4152
    foreach ($item in $plan.formulas) {
        $target = $sheet.Range($item.address)
        Invoke-Excel { $excel.DisplayAlerts = $false; $target.Formula2 = [string]$item.formula }
    }
    Wait-Excel

    Step 'conditional formats'
    for ($row = 7; $row -le 10; $row++) {
        $unused = $sheet.Range("D$row").FormatConditions.Add(2, 0, $plan.conditional.unused_parameter.Replace('{slot}', [string]($row - 6)))
        $unused.Font.Color = 0x808080
        $unused.Interior.Color = 0xF2F2F2
    }
    $status = $sheet.Range($at.status)
    $bad = $status.FormatConditions.Add(2, 0, $plan.conditional.status_bad)
    $bad.Font.Color = 0x0000C0
    $bad.Interior.Color = 0xCEC7FF

    Step 'inputs unlocked, outputs locked'
    $sheet.Range($at.inputs).Locked = $false
    foreach ($address in $at.outputs) { $sheet.Range($address).Locked = $true }

    Step 'the two charts'
    $places = @{ 'ParameterEffectsDensity' = $explorer.ChartObjects('ExplorerDensity'); 'ParameterEffectsCdf' = $explorer.ChartObjects('ExplorerCDF') }
    foreach ($spec in $plan.charts) {
        foreach ($old in @($sheet.ChartObjects())) { if ($old.Name -eq $spec.name) { $old.Delete() } }
        $donor = $places[$spec.name]
        $item = $sheet.ChartObjects().Add($donor.Left, $donor.Top, $donor.Width, $donor.Height)
        $item.Name = $spec.name
        $chart = $item.Chart
        $chart.ChartType = 4
        # The data sit in hidden columns: with the default, Excel drops a series whose cells are all hidden.
        $chart.PlotVisibleOnly = $false
        while ($chart.SeriesCollection().Count -gt 0) { $chart.SeriesCollection(1).Delete() }
        $order = 1
        foreach ($series in $spec.series) {
            $created = $chart.SeriesCollection().NewSeries()
            $nameCell = $plan.sheet + '!' + $sheet.Range($series.name).Address()
            $created.Formula = '=SERIES(' + $nameCell + ',' + $prefix + $series.x + ',' + $prefix + $series.y + ',' + $order + ')'
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
    Step 'defaults and charts through the Parameter Effects module'
    $processId = 0
    [void][EvolveNative]::GetWindowThreadProcessId([IntPtr][long]$excel.Hwnd, [ref]$processId)
    $watchdog = Start-Watchdog $processId 120
    [void]$excel.Run($prefix + 'ResetParameterEffectsInputs', $false)
    [void]$excel.Run($prefix + 'RefreshParameterEffects')
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
    $sheet.ExportAsFixedFormat(0, (Join-Path $directory 'Parameters-preview.pdf'))
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    Write-Output ('Parameters evolved: ' + (Get-FileHash $path -Algorithm SHA256).Hash)
} catch {
    $note = Join-Path $env:TEMP 'parameters-watchdog.txt'
    if (Test-Path $note) { Get-Content $note | Write-Output }
    Write-Output $_.InvocationInfo.PositionMessage
    Write-Output $_.ScriptStackTrace
    throw
} finally {
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
