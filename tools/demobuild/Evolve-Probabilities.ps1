param([Parameter(Mandatory=$true)][string]$PackageDirectory)
# Applies probabilities-plan.json to the Probabilities sheet of the package in
# desktop Excel, after Evolve-Explorer.ps1 has built the Explorer: the sheet
# takes the formats, widths and merges of the Explorer cell by cell (the same
# roles: inputs, status, results, notes, grid), then the names, the spilled
# formulas, the validation, the conditional formats and the shaded chart. The
# plan is made by tools/demobuild/probabilities_plan.py.
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$plan = Get-Content (Join-Path $directory 'probabilities-plan.json') -Raw | ConvertFrom-Json
$excel = $null
$book = $null
Remove-Item (Join-Path $env:TEMP 'probabilities-watchdog.txt') -ErrorAction SilentlyContinue
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
        $text | Out-File -FilePath (Join-Path $env:TEMP 'probabilities-watchdog.txt') -Encoding UTF8
        Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
    }
}

function Step([string]$Text) {
    if ($null -ne $excel) { $excel.DisplayAlerts = $false }
    Write-Output ('[' + (Get-Date -Format 'HH:mm:ss') + '] ' + $Text)
}

# Copies the formats (and merges) of one Explorer range onto a range of the
# Probabilities sheet.
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
    $first = [int]$plan.grid.first_row
    $last = [int]$plan.grid.last_row
    $dec = [string]$excel.International(3)
    $at = $plan.layout
    $prefix = "'" + $book.Name + "'!"

    Step 'clear the sheet and take the Explorer formats'
    $sheet.Columns.Hidden = $false
    $sheet.Range('A1:P1100').UnMerge()
    $sheet.Range('A1:P1100').ClearContents() | Out-Null
    $sheet.Range("F${first}:I1100").FormatConditions.Delete()
    $sheet.Range('D6:D18').FormatConditions.Delete()
    Copy-Format $explorer.Range('A1:P18') $sheet.Range('A1')
    try { $excel.CutCopyMode = $false } catch { }
    # The inputs are D6:D14 here: the rows the Explorer keeps for its other inputs are spacers.
    Copy-Format $explorer.Range('C17:D18') $sheet.Range('C15:D16')
    Copy-Format $explorer.Range('C19:D20') $sheet.Range('C17:D18')
    foreach ($row in 20..28) { Copy-Format $explorer.Range('C21:D21') $sheet.Range("C${row}:D$row") }
    Copy-Format $explorer.Range('C26:D27') $sheet.Range('C30:D31')
    Copy-Format $explorer.Range('F19:I19') $sheet.Range('F19:I19')
    $widths = 1..16 | ForEach-Object { $explorer.Columns.Item($_).ColumnWidth }
    for ($column = 1; $column -le 16; $column++) { $sheet.Columns.Item($column).ColumnWidth = $widths[$column - 1] }
    $sheet.Range('Q:XFD').EntireColumn.Hidden = $true
    $sheet.Rows.Item("1:$last").RowHeight = 20.1
    foreach ($row in 1..18) { $sheet.Rows.Item($row).RowHeight = $explorer.Rows.Item($row).RowHeight }
    $sheet.Rows.Item(19).RowHeight = $explorer.Rows.Item(19).RowHeight
    $sheet.Tab.Color = $explorer.Tab.Color
    try { $excel.CutCopyMode = $false } catch { }
    # The plain-text titles and the "Inputs" heading of the Explorer.
    $sheet.Range('C5').Value2 = [string]$explorer.Range('C5').Value2

    # The copied formats bring the Explorer's conditional rules, whose formulas read the Explorer's
    # inputs (ExplorerParameterUnused*, ExplorerStatusBad, the grid borders): remove every formula
    # rule here; the sheet's own are added below. A rule on a cell's value (the green OK) is kept.
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

    Step 'texts and validation'
    foreach ($text in $plan.texts) { $sheet.Range($text.address).Value2 = [string]$text.value }
    $sheet.Range($at.probability).Value2 = [double]$plan.probability.default
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
    $limit = [double]$plan.validation.point_limit
    foreach ($address in @($at.lower, $at.upper)) {
        $cell = $sheet.Range($address)
        $cell.Validation.Delete()
        $cell.Validation.Add(2, 1, 1, [string](-$limit), [string]$limit)
        $cell.Validation.ShowError = $true
    }
    $cell = $sheet.Range($at.probability)
    $cell.Validation.Delete()
    $cell.Validation.Add(2, 1, 1, '0', '1')
    $cell.Validation.ShowError = $true
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

    Step 'result formulas and the spilled grid'
    $grid = $sheet.Range("F${first}:I$last")
    $donor = $explorer.Range('F21')
    $grid.Font.Name = $donor.Font.Name
    $grid.Font.Size = $donor.Font.Size
    $grid.Font.Bold = $false
    $grid.Interior.Pattern = -4142
    $grid.VerticalAlignment = -4108
    $grid.Borders.LineStyle = -4142
    $grid.NumberFormatLocal = '0' + $dec + '000000'
    $grid.WrapText = $false
    $sheet.Range("D20:D27").NumberFormatLocal = '0' + $dec + '000000'
    $sheet.Range('D28').HorizontalAlignment = -4131
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
    $inside = $grid.FormatConditions.Add(2, 0, $plan.conditional.grid_inside)
    foreach ($edge in @(@(-4131, -4115), @(-4152, 1), @(-4160, 1), @(-4107, -4115))) {
        try { $inside.Borders.Item($edge[0]).LineStyle = $edge[1] } catch { $inside.Borders.Item($edge[0]).LineStyle = 1 }
    }

    Step 'inputs unlocked, outputs locked'
    $sheet.Range($at.inputs).Locked = $false
    foreach ($address in $at.outputs) { $sheet.Range($address).Locked = $true }
    $sheet.Range("F${first}:I$last").Locked = $true

    Step 'the shaded chart'
    foreach ($old in @($sheet.ChartObjects())) { if ($old.Name -eq $plan.chart.name) { $old.Delete() } }
    $density = $explorer.ChartObjects('ExplorerDensity')
    $second = $explorer.ChartObjects('ExplorerCDF')
    $left = $density.Left
    $top = $density.Top
    $width = $density.Width
    $height = ($second.Top + $second.Height) - $density.Top
    $item = $sheet.ChartObjects().Add($left, $top, $width, $height)
    $item.Name = $plan.chart.name
    $chart = $item.Chart
    $chart.ChartType = 4
    while ($chart.SeriesCollection().Count -gt 0) { $chart.SeriesCollection(1).Delete() }
    $order = 1
    foreach ($series in $plan.chart.series) {
        $created = $chart.SeriesCollection().NewSeries()
        $created.Formula = '=SERIES("' + $series.name + '",' + $prefix + $series.x + ',' + $prefix + $series.y + ',' + $order + ')'
        $order++
    }
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = 'Density f(x)'
    $chart.HasLegend = $true
    $item.Placement = 2

    $excel.Calculation = -4105
    Invoke-Excel { $excel.Calculate() }
    Invoke-Excel { $excel.CalculateFull() }
    Wait-Excel
    $excel.ScreenUpdating = $true
    Step 'defaults and chart through the Probabilities module'
    $processId = 0
    [void][EvolveNative]::GetWindowThreadProcessId([IntPtr][long]$excel.Hwnd, [ref]$processId)
    $watchdog = Start-Watchdog $processId 120
    [void]$excel.Run($prefix + 'ResetProbabilitiesInputs', $false)
    [void]$excel.Run($prefix + 'RefreshProbabilities')
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
    $sheet.ExportAsFixedFormat(0, (Join-Path $directory 'Probabilities-preview.pdf'))
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    Write-Output ('Probabilities evolved: ' + (Get-FileHash $path -Algorithm SHA256).Hash)
} catch {
    $note = Join-Path $env:TEMP 'probabilities-watchdog.txt'
    if (Test-Path $note) { Get-Content $note | Write-Output }
    Write-Output $_.InvocationInfo.PositionMessage
    Write-Output $_.ScriptStackTrace
    throw
} finally {
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
