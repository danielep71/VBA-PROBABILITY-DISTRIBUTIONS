param([Parameter(Mandatory=$true)][string]$PackageDirectory)
# Applies tails-plan.json to the Tails sheet (Extreme Tails and Numerical
# Limits) of the package in desktop Excel, after the Explorer, Probabilities,
# Parameters and Comparison sheets: the sheet takes the formats and widths of
# the Explorer (the same roles: inputs, status, results, notes, headings), then
# the names, the stored scenarios and references, the spilled ladder (the
# library's direct result, the alternative, the digits and the statuses), the
# validation, the conditional formats and the two charts. The plan is made by
# tools/demobuild/tails_plan.py.
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$plan = Get-Content (Join-Path $directory 'tails-plan.json') -Raw | ConvertFrom-Json
$excel = $null
$book = $null
Remove-Item (Join-Path $env:TEMP 'tails-watchdog.txt') -ErrorAction SilentlyContinue
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
        $text | Out-File -FilePath (Join-Path $env:TEMP 'tails-watchdog.txt') -Encoding UTF8
        Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
    }
}

function Step([string]$Text) {
    if ($null -ne $excel) { $excel.DisplayAlerts = $false }
    Write-Output ('[' + (Get-Date -Format 'HH:mm:ss') + '] ' + $Text)
}

# Copies the formats (and merges) of one Explorer range onto a range of the
# Tails sheet.
function Copy-Format($Source, $Target) {
    # The Windows clipboard is shared with every program on the machine: another one can take it between the
    # copy and the paste, which Excel then refuses. Copy again and retry, up to five times.
    for ($attempt = 1; ; $attempt++) {
        try {
            [void]$Source.Copy()
            [void]$Target.PasteSpecial(-4122)
            return
        } catch {
            if ($attempt -ge 5) { throw }
            Write-Host "  the paste of formats was refused (attempt $attempt); copying again"
            Start-Sleep -Milliseconds 500
        }
    }
}

# Writes one column of the plan's hidden storage: a text column is formatted as text first, so a call
# kept as text ("=K_STATS_...") and a reference ("6.2E-16") stay what they are.
function Set-Column($Sheet, $Column) {
    $range = $Sheet.Range([string]$Column.address)
    $values = @($Column.values)
    $array = New-Object 'object[,]' $values.Count, 1
    if ($Column.kind -eq 'text') { $range.NumberFormat = '@' }
    for ($index = 0; $index -lt $values.Count; $index++) {
        if ($Column.kind -eq 'text') { $array[$index, 0] = [string]$values[$index] }
        else { $array[$index, 0] = [double]$values[$index] }
    }
    $range.Value2 = $array
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
    $prefix = "'" + $book.Name + "'!"

    Step 'clear the sheet and take the Explorer formats'
    foreach ($old in @($sheet.ChartObjects())) { $old.Delete() }
    $sheet.Columns.Hidden = $false
    $sheet.Range('A1:P80').UnMerge()
    $sheet.Range('A1:AZ400').ClearContents() | Out-Null
    $sheet.Range('B1:AZ400').ClearFormats() | Out-Null
    $sheet.Range('A1:Z80').Validation.Delete() | Out-Null
    Copy-Format $explorer.Range('A1:P20') $sheet.Range('A1')
    try { $excel.CutCopyMode = $false } catch { }
    # The inputs are D6:D8 here: the rows the Explorer keeps for its other inputs are spacers.
    foreach ($address in $at.spacers) {
        Copy-Format $explorer.Range('C17:D18') $sheet.Range($address)
    }
    # Rows 19 and 20 of the Explorer carry its grid headings and first grid row in F:I.
    $sheet.Range('F19:I20').ClearFormats()
    # The scenario: a label in C, its text over D:I.
    foreach ($property in $at.detail_rows.PSObject.Properties) {
        $row = [int]$property.Value
        Copy-Format $explorer.Range('C21:D21') $sheet.Range("C${row}:D$row")
        Copy-Format $explorer.Range('D21') $sheet.Range("E${row}:I$row")
        if ($property.Name -ne 'statuses') { $sheet.Range($at.detail_merge.Replace('{row}', [string]$row)).Merge() }
    }
    foreach ($row in 32..33) {
        Copy-Format $explorer.Range('C21:D21') $sheet.Range("C${row}:D$row")
        Copy-Format $explorer.Range('D21') $sheet.Range("E${row}:I$row")
    }
    $sheet.Range("C31:C33").Merge()
    $sheet.Range($at.statuses_merge).Merge()
    $detail = $sheet.Range('D21:I35')
    $detail.Font.Size = 9
    $detail.HorizontalAlignment = -4131
    $detail.VerticalAlignment = -4108
    $detail.WrapText = $true
    $sheet.Range('C31').VerticalAlignment = -4108
    # The ladder: a heading like the Explorer's section titles, the column headings like its grid's.
    Copy-Format $explorer.Range('C5:D5') $sheet.Range("C$($header - 1)")
    foreach ($column in @('C', 'D', 'F', 'G', 'H', 'I', 'K', 'L', 'M')) {
        Copy-Format $explorer.Range('F19') $sheet.Range("$column$header")
        if ($column -eq 'C') { Copy-Format $explorer.Range('C21') $sheet.Range("C${first}:C$last") }
        else { Copy-Format $explorer.Range('D21') $sheet.Range("${column}${first}:$column$last") }
    }
    $widths = 1..16 | ForEach-Object { $explorer.Columns.Item($_).ColumnWidth }
    for ($column = 1; $column -le 16; $column++) { $sheet.Columns.Item($column).ColumnWidth = $widths[$column - 1] }
    $sheet.Range("$($at.hidden_from):XFD").EntireColumn.Hidden = $true
    $sheet.Rows.Item("1:$last").RowHeight = 20.1
    foreach ($row in 1..20) { $sheet.Rows.Item($row).RowHeight = $explorer.Rows.Item($row).RowHeight }
    # The long texts of the scenario wrap over two or three lines.
    foreach ($pair in @(@('title', 26), @('parameters', 26), @('reference', 26), @('guarantee', 38), @('alternative_usable', 26),
                        @('direct_usable', 26), @('limit', 38), @('count', 26), @('provenance', 36))) {
        $sheet.Rows.Item([int]$at.detail_rows.($pair[0])).RowHeight = [double]$pair[1]
    }
    $sheet.Rows.Item($header).RowHeight = 36
    $sheet.Range("C${header}:M$header").WrapText = $true
    $table = $sheet.Range("C${first}:M$last")
    $table.Font.Size = 9
    $sheet.Range($at.values).NumberFormatLocal = '0' + $dec + '000000E+00'
    $sheet.Range($at.values).HorizontalAlignment = -4152
    $sheet.Range($at.digits).NumberFormatLocal = '0' + $dec + '0'
    $sheet.Range($at.depth).NumberFormatLocal = '0' + $dec + '0'
    $sheet.Range($at.digits).HorizontalAlignment = -4108
    $sheet.Range($at.depth).HorizontalAlignment = -4108
    $sheet.Range($at.statuses).HorizontalAlignment = -4108
    $sheet.Range($at.points).HorizontalAlignment = -4152
    $sheet.Range($at.reference_column).HorizontalAlignment = -4152
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

    Step 'the scenarios and their references, texts and validation'
    foreach ($column in $plan.columns) { Set-Column $sheet $column }
    foreach ($cell in $plan.numbers) { $sheet.Range($cell.address).Value2 = [double]$cell.value }
    foreach ($text in $plan.texts) { $sheet.Range($text.address).Value2 = [string]$text.value }
    $sheet.Range($at.scenario).Value2 = [string]$plan.defaults.scenario
    $sheet.Range($at.reference_grade).Value2 = [double]$plan.defaults.reference_grade
    $sheet.Range($at.report_grade).Value2 = [double]$plan.defaults.report_grade
    $scenario = $sheet.Range($at.scenario)
    $scenario.Validation.Delete()
    $scenario.Validation.Add(3, 1, 1, $plan.validation.scenario_list)
    foreach ($address in @($at.reference_grade, $at.report_grade)) {
        $cell = $sheet.Range($address)
        $cell.Validation.Delete()
        $cell.Validation.Add(2, 1, 1, [string]$plan.validation.digits[0], [string]$plan.validation.digits[1])
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

    Step 'the formulas: the ladder, the helpers and the scenario'
    foreach ($item in $plan.formulas) {
        $target = $sheet.Range($item.address)
        Invoke-Excel { $excel.DisplayAlerts = $false; $target.Formula2 = [string]$item.formula }
    }
    Wait-Excel

    Step 'conditional formats'
    $status = $sheet.Range($at.status)
    $bad = $status.FormatConditions.Add(2, 0, $plan.conditional.status_bad)
    $bad.Font.Color = 0x0000C0
    $bad.Interior.Color = 0xCEC7FF
    # The statuses: red where a value is wrong or died by cancellation, amber where it is the Double range.
    # A cell-value rule has no relative reference, so it does not depend on the active cell.
    foreach ($rule in $plan.conditional.status_rules) {
        $cell = $sheet.Range($at.statuses).FormatConditions.Add(1, 3, '="' + [string]$rule.text + '"')
        $cell.Font.Color = [int]$rule.font
        $cell.Font.Bold = $true
        $cell.Interior.Color = [int]$rule.fill
    }
    # The ladder's frame follows the number of points of the scenario.
    $inside = $sheet.Range("C${first}:M$last").FormatConditions.Add(2, 0, $plan.conditional.grid_inside)
    foreach ($edge in @(@(-4131, -4115), @(-4152, 1), @(-4160, 1), @(-4107, -4115))) {
        try { $inside.Borders.Item($edge[0]).LineStyle = $edge[1] } catch { $inside.Borders.Item($edge[0]).LineStyle = 1 }
    }

    Step 'inputs unlocked, outputs locked'
    $sheet.Range($at.inputs).Locked = $false
    foreach ($address in $at.outputs) { $sheet.Range($address).Locked = $true }

    Step 'the two charts'
    $places = @{ 'TailsDigits' = $explorer.ChartObjects('ExplorerDensity'); 'TailsValues' = $explorer.ChartObjects('ExplorerCDF') }
    foreach ($spec in $plan.charts) {
        foreach ($old in @($sheet.ChartObjects())) { if ($old.Name -eq $spec.name) { $old.Delete() } }
        $donor = $places[$spec.name]
        $item = $sheet.ChartObjects().Add($donor.Left, $donor.Top, $donor.Width, $donor.Height)
        $item.Name = $spec.name
        $chart = $item.Chart
        $chart.ChartType = 74
        # The data of the second chart sit in hidden columns: with the default, Excel drops such a series.
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
    Step 'defaults and charts through the Tails module'
    $processId = 0
    [void][EvolveNative]::GetWindowThreadProcessId([IntPtr][long]$excel.Hwnd, [ref]$processId)
    $watchdog = Start-Watchdog $processId 120
    [void]$excel.Run($prefix + 'ResetTailsInputs', $false)
    [void]$excel.Run($prefix + 'RefreshTails')
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
    $sheet.ExportAsFixedFormat(0, (Join-Path $directory 'Tails-preview.pdf'))
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    Write-Output ('Tails evolved: ' + (Get-FileHash $path -Algorithm SHA256).Hash)
} catch {
    $note = Join-Path $env:TEMP 'tails-watchdog.txt'
    if (Test-Path $note) { Get-Content $note | Write-Output }
    Write-Output $_.InvocationInfo.PositionMessage
    Write-Output $_.ScriptStackTrace
    throw
} finally {
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
