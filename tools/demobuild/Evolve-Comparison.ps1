param([Parameter(Mandatory=$true)][string]$PackageDirectory)
# Applies comparison-plan.json to the Comparison sheet (Excel vs VBA vs Python)
# of the package in desktop Excel, after the Explorer, Probabilities and
# Parameters sheets: the sheet takes the formats and widths of the Explorer
# (the same roles: inputs, status, results, notes, headings), then the names,
# the stored values of the cases, the live Excel and library formulas, the
# digits and grades, the validation and the bar chart of correct digits. The
# plan is made by tools/demobuild/comparison_plan.py.
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$plan = Get-Content (Join-Path $directory 'comparison-plan.json') -Raw | ConvertFrom-Json
$excel = $null
$book = $null
Remove-Item (Join-Path $env:TEMP 'comparison-watchdog.txt') -ErrorAction SilentlyContinue
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
        $text | Out-File -FilePath (Join-Path $env:TEMP 'comparison-watchdog.txt') -Encoding UTF8
        Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
    }
}

function Step([string]$Text) {
    if ($null -ne $excel) { $excel.DisplayAlerts = $false }
    Write-Output ('[' + (Get-Date -Format 'HH:mm:ss') + '] ' + $Text)
}

# Copies the formats (and merges) of one Explorer range onto a range of the
# Comparison sheet.
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
    $prefix = "'" + $book.Name + "'!"

    Step 'clear the sheet and take the Explorer formats'
    foreach ($old in @($sheet.ChartObjects())) { $old.Delete() }
    $sheet.Columns.Hidden = $false
    $sheet.Range('A1:Z200').UnMerge()
    $sheet.Range('A1:Z200').ClearContents() | Out-Null
    $sheet.Range('B1:Z200').ClearFormats() | Out-Null
    $sheet.Range('A1:Z200').Validation.Delete() | Out-Null
    Copy-Format $explorer.Range('A1:P20') $sheet.Range('A1')
    try { $excel.CutCopyMode = $false } catch { }
    # The inputs are D6:D9 here: the rows the Explorer keeps for its other inputs are spacers.
    foreach ($address in $at.spacers) {
        $source = if ($sheet.Range($address).Rows.Count -eq 2) { 'C17:D18' } else { 'C17:D17' }
        Copy-Format $explorer.Range($source) $sheet.Range($address)
    }
    # Rows 19 and 20 of the Explorer carry its grid headings and first grid row in F:I.
    $sheet.Range('F19:I20').ClearFormats()
    # The selected case: a label in C, its value over D:I.
    foreach ($property in $at.detail_rows.PSObject.Properties) {
        $row = [int]$property.Value
        Copy-Format $explorer.Range('C21:D21') $sheet.Range("C${row}:D$row")
        Copy-Format $explorer.Range('D21') $sheet.Range("E${row}:I$row")
        $sheet.Range($at.detail_merge.Replace('{row}', [string]$row)).Merge()
    }
    Copy-Format $explorer.Range('C21:D21') $sheet.Range('C33:D33')
    Copy-Format $explorer.Range('D21') $sheet.Range('E33:I33')
    $sheet.Range('D32:I32').UnMerge()
    $sheet.Range($at.why_merge).Merge()
    $sheet.Range('C32:C33').Merge()
    $detail = $sheet.Range('D21:I35')
    $detail.Font.Size = 9
    $detail.HorizontalAlignment = -4131
    $detail.VerticalAlignment = -4108
    $sheet.Range($at.why_merge).WrapText = $true
    $sheet.Range('D35:I35').WrapText = $true
    $sheet.Range('C32').VerticalAlignment = -4108
    # The table: a heading like the Explorer's section titles, the column headings like its grid's.
    Copy-Format $explorer.Range('C5:D5') $sheet.Range("C$($header - 1)")
    foreach ($column in @('C', 'D', 'F', 'G', 'H', 'I', 'K', 'L', 'M', 'N', 'O')) {
        Copy-Format $explorer.Range('F19') $sheet.Range("$column$header")
        if ($column -eq 'C') { Copy-Format $explorer.Range('C21') $sheet.Range("C${first}:C$last") }
        else { Copy-Format $explorer.Range('D21') $sheet.Range("${column}${first}:$column$last") }
    }
    $widths = 1..16 | ForEach-Object { $explorer.Columns.Item($_).ColumnWidth }
    for ($column = 1; $column -le 16; $column++) { $sheet.Columns.Item($column).ColumnWidth = $widths[$column - 1] }
    $sheet.Range('Q:XFD').EntireColumn.Hidden = $true
    $sheet.Rows.Item("1:$last").RowHeight = 20.1
    foreach ($row in 1..20) { $sheet.Rows.Item($row).RowHeight = $explorer.Rows.Item($row).RowHeight }
    $sheet.Rows.Item(35).RowHeight = 36
    $sheet.Rows.Item($header).RowHeight = 30
    $sheet.Range("C${header}:O$header").WrapText = $true
    $table = $sheet.Range("C${first}:O$last")
    $table.Font.Size = 9
    # A union address takes the regional list separator: each block is formatted on its own. The
    # formats are written in the regional language (NumberFormat through COM is read that way too).
    foreach ($address in ([string]$at.values).Split(',')) {
        $sheet.Range($address).NumberFormatLocal = '0' + $dec + '000000E+00'
        $sheet.Range($address).HorizontalAlignment = -4152
    }
    foreach ($address in ([string]$at.digits).Split(',')) {
        $sheet.Range($address).NumberFormatLocal = '0' + $dec + '0'
        $sheet.Range($address).HorizontalAlignment = -4108
    }
    # The values of the selected case keep every digit a Double shows; the thresholds are whole numbers.
    foreach ($key in @('excel_value', 'library_value', 'scipy_value')) {
        $row = [int]$at.detail_rows.$key
        $sheet.Range("D${row}:I$row").NumberFormatLocal = '0' + $dec + '000000000000000E+00'
    }
    $sheet.Range("$($at.reference_grade):$($at.report_grade)").NumberFormatLocal = '0'
    $sheet.Range($at.grades).HorizontalAlignment = -4108
    $sheet.Range("C${first}:C$last").HorizontalAlignment = -4131
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

    Step 'the cases, texts and validation'
    foreach ($cell in $plan.cells) {
        $target = $sheet.Range($cell.address)
        if ($cell.value -is [string]) { $target.Value2 = [string]$cell.value }
        else { $target.Value2 = [double]$cell.value }
    }
    foreach ($text in $plan.texts) { $sheet.Range($text.address).Value2 = [string]$text.value }
    $sheet.Range($at.case).Value2 = [string]$plan.defaults.case
    $sheet.Range($at.reference_grade).Value2 = [double]$plan.defaults.reference_grade
    $sheet.Range($at.report_grade).Value2 = [double]$plan.defaults.report_grade
    $sheet.Range($at.python).ClearContents() | Out-Null
    $case = $sheet.Range($at.case)
    $case.Validation.Delete()
    $case.Validation.Add(3, 1, 1, $plan.validation.case_list)
    foreach ($address in @($at.reference_grade, $at.report_grade)) {
        $cell = $sheet.Range($address)
        $cell.Validation.Delete()
        $cell.Validation.Add(2, 1, 1, [string]$plan.validation.digits[0], [string]$plan.validation.digits[1])
    }
    # The =PY() cell takes any formula: its rule only shows the instructions.
    $python = $sheet.Range($at.python)
    $python.Validation.Delete()
    $python.Validation.Add(0)
    foreach ($entry in $plan.messages.PSObject.Properties) {
        $rule = $sheet.Range($entry.Name).Validation
        $rule.InputTitle = [string]$entry.Value[0]
        $rule.InputMessage = [string]$entry.Value[1]
        $rule.ShowInput = $true
        if (-not [string]::IsNullOrEmpty([string]$entry.Value[2])) {
            $rule.ErrorTitle = [string]$entry.Value[2]
            $rule.ErrorMessage = [string]$entry.Value[3]
            $rule.ShowError = $true
        }
    }

    Step 'the formulas: Excel, the library, the digits and the grades'
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

    Step 'inputs unlocked, outputs locked'
    $sheet.Range($at.inputs).Locked = $false
    foreach ($address in $at.outputs) { $sheet.Range($address).Locked = $true }

    Step 'the chart of correct digits'
    $place = $sheet.Range($at.chart)
    $item = $sheet.ChartObjects().Add($place.Left, $place.Top, $place.Width, $place.Height)
    $item.Name = $plan.chart.name
    $chart = $item.Chart
    $chart.ChartType = 57
    while ($chart.SeriesCollection().Count -gt 0) { $chart.SeriesCollection(1).Delete() }
    $order = 1
    foreach ($series in $plan.chart.series) {
        $created = $chart.SeriesCollection().NewSeries()
        $nameCell = $plan.sheet + '!' + $sheet.Range($series.name).Address()
        $created.Formula = '=SERIES(' + $nameCell + ',' + $prefix + $plan.chart.categories + ',' + $prefix + $series.values + ',' + $order + ')'
        $order++
    }
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = [string]$plan.chart.title
    $chart.HasLegend = $true
    $item.Placement = 2

    $excel.Calculation = -4105
    Invoke-Excel { $excel.Calculate() }
    Invoke-Excel { $excel.CalculateFull() }
    Wait-Excel
    $excel.ScreenUpdating = $true
    Step 'defaults and chart through the Comparison module'
    $processId = 0
    [void][EvolveNative]::GetWindowThreadProcessId([IntPtr][long]$excel.Hwnd, [ref]$processId)
    $watchdog = Start-Watchdog $processId 120
    [void]$excel.Run($prefix + 'ResetComparisonInputs', $false)
    [void]$excel.Run($prefix + 'RefreshComparison')
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
    $sheet.ExportAsFixedFormat(0, (Join-Path $directory 'Comparison-preview.pdf'))
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    Write-Output ('Comparison evolved: ' + (Get-FileHash $path -Algorithm SHA256).Hash)
} catch {
    $note = Join-Path $env:TEMP 'comparison-watchdog.txt'
    if (Test-Path $note) { Get-Content $note | Write-Output }
    Write-Output $_.InvocationInfo.PositionMessage
    Write-Output $_.ScriptStackTrace
    throw
} finally {
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
