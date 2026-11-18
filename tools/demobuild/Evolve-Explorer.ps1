param([Parameter(Mandatory=$true)][string]$PackageDirectory)
# Applies explorer-plan.json to the Explorer sheet of the package in desktop
# Excel: the family table, the workbook names, the spilled formulas, the
# validation and conditional formats, and the two charts. Nothing here draws a
# new layout: the sheet keeps the formats, merges, charts and navigation bar of
# the workbook it started from. The plan is made by tools/demobuild/explorer_plan.py.
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$plan = Get-Content (Join-Path $directory 'explorer-plan.json') -Raw | ConvertFrom-Json
$excel = $null
$book = $null

# Names.Add reads its RefersTo with English function names but the list
# separator of the user's regional settings (a semicolon in many locales).
function Localize-Formula([string]$Formula, [string]$Separator) {
    $builder = New-Object System.Text.StringBuilder
    $quoted = $false
    foreach ($character in $Formula.ToCharArray()) {
        if ($character -eq '"') { $quoted = -not $quoted }
        if ($character -eq ',' -and -not $quoted) { [void]$builder.Append($Separator) } else { [void]$builder.Append($character) }
    }
    return $builder.ToString()
}

# Excel rejects a COM call while it is busy (RPC_E_CALL_REJECTED), for
# instance while it computes a spilled grid it was just given; wait and retry,
# for up to five minutes, and say so in the log. The log goes to the host, not
# the output stream, so it never becomes part of the value returned.
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

# Excel turns alerts back on once a worksheet function of the VBA project has
# run, and a hidden confirmation would wait for ever; switch them off at each step.
function Step([string]$Text) {
    if ($null -ne $excel) { $excel.DisplayAlerts = $false }
    Write-Output ('[' + (Get-Date -Format 'HH:mm:ss') + '] ' + $Text)
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
    $sheet = $book.Worksheets.Item($plan.sheet)
    try { $sheet.Unprotect() } catch { }
    $excel.Calculation = -4135
    $first = [int]$plan.grid.first_row
    $last = [int]$plan.grid.last_row

    $dec = [string]$excel.International(3)
    $at = $plan.layout
    # The workbook this starts from has the user's layout already, with the
    # formulas of an earlier build in it: every output, rule and format the
    # plan owns is cleared here and written again, nothing is moved.
    Step 'clear the old result and grid formulas'
    foreach ($address in @($at.status, $at.macros, $at.quantile, $at.tail, $at.support)) { $sheet.Range($address).MergeArea.ClearContents() | Out-Null }
    $sheet.Range($at.results).Resize(3, 1).ClearContents() | Out-Null
    $sheet.Range("F${first}:I$last").ClearContents() | Out-Null
    $sheet.Range("F${first}:I$last").Validation.Delete()
    $sheet.Range("F${first}:I$last").FormatConditions.Delete()
    $sheet.Range('D7:D10').FormatConditions.Delete()
    # Keep the user's own "OK" format on the status cell; drop the plan's rules.
    $status = $sheet.Range($at.status)
    for ($index = $status.FormatConditions.Count; $index -ge 1; $index--) {
        if ($status.FormatConditions.Item($index).Type -eq 2) { $status.FormatConditions.Item($index).Delete() }
    }

    Step 'family table'
    foreach ($cell in $plan.table.cells) {
        $target = $sheet.Range($cell.address)
        if ($null -ne $cell.PSObject.Properties['formula']) { continue }
        if ($cell.value -is [string]) { $target.Value2 = [string]$cell.value }
        else { $target.Value2 = [double]$cell.value }
    }
    foreach ($item in $plan.names) { try { $book.Names.Item($item.name).Delete() } catch { } }
    # A name of an earlier plan that the sheet no longer uses.
    try { $book.Names.Item('ExplorerGridDiscrete').Delete() } catch { }
    $failed = @()
    foreach ($item in $plan.names) {
        try { [void]$book.Names.Add($item.name, (Localize-Formula $item.refers ([string]$excel.International(5)))) }
        catch { $failed += ($item.name + ' = ' + $item.refers + ' -> ' + $_.Exception.Message) }
    }
    if ($failed.Count -gt 0) { throw ("Names rejected:`n" + ($failed -join "`n")) }
    foreach ($cell in $plan.table.cells) {
        if ($null -ne $cell.PSObject.Properties['formula']) { $sheet.Range($cell.address).Formula2 = [string]$cell.formula }
    }

    Step 'texts and validation'
    foreach ($text in $plan.texts) { $sheet.Range($text.address).Value2 = [string]$text.value }
    $sheet.Range($at.probability).Value2 = [double]$plan.probability.default
    $sheet.Range($at.decimals).Value2 = [double]$plan.decimals.default
    $family = $sheet.Range($at.family)
    $family.Validation.Delete()
    $family.Validation.Add(3, 1, 1, $plan.validation.family_list)
    $family.Validation.ErrorTitle = 'Choose a listed distribution'
    $family.Validation.ErrorMessage = 'Pick a name from the list.'
    $family.Validation.ShowError = $true
    for ($row = 7; $row -le 10; $row++) {
        $cell = $sheet.Range("D$row")
        $cell.Validation.Delete()
        $cell.Validation.Add(7, 1, 1, $plan.validation.parameter.Replace('{slot}', [string]($row - 6)))
        $cell.Validation.ErrorTitle = 'Parameter out of range'
        $cell.Validation.ErrorMessage = 'The value is outside the limits of this parameter. The status line names the first one to check.'
        $cell.Validation.ShowError = $true
    }
    $limit = [double]$plan.validation.x_limit
    $bound = [double]$plan.validation.bound_limit
    $rules = @(@($at.x, -$limit, $limit), @($at.probability, 0, 1), @($at.lower, -$bound, $bound), @($at.upper, -$bound, $bound))
    foreach ($rule in $rules) {
        $cell = $sheet.Range($rule[0])
        $cell.Validation.Delete()
        $cell.Validation.Add(2, 1, 1, [string]$rule[1], [string]$rule[2])
        $cell.Validation.ShowError = $true
    }
    $sheet.Range($at.probability).Validation.ErrorTitle = 'Probability p'
    $sheet.Range($at.probability).Validation.ErrorMessage = 'Enter a probability from 0 to 1.'
    $points = $sheet.Range($at.points)
    $points.Validation.Delete()
    $points.Validation.Add(1, 1, 1, [string]$plan.validation.points[0], [string]$plan.validation.points[1])
    $points.Validation.ErrorTitle = 'Number of points'
    $points.Validation.ErrorMessage = 'Enter a whole number from 2 to ' + $plan.validation.points[1] + '.'
    $points.Validation.ShowError = $true
    $decimals = $sheet.Range($at.decimals)
    $decimals.Validation.Delete()
    $decimals.Validation.Add(3, 1, 1, $plan.validation.decimals_list)
    $decimals.Validation.ErrorTitle = 'Number of decimals'
    $decimals.Validation.ErrorMessage = 'Pick a number of decimals from 0 to ' + $plan.decimals.max + '.'
    $decimals.Validation.ShowError = $true

    Step 'result formulas and the spilled grid'
    $donor = $sheet.Range('F' + ($first + 1))
    $grid = $sheet.Range("F${first}:I$last")
    $grid.Font.Name = $donor.Font.Name
    $grid.Font.Size = $donor.Font.Size
    $grid.Font.Bold = $false
    $grid.Interior.Pattern = -4142
    $grid.VerticalAlignment = -4108
    $grid.Borders.LineStyle = -4142
    # NumberFormat is read with the regional decimal separator here, so write the local form.
    $grid.NumberFormatLocal = '0' + $dec + '000000'
    $grid.WrapText = $false
    $grid.Locked = $true
    $sheet.Rows.Item("${first}:$last").RowHeight = 20.1
    foreach ($item in $plan.formulas) {
        $target = $sheet.Range($item.address)
        Invoke-Excel { $target.Formula2 = [string]$item.formula }
    }
    Wait-Excel

    Step 'conditional formats'
    for ($row = 7; $row -le 10; $row++) {
        $unused = $sheet.Range("D$row").FormatConditions.Add(2, 0, $plan.conditional.unused_parameter.Replace('{slot}', [string]($row - 6)))
        $unused.Font.Color = 0x808080
        $unused.Interior.Color = 0xF2F2F2
    }
    $bad = $status.FormatConditions.Add(2, 0, $plan.conditional.status_bad)
    $bad.Font.Color = 0x0000C0
    $bad.Interior.Color = 0xCEC7FF
    Step 'grid borders'
    $inside = $grid.FormatConditions.Add(2, 0, $plan.conditional.grid_inside)
    foreach ($edge in @(@(-4131, -4115), @(-4152, 1), @(-4160, 1), @(-4107, -4115))) {
        try { $inside.Borders.Item($edge[0]).LineStyle = $edge[1] } catch { $inside.Borders.Item($edge[0]).LineStyle = 1 }
    }

    Step 'inputs unlocked, outputs locked'
    $sheet.Range($at.inputs).Locked = $false
    foreach ($address in $at.outputs) { $sheet.Range($address).Locked = $true }

    Step 'charts follow the spilled columns through their names'
    $prefix = "'" + $book.Name + "'!"
    foreach ($spec in @(@('ExplorerDensity', 'ExplorerPDF'), @('ExplorerCDF', 'ExplorerCDF'))) {
        $chart = $sheet.ChartObjects($spec[0]).Chart
        while ($chart.SeriesCollection().Count -gt 1) { $chart.SeriesCollection($chart.SeriesCollection().Count).Delete() }
        $chart.SeriesCollection(1).Formula = '=SERIES(,' + $prefix + 'ExplorerX,' + $prefix + $spec[1] + ',1)'
        $chart.HasLegend = $false
    }

    $excel.Calculation = -4105
    Invoke-Excel { $excel.Calculate() }
    Invoke-Excel { $excel.CalculateFull() }
    Wait-Excel
    $excel.ScreenUpdating = $true
    Step 'defaults and charts through the Explorer module'
    [void]$excel.Run($prefix + 'ResetExplorerInputs', $false)
    [void]$excel.Run($prefix + 'RefreshExplorer')
    $sheet.Activate()
    $sheet.Range('D6').Select()
    $excel.EnableEvents = $true
    [void]$excel.Run($prefix + 'RefreshRibbon')
    Step 'save'
    $excel.CommandBars.ExecuteMso('FileSave')
    $sheet.ExportAsFixedFormat(0, (Join-Path $directory 'Explorer-preview.pdf'))
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    Write-Output ('Explorer evolved: ' + (Get-FileHash $path -Algorithm SHA256).Hash)
} catch {
    Write-Output $_.InvocationInfo.PositionMessage
    Write-Output $_.ScriptStackTrace
    throw
} finally {
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
