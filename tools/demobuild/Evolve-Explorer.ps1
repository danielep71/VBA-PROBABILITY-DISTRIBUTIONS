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

# Excel turns alerts back on once a worksheet function of the VBA project has
# run, and a hidden confirmation would wait for ever; switch them off at each step.
# Excel rejects a COM call while it is busy (RPC_E_CALL_REJECTED); wait and retry.
function Invoke-Excel([scriptblock]$Action) {
    for ($attempt = 1; ; $attempt++) {
        try { return & $Action }
        catch {
            $busy = $_.Exception.HResult -eq -2147418111 -or $_.Exception.Message -match '0x80010001'
            if (-not $busy -or $attempt -ge 40) { throw }
            Start-Sleep -Milliseconds 500
        }
    }
}

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
    Step 'clear the old result and grid formulas'
    $sheet.Range('G7:G9').ClearContents() | Out-Null
    $sheet.Range('G14').ClearContents() | Out-Null
    $sheet.Range("F${first}:J$last").ClearContents() | Out-Null
    $sheet.Range("F${first}:J$last").Validation.Delete()
    $sheet.Range("F${first}:J$last").FormatConditions.Delete()

    # The workbook this starts from (v2) has the plot window in G10:G12, the
    # decimals in G13, the right tail in G14 and the support text in F15:H16.
    # Move that block down one row, inside F:H only, to make room for the
    # quantile in row 10; the grid stays at row 19.
    Step 'room for the quantile and the probability input'
    [void]$sheet.Range('F10:H16').Cut($sheet.Range('F11'))
    [void]$sheet.Range('F9:H9').Copy($sheet.Range('F10'))
    [void]$sheet.Range('C11:D11').Copy($sheet.Range('C12'))
    $sheet.Range('G10').ClearContents() | Out-Null
    $sheet.Range("J18:J$last").Clear() | Out-Null
    [void]$sheet.Range('G18').Copy($sheet.Range('H18'))
    $support = $sheet.Range('F16:H17')
    $support.UnMerge()
    $sheet.Range('F17:H17').ClearContents() | Out-Null
    $support.Merge()
    $support.WrapText = $true
    $support.VerticalAlignment = -4160

    Step 'family table'
    foreach ($cell in $plan.table.cells) {
        $target = $sheet.Range($cell.address)
        if ($null -ne $cell.PSObject.Properties['formula']) { continue }
        if ($cell.value -is [string]) { $target.Value2 = [string]$cell.value }
        else { $target.Value2 = [double]$cell.value }
    }
    foreach ($item in $plan.names) { try { $book.Names.Item($item.name).Delete() } catch { } }
    # A name of the earlier plan that the sheet no longer uses.
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
    $sheet.Range('D12').Value2 = [double]$plan.probability.default
    $sheet.Range('G14').Value2 = [double]$plan.decimals.default
    $family = $sheet.Range('D6')
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
    $rules = @(@('D11', -$limit, $limit), @('D12', 0, 1), @('G11', -$bound, $bound), @('G12', -$bound, $bound))
    foreach ($rule in $rules) {
        $cell = $sheet.Range($rule[0])
        $cell.Validation.Delete()
        $cell.Validation.Add(2, 1, 1, [string]$rule[1], [string]$rule[2])
        $cell.Validation.ShowError = $true
    }
    $sheet.Range('D12').Validation.ErrorTitle = 'Probability p'
    $sheet.Range('D12').Validation.ErrorMessage = 'Enter a probability from 0 to 1.'
    $points = $sheet.Range('G13')
    $points.Validation.Delete()
    $points.Validation.Add(1, 1, 1, [string]$plan.validation.points[0], [string]$plan.validation.points[1])
    $points.Validation.ErrorTitle = 'Number of points'
    $points.Validation.ErrorMessage = 'Enter a whole number from 2 to ' + $plan.validation.points[1] + '.'
    $points.Validation.ShowError = $true
    $decimals = $sheet.Range('G14')
    $decimals.Validation.Delete()
    $decimals.Validation.Add(3, 1, 1, $plan.validation.decimals_list)
    $decimals.Validation.ErrorTitle = 'Number of decimals'
    $decimals.Validation.ErrorMessage = 'Pick a number of decimals from 0 to ' + $plan.decimals.max + '.'
    $decimals.Validation.ShowError = $true

    Step 'result formulas and the spilled grid'
    $donor = $sheet.Range('F20')
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
    foreach ($address in @('G5', 'G6', 'C7', 'C8', 'C9', 'C10', 'F16', 'F19', 'G19', 'H19', 'I19', 'G7', 'G10', 'G15')) {
        $item = $plan.formulas | Where-Object { $_.address -eq $address }
        if ($null -eq $item) { throw "The plan has no formula for $address" }
        Invoke-Excel { $sheet.Range($address).Formula2 = [string]$item.formula }
    }

    Step 'conditional formats'
    for ($row = 7; $row -le 10; $row++) {
        $unused = $sheet.Range("D$row").FormatConditions.Add(2, 0, $plan.conditional.unused_parameter.Replace('{slot}', [string]($row - 6)))
        $unused.Font.Color = 0x808080
        $unused.Interior.Color = 0xF2F2F2
    }
    $bad = $sheet.Range('G5').FormatConditions.Add(2, 0, $plan.conditional.status_bad)
    $bad.Font.Color = 0x0000C0
    $bad.Interior.Color = 0xCEC7FF
    Step 'grid borders'
    $inside = $grid.FormatConditions.Add(2, 0, $plan.conditional.grid_inside)
    foreach ($edge in @(@(-4131, -4115), @(-4152, 1), @(-4160, 1), @(-4107, -4115))) {
        try { $inside.Borders.Item($edge[0]).LineStyle = $edge[1] } catch { $inside.Borders.Item($edge[0]).LineStyle = 1 }
    }

    Step 'inputs unlocked, outputs locked'
    $sheet.Range('D6:D12').Locked = $false
    $sheet.Range('G11:G14').Locked = $false
    $sheet.Range('G5:G10').Locked = $true
    $sheet.Range('G15').Locked = $true

    Step 'charts follow the spilled columns through their names'
    $prefix = "'" + $book.Name + "'!"
    foreach ($spec in @(@('ExplorerDensity', 'ExplorerPDF'), @('ExplorerCDF', 'ExplorerCDF'))) {
        $chart = $sheet.ChartObjects($spec[0]).Chart
        while ($chart.SeriesCollection().Count -gt 1) { $chart.SeriesCollection($chart.SeriesCollection().Count).Delete() }
        $chart.SeriesCollection(1).Formula = '=SERIES(,' + $prefix + 'ExplorerX,' + $prefix + $spec[1] + ',1)'
        $chart.HasLegend = $false
    }

    $excel.Calculation = -4105
    $excel.Calculate()
    $excel.CalculateFull()
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
