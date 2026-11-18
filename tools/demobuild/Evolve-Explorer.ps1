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

function Step([string]$Text) { Write-Output ('[' + (Get-Date -Format 'HH:mm:ss') + '] ' + $Text) }

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($path, 0, $false)
    try { $book.Unprotect() } catch { }
    $sheet = $book.Worksheets.Item($plan.sheet)
    try { $sheet.Unprotect() } catch { }
    $excel.Calculation = -4135
    $first = [int]$plan.grid.first_row
    $last = [int]$plan.grid.last_row

    Step 'clear the old fixed grid and result formulas'
    $sheet.Range('G7:G9').ClearContents() | Out-Null
    $sheet.Range('G13').ClearContents() | Out-Null
    $sheet.Range("F${first}:H$last").ClearContents() | Out-Null
    $sheet.Range("F${first}:H$last").Validation.Delete()
    $sheet.Range("F${first}:H$last").FormatConditions.Delete()

    Step 'family table'
    foreach ($cell in $plan.table.cells) {
        $target = $sheet.Range($cell.address)
        if ($null -ne $cell.PSObject.Properties['formula']) { continue }
        if ($cell.value -is [string]) { $target.Value2 = [string]$cell.value }
        else { $target.Value2 = [double]$cell.value }
    }
    foreach ($item in $plan.names) { try { $book.Names.Item($item.name).Delete() } catch { } }
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
    $sheet.Range('F14:H17').ClearContents() | Out-Null
    $sheet.Range('F14:H17').Merge()
    $sheet.Range('F14').WrapText = $true
    $sheet.Range('F14').VerticalAlignment = -4160
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
    $rules = @(@('D11', -$limit, $limit), @('G10', -$bound, $bound), @('G11', -$bound, $bound))
    foreach ($rule in $rules) {
        $cell = $sheet.Range($rule[0])
        $cell.Validation.Delete()
        $cell.Validation.Add(2, 1, 1, [string]$rule[1], [string]$rule[2])
        $cell.Validation.ShowError = $true
    }
    $points = $sheet.Range('G12')
    $points.Validation.Delete()
    $points.Validation.Add(1, 1, 1, [string]$plan.validation.points[0], [string]$plan.validation.points[1])
    $points.Validation.ErrorTitle = 'Number of points'
    $points.Validation.ErrorMessage = 'Enter a whole number from 2 to ' + $plan.validation.points[1] + '.'
    $points.Validation.ShowError = $true

    Step 'result formulas and the spilled grid'
    $donor = $sheet.Range('F20')
    $body = $sheet.Range("F${first}:H$last")
    $body.Font.Name = $donor.Font.Name
    $body.Font.Size = $donor.Font.Size
    $body.Font.Bold = $false
    $body.Interior.Pattern = -4142
    $body.VerticalAlignment = -4108
    $sheet.Range("F${first}:H$last").Borders.LineStyle = -4142
    # NumberFormat is read with the regional decimal separator here, so write the local form.
    $sheet.Range("F${first}:H$last").NumberFormatLocal = '0' + [string]$excel.International(3) + '000000'
    $sheet.Range("F${first}:H$last").WrapText = $false
    $sheet.Range("F${first}:H$last").Locked = $true
    $sheet.Rows.Item("${first}:$last").RowHeight = 20.1
    foreach ($item in $plan.formulas) {
        $address = $item.address
        if ($address -in @('G5', 'F14', 'G6', 'C7', 'C8', 'C9', 'C10')) { $sheet.Range($address).Formula2 = [string]$item.formula }
    }
    foreach ($address in @('F19', 'G19', 'H19', 'G7', 'G13')) {
        $item = $plan.formulas | Where-Object { $_.address -eq $address }
        $sheet.Range($address).Formula2 = [string]$item.formula
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
    $inside = $sheet.Range("F${first}:H$last").FormatConditions.Add(2, 0, $plan.conditional.grid_inside)
    foreach ($edge in @(@(-4131, -4115), @(-4152, 1), @(-4160, 1), @(-4107, -4115))) {
        try { $inside.Borders.Item($edge[0]).LineStyle = $edge[1] } catch { $inside.Borders.Item($edge[0]).LineStyle = 1 }
    }
    $integer = $sheet.Range("F${first}:F$last").FormatConditions.Add(2, 0, $plan.conditional.grid_discrete)
    $integer.NumberFormat = '0'

    Step 'inputs unlocked, outputs locked'
    $sheet.Range('D6:D11').Locked = $false
    $sheet.Range('G10:G12').Locked = $false
    $sheet.Range('G5:G9').Locked = $true
    $sheet.Range('G13').Locked = $true

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
