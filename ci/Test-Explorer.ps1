param([Parameter(Mandatory=$true)][string]$PackageDirectory)
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$fixture = Get-Content (Join-Path $PSScriptRoot '../examples/demo/explorer-references.json') -Raw | ConvertFrom-Json
$report = @{ status='FAIL'; sha256=(Get-FileHash $path -Algorithm SHA256).Hash; checks=@() }
$excel = $null
$book = $null
$other = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($path,0,$true)
    $sheet = $book.Worksheets.Item('Explorer')
    $prefix = "'" + $book.Name + "'!"
    $report.excel_version = $excel.Version
    $report.excel_build = $excel.Build
    $decimal = [string]$excel.International(3)
    if ($sheet.Range('D8').Text -ne ('1' + $decimal + '0000')) { throw 'Default sigma is formatted incorrectly' }
    if ($sheet.Range('G8').Text -ne ('0' + $decimal + '500000')) { throw 'Default CDF is formatted incorrectly' }
    if ($null -ne $sheet.UsedRange.Find('Under development')) { throw 'Placeholder text remains in Explorer' }
    foreach ($cell in $sheet.Range('A1:A7').Cells) {
        if ($cell.Interior.Color -ne $sheet.Range('A8').Interior.Color) { throw 'Navigation background differs in frozen/header rows' }
    }
    # Read the formula array directly; SpecialCells is unavailable on protected sheets.
    foreach ($formula in $sheet.UsedRange.Formula2) {
        if ([string]$formula -like '=*' -and [string]$formula -match '@|_xlfn.SINGLE') { throw "Implicit intersection remains: $formula" }
    }
    $validation = $sheet.Range('D6').Validation
    if ($validation.Type -ne 3 -or $validation.Formula1 -ne '=$Q$1:$Q$3') { throw 'Distribution validation is not a range-backed list' }
    $choices = @($sheet.Range('Q1:Q3').Cells | ForEach-Object { [string]$_.Value2 })
    if (($choices -join '|') -ne 'Normal|Exponential|Poisson') { throw 'Distribution list does not have three separate choices' }
    foreach ($choice in $choices) {
        $sheet.Range('D6').Value2 = $choice
        if (-not $validation.Value -or $sheet.Range('G5').Value2 -ne 'OK') { throw "Validation rejects $choice" }
    }
    $sheet.Range('D6').Value2 = 'Normal,Exponential,Poisson'
    if ($validation.Value) { throw 'Validation accepts concatenated family list' }
    $excel.Run($prefix + 'ResetExplorerInputs',$false)
    $report.navigation_background = 'PASS'
    $report.formula2_no_implicit_intersection = 'PASS'
    $report.range_backed_family_validation = 'PASS'
    $report.display_formats = 'PASS'
    if ($sheet.Range('D6:D11').Locked) { throw 'Explorer inputs are locked' }
    if (-not $sheet.Range('G7').Locked -or -not $sheet.ProtectContents) { throw 'Results are not protected' }
    foreach ($case in $fixture.cases) {
        $excel.EnableEvents = $false
        $sheet.Range('D6').Value2 = $case.family
        if ($null -ne $case.mean) { $sheet.Range('D7').Value2 = [double]$case.mean }
        if ($null -ne $case.sigma) { $sheet.Range('D8').Value2 = [double]$case.sigma }
        if ($null -ne $case.rate) { $sheet.Range('D9').Value2 = [double]$case.rate }
        if ($null -ne $case.poisson_mean) { $sheet.Range('D10').Value2 = [double]$case.poisson_mean }
        $sheet.Range('D11').Value2 = [double]$case.x
        $excel.EnableEvents = $true
        $excel.Run($prefix + 'RefreshExplorer')
        if ($sheet.Range('G5').Value2 -ne 'OK') { throw 'A valid case is rejected' }
        $pdf = [double]$sheet.Range('G7').Value2
        $cdf = [double]$sheet.Range('G8').Value2
        if ([Math]::Abs($pdf-$case.density) -gt $fixture.absolute_tolerance -or
            [Math]::Abs($cdf-$case.cdf) -gt $fixture.absolute_tolerance) {
            throw "Reference mismatch for $($case.family), x=$($case.x): $pdf / $cdf"
        }
        $expectedType = if ($case.family -eq 'Poisson') { 51 } else { 75 }
        if ($sheet.ChartObjects('ExplorerDensity').Chart.ChartType -ne $expectedType) { throw 'Wrong density chart type' }
        if (-not $sheet.ProtectContents -or -not $excel.EnableEvents) { throw 'Refresh leaked Excel state' }
        $report.checks += @{family=$case.family;x=$case.x;density=$pdf;cdf=$cdf;status='PASS'}
    }
    # Exercise the actual SheetChange event, not only the public refresh routine.
    $excel.Run($prefix + 'ResetExplorerInputs',$false)
    $sheet.Range('D8').Value2 = 2.0
    if ([Math]::Abs([double]$sheet.Range('G7').Value2-0.19947114020071635) -gt 1e-12) { throw 'Input event did not recalculate Normal density' }
    if ([double]$sheet.Range('F18').Value2 -ne -8.0 -or [double]$sheet.Range('F218').Value2 -ne 8.0) { throw 'Chart domain did not follow sigma' }
    $series = $sheet.ChartObjects('ExplorerDensity').Chart.SeriesCollection(1)
    $chartValues = @($series.Values)
    if ([Math]::Abs([double]$chartValues[100]-[double]$sheet.Range('G118').Value2) -gt 1e-12) { throw 'Chart values did not update' }
    # Invalid paste bypasses data validation but must still be visibly rejected.
    $sheet.Range('D8').Value2 = -1.0
    if ($sheet.Range('G5').Value2 -ne 'Check inputs') { throw 'Invalid sigma not rejected' }
    $other = $excel.Workbooks.Add()
    $other.Worksheets.Item(1).Range('D6').Value2 = 'sentinel'
    $other.Activate()
    $excel.Run($prefix + 'ResetExplorerInputs',$false)
    if ($other.Worksheets.Item(1).Range('D6').Value2 -ne 'sentinel') { throw 'Reset changed another workbook' }
    if ($sheet.Range('D6').Value2 -ne 'Normal' -or $sheet.Range('D8').Value2 -ne 1.0) { throw 'Reset did not restore defaults' }
    $report.changed_inputs_and_charts = 'PASS'
    $report.invalid_input = 'PASS'
    $report.reset_isolation = 'PASS'
    $report.status = 'PASS'
} catch {
    $report.error = $_.Exception.Message
    throw
} finally {
    if ($null -ne $other) { $other.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($other) }
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    $report | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 (Join-Path $directory 'explorer-report.json')
    $report | ConvertTo-Json -Depth 5 | Write-Output
}
