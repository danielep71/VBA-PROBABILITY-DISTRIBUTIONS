param([Parameter(Mandatory=$true)][string]$PackageDirectory)
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$copy = Join-Path $directory 'Excel-roundtrip.xlsm'
$report = @{ status = 'FAIL'; workbook = $path; sha256 = (Get-FileHash $path -Algorithm SHA256).Hash }
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $report.excel_version = $excel.Version
    $report.excel_build = $excel.Build
    # Normal open only: never request Excel's repair/extract-data modes.
    $book = $excel.Workbooks.Open($path, 0, $false)
    if ($book.Worksheets.Count -ne 9) { throw 'Expected nine worksheets' }
    if ($book.HasVBProject -ne $true) { throw 'VBA project missing' }
    $report.open = 'PASS'
    $book.SaveCopyAs($copy)
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    $book = $excel.Workbooks.Open($copy, 0, $true)
    if ($book.Worksheets.Count -ne 9) { throw 'Round-trip lost worksheets' }
    if ($book.HasVBProject -ne $true) { throw 'Round-trip lost VBA project' }
    $report.roundtrip = 'PASS'
    $report.status = 'PASS'
} catch {
    $report.error = $_.Exception.Message
    throw
} finally {
    if ($null -ne $book) {
        $book.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    }
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    $report | ConvertTo-Json | Set-Content -Encoding UTF8 (Join-Path $directory 'excel-open-report.json')
}
