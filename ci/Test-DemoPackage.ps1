param([Parameter(Mandatory=$true)][string]$PackageDirectory)
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$copy = Join-Path $directory 'Excel-roundtrip.xlsm'
$report = @{ status = 'FAIL'; workbook = $path; sha256 = (Get-FileHash $path -Algorithm SHA256).Hash }
$excel = $null
$book = $null
$other = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $report.excel_version = $excel.Version
    $report.excel_build = $excel.Build
    # Normal open only: never request Excel's repair/extract-data modes.
    $book = $excel.Workbooks.Open($path, 0, $false)
    if ($book.Worksheets.Count -ne 10) { throw 'Expected ten worksheets' }
    if ($book.HasVBProject -ne $true) { throw 'VBA project missing' }
    $report.open = 'PASS'
    $book.SaveCopyAs($copy)
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    $book = $excel.Workbooks.Open($copy, 0, $true)
    if ($book.Worksheets.Count -ne 10) { throw 'Round-trip lost worksheets' }
    if ($book.HasVBProject -ne $true) { throw 'Round-trip lost VBA project' }
    $report.roundtrip = 'PASS'
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    # Macro execution is confined to this owned Excel process; Trust Center is unchanged.
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($path, 0, $false)
    $prefix = "'" + $book.Name + "'!"
    $value = $excel.Run($prefix + 'K_STATS_NormalStandard_Cumulative', 0.0)
    if ([Math]::Abs([double]$value - 0.5) -gt 1e-15) { throw 'Embedded VBA smoke result differs from 0.5' }
    $report.embedded_vba = 'PASS'
    $sheetMap = Get-Content (Join-Path $PSScriptRoot '..\examples\demo\sheets.json') -Raw | ConvertFrom-Json
    foreach ($section in $sheetMap.sections) {
        foreach ($entry in $section.sheets) {
            $control = 'pdGo' + $entry.code.Substring(2)
            $actual = $excel.Run($prefix + 'RibbonSheet', $control)
            if ($actual -ne $entry.code) { throw "Wrong Ribbon mapping: $control" }
            $target = $book.Names.Item('nav_' + $entry.code).RefersToRange
            if ($target.Worksheet.CodeName -ne $entry.code) { throw "Wrong navigation name: $control" }
        }
    }
    if ($excel.Run($prefix + 'RibbonSheet', 'unknown-control') -ne '') { throw 'Unknown control has a destination' }
    $report.navigation_targets = 'PASS'
    $other = $excel.Workbooks.Add()
    $other.Worksheets.Item(1).Range('A1').Value2 = 'isolation sentinel'
    $other.Activate()
    $savedScreen = $excel.ScreenUpdating
    $savedFlag = $book.Saved
    $excel.Run($prefix + 'ShowNavigationBar', $false)
    foreach ($sheet in $book.Worksheets) {
        if (-not $sheet.Columns.Item(1).Hidden) { throw "Navigation column still visible: $($sheet.Name)" }
        foreach ($name in @('NavBarTop','NavBar')) {
            if ($sheet.Shapes.Item($name).Visible -ne 0) { throw "Navigation shape still visible: $name" }
        }
        if (-not $sheet.ProtectContents) { throw 'Navigation toggle removed sheet protection' }
    }
    if ($excel.ScreenUpdating -ne $savedScreen -or $book.Saved -ne $savedFlag) { throw 'Navigation toggle leaked state' }
    if ($other.Worksheets.Item(1).Columns.Item(1).Hidden) { throw 'Navigation toggle changed another workbook' }
    if ($other.Worksheets.Item(1).Range('A1').Value2 -ne 'isolation sentinel') { throw 'Other workbook changed' }
    $report.hide_and_isolation = 'PASS'
    $excel.Run($prefix + 'RecalculateWorkbook')
    $book.Save()
    if (-not $book.Worksheets.Item(1).Columns.Item(1).Hidden) { throw 'Save did not restore session-hidden bar' }
    $report.save_events = 'PASS'
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($path, 0, $true)
    foreach ($sheet in $book.Worksheets) {
        if ($sheet.Columns.Item(1).Hidden) { throw 'Saved file hides its navigation with macros disabled' }
        foreach ($name in @('NavBarTop','NavBar')) {
            if ($sheet.Shapes.Item($name).Visible -eq 0) { throw 'Saved file hides a navigation group' }
        }
    }
    $report.saved_navigation = 'PASS'
    $report.scope = 'Embedded VBA, navigation mapping, toggle, isolation, save events; Ribbon clicks remain manual.'
    $report.status = 'PASS'
} catch {
    $report.error = $_.Exception.Message
    throw
} finally {
    if ($null -ne $book) {
        $book.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    }
    if ($null -ne $other) {
        $other.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($other)
    }
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    $report | ConvertTo-Json | Set-Content -Encoding UTF8 (Join-Path $directory 'excel-open-report.json')
}
