param([Parameter(Mandatory=$true)][string]$PackageDirectory)
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$root = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$map = Get-Content (Join-Path $root 'examples/demo/explorer-capabilities.json') -Raw | ConvertFrom-Json
$excel = $null
$book = $null
function DistributionFormula([string]$Operation, [string]$Point) {
    $expression = 'NA()'
    foreach ($family in @($map.families)[2..0]) {
        $call = $family.$Operation + '(' + $Point + ',' + $family.arguments + ')'
        $expression = 'IF($D$6="' + $family.name + '",' + $call + ',' + $expression + ')'
    }
    return '=IF($G$5="OK",' + $expression + ',NA())'
}
# A grid row outside the plotted domain (Poisson beyond its last count) holds
# #N/A in column F; give the same #N/A, not the #VALUE! of a call on an error.
function GridFormula([string]$Operation, [string]$Point) {
    return '=IF(ISNUMBER(' + $Point + '),' + (DistributionFormula $Operation $Point).Substring(1) + ',NA())'
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $decimal = [string]$excel.International(3)
    $book = $excel.Workbooks.Open($path, 0, $false)
    $book.Unprotect()
    $sheet = $book.Worksheets.Item('Explorer')
    $sheet.Unprotect()
    # Temporary style donor; deleted before saving. No independent visual theme.
    $sheet.Copy([Type]::Missing, $book.Worksheets.Item($book.Worksheets.Count))
    $donor = $book.Worksheets.Item($book.Worksheets.Count)
    # Keep the template navigation background, including the frozen rows.
    $sheet.Range('B1:P218').Clear()
    $sheet.Columns.Item('C').ColumnWidth = 26
    $sheet.Columns.Item('D').ColumnWidth = 18
    $sheet.Columns.Item('E').ColumnWidth = 3
    $sheet.Columns.Item('F:H').Hidden = $false
    $sheet.Columns.Item('F:H').ColumnWidth = 17
    $sheet.Columns.Item('I').Hidden = $false
    $sheet.Columns.Item('I').ColumnWidth = 3
    $sheet.Columns.Item('J:P').Hidden = $false
    $sheet.Columns.Item('J:P').ColumnWidth = 12
    $sheet.Rows.Item('1:218').RowHeight = 16
    $donor.Range('B1:D3').Copy($sheet.Range('B1'))
    $sheet.Range('B1:P1').Merge()
    $sheet.Range('B2:P2').Merge()
    $sheet.Range('C3:P3').Merge()
    $sheet.Range('C3').Value2 = 'Choose a family and edit the highlighted inputs. Formulas use the embedded VBA API; macros are required. Other families are not included in this first Explorer.'
    $sheet.Range('C3').WrapText = $true
    $sheet.Rows.Item(1).RowHeight = 24
    $sheet.Rows.Item(2).RowHeight = 26.1
    $sheet.Rows.Item(3).RowHeight = 34
    foreach ($address in @('C5:D5','F5:H5','F17:H17','C14:D14')) {
        $donor.Range('C5').Copy($sheet.Range($address))
        $sheet.Range($address).ClearContents()
    }
    $sheet.Range('C5').Value2 = 'Editable inputs'
    $labels = @('Distribution','Normal mean (x units)','Normal standard deviation','Exponential rate (1/x)','Poisson mean (count)','Evaluation point x / k')
    for ($i=0; $i -lt $labels.Count; $i++) {
        $row = $i + 6
        $donor.Range('C6').Copy($sheet.Range("C$row"))
        $donor.Range('D6').Copy($sheet.Range("D$row"))
        $sheet.Range("C$row").Value2 = $labels[$i]
    }
    $sheet.Range('C6:D11').WrapText = $true
    $sheet.Rows.Item('6:11').RowHeight = 28
    $sheet.Range('D6:D11').Locked = $false
    # K-SACCR input palette; preserve the template's typeface and borders.
    $sheet.Range('D6:D11').Interior.Color = 0xCCF2FF
    $sheet.Range('D6:D11').NumberFormatLocal = '0' + $decimal + '0000'
    # Range-backed lists work with every Excel regional list separator.
    $sheet.Range('Q1').Value2 = 'Normal'
    $sheet.Range('Q2').Value2 = 'Exponential'
    $sheet.Range('Q3').Value2 = 'Poisson'
    $sheet.Columns.Item('Q').Hidden = $true
    $sheet.Range('D6').Validation.Delete()
    $sheet.Range('D6').Validation.Add(3, 1, 1, '=$Q$1:$Q$3')
    $sheet.Range('D6').Validation.ErrorTitle = 'Choose a listed family'
    $sheet.Range('D6').Validation.ShowError = $true
    $limits = @(@('D7',-1000000,1000000),@('D8',0.001,10000),@('D9',0.001,1000),@('D10',0.01,100),@('D11',-10000000,10000000))
    foreach ($limit in $limits) {
        $inputCell = $sheet.Range($limit[0])
        $inputCell.Validation.Add(2, 1, 1, [double]$limit[1], [double]$limit[2])
        $inputCell.Validation.ShowError = $true
    }
    $donor.Range('D6').Copy($sheet.Range('F6:H15'))
    $sheet.Range('F6:H15').ClearContents()
    $sheet.Range('G5:H5').Merge()
    $sheet.Range('G5').Formula2 = '=IF(IFERROR(AND(ISNUMBER(D11),ABS(D11)<=10000000,OR(AND(D6="Normal",ISNUMBER(D7),ABS(D7)<=1000000,ISNUMBER(D8),D8>=0.001,D8<=10000),AND(D6="Exponential",ISNUMBER(D9),D9>=0.001,D9<=1000),AND(D6="Poisson",ISNUMBER(D10),D10>=0.01,D10<=100))),FALSE),"OK","Check inputs")'
    $sheet.Range('F5').Value2 = 'Input status'
    # With macros off every result is #NAME? and the charts do not follow the
    # inputs. Say so on the sheet: the formula depends on the six inputs, so it
    # is evaluated again at the first edit, which is when the results break.
    $sheet.Range('F6').Value2 = 'Macros'
    $sheet.Range('G6:H6').Merge()
    $sheet.Range('G6').Formula2 = '=IF(ISERROR(K_STATS_NormalStandard_Cumulative(0)+0*COUNTA($D$6:$D$11)),"DISABLED","Enabled")'
    # A line in the gap above the tables, empty while macros run. Its red bold
    # font is static: no conditional format is needed.
    $sheet.Range('F4:H4').Merge()
    $sheet.Range('F4').Formula2 = '=IF($G$6="Enabled","","MACROS ARE DISABLED: the results are not live")'
    $sheet.Range('F4').Font.Bold = $true
    $sheet.Range('F4').Font.Color = 0x0000C0
    $sheet.Range('F7').Value2 = 'PDF / PMF at x'
    $sheet.Range('F8').Value2 = 'CDF at x'
    $sheet.Range('F9').Value2 = 'P(X > x)'
    $sheet.Range('F10').Value2 = 'Plot lower bound'
    $sheet.Range('F11').Value2 = 'Plot upper bound'
    $sheet.Range('F12').Value2 = 'Right tail omitted'
    $sheet.Range('G7').Formula2 = DistributionFormula 'density' '$D$11'
    $sheet.Range('G8').Formula2 = DistributionFormula 'cdf' '$D$11'
    $sheet.Range('G9').Formula2 = DistributionFormula 'survival' '$D$11'
    $sheet.Range('G10').Formula2 = '=IF($G$5="OK",IF(D6="Normal",D7-4*D8,0),NA())'
    $sheet.Range('G11').Formula2 = '=IF($G$5="OK",IF(D6="Normal",D7+4*D8,IF(D6="Exponential",8/D9,ROUNDUP(D10+8*SQRT(D10)+10,0))),NA())'
    $sheet.Range('G12').Formula2 = DistributionFormula 'survival' '$G$11'
    $sheet.Range('G7:G12').NumberFormatLocal = '0' + $decimal + '000000'
    $sheet.Range('F14:H15').Merge()
    $sheet.Range('F14').Formula2 = '=IF(D6="Normal","Support: all real x. Parameters: mean, standard deviation.",IF(D6="Exponential","Support: x >= 0. Parameter: rate, not scale.","Support: integer k >= 0. Fractional probe k is truncated by the library."))'
    $sheet.Range('F14').WrapText = $true
    $sheet.Range('C14:D14').Merge()
    $sheet.Range('C14').Value2 = 'Read the results'
    # Excel turns alerts back on once a worksheet function of the VBA project has run.
    $excel.DisplayAlerts = $false
    $notes = @('Yellow cells are editable; only the selected family uses its parameters. Results are protected.','Select a result or grid cell to inspect its formula. Continuous plots use 201 points.','Normal: mean +/- 4 SD. Exponential: 0 to 8/rate. Poisson: integer bars and CDF points.','Reset Explorer inputs on the Ribbon restores Normal(0,1), after confirmation.','This is an educational demo, not accuracy certification. Other families and quantiles are not offered here.','Needs macros. If Macros says DISABLED: unblock the file or enable content, then reopen.')
    for ($i=0; $i -lt $notes.Count; $i++) {
        $row = 15 + $i*2
        $area = $sheet.Range("C${row}:D$($row+1)")
        $donor.Range('D6').Copy($area)
        # The copy fills every cell of the area. Empty them first: merging
        # cells that hold values asks for confirmation when alerts are on.
        $area.ClearContents()
        $area.Merge()
        $area.Value2 = $notes[$i]
        $area.WrapText = $true
    }
    $sheet.Range('F17').Value2 = 'x / integer k'
    $sheet.Range('G17').Value2 = 'PDF / PMF'
    $sheet.Range('H17').Value2 = 'CDF'
    $donor.Range('D6').Copy($sheet.Range('F18:H218'))
    $sheet.Range('F18:H218').NumberFormatLocal = '0' + $decimal + '000000'
    for ($row=18; $row -le 218; $row++) {
        $index = $row - 18
        $sheet.Range("F$row").Formula2 = '=IF($G$5="OK",IF($D$6="Poisson",IF(' + $index + '<=$G$11,' + $index + ',NA()),$G$10+($G$11-$G$10)*' + $index + '/200),NA())'
        $sheet.Range("G$row").Formula2 = GridFormula 'density' "F$row"
        $sheet.Range("H$row").Formula2 = GridFormula 'cdf' "F$row"
    }
    foreach ($spec in @(@('ExplorerDensity',5,'G'),@('ExplorerCDF',18,'H'))) {
        $anchor = $sheet.Range('J' + $spec[1])
        $chart = $sheet.ChartObjects().Add($anchor.Left,$anchor.Top,$sheet.Range('J:P').Width,205)
        $chart.Name = $spec[0]
        $chart.Chart.ChartType = 75
        $series = $chart.Chart.SeriesCollection().NewSeries()
        $series.Name = $spec[0]
        $series.XValues = $sheet.Range('F18:F218')
        $series.Values = $sheet.Range($spec[2]+'18:'+$spec[2]+'218')
        $chart.Chart.HasTitle = $true
        $chart.Chart.ChartTitle.Text = $spec[0]
        $chart.Chart.HasLegend = $false
        $chart.Chart.ChartArea.Font.Name = $donor.Range('D6').Font.Name
        $chart.Chart.ChartArea.Font.Size = 10
        $series.Format.Line.ForeColor.RGB = 0xB6752E
        $series.Format.Fill.ForeColor.RGB = 0xB6752E
    }
    $donor.Delete()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($donor)
    $prefix = "'" + $book.Name + "'!"
    $excel.Run($prefix + 'ResetExplorerInputs', $false)
    $sheet.Protect()
    $book.Protect([Type]::Missing,$true)
    $sheet.PageSetup.PrintArea = '$B$1:$P$30'
    $sheet.PageSetup.Zoom = $false
    $sheet.PageSetup.FitToPagesWide = 1
    $sheet.PageSetup.FitToPagesTall = $false
    $sheet.PageSetup.Orientation = 2
    $sheet.Activate()
    $excel.ActiveWindow.Zoom = 85
    $sheet.Range('C5').Select()
    $excel.EnableEvents = $true
    $excel.Run($prefix + 'RefreshRibbon')
    $excel.CommandBars.ExecuteMso('FileSave')
    $sheet.ExportAsFixedFormat(0, (Join-Path $directory 'Explorer-preview.pdf'))
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    Write-Output ('Explorer built: ' + (Get-FileHash $path -Algorithm SHA256).Hash)
} catch {
    Write-Output $_.InvocationInfo.PositionMessage
    Write-Output $_.ScriptStackTrace
    throw
} finally {
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
