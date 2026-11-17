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
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($path, 0, $false)
    $book.Unprotect()
    $sheet = $book.Worksheets.Item('Explorer')
    $sheet.Unprotect()
    # Temporary style donor; deleted before saving. No independent visual theme.
    $sheet.Copy([Type]::Missing, $book.Worksheets.Item($book.Worksheets.Count))
    $donor = $book.Worksheets.Item($book.Worksheets.Count)
    $sheet.UsedRange.Clear()
    $sheet.Columns.Item('C').ColumnWidth = 26
    $sheet.Columns.Item('D').ColumnWidth = 18
    $sheet.Columns.Item('E').ColumnWidth = 3
    $sheet.Columns.Item('F:H').Hidden = $false
    $sheet.Columns.Item('F:H').ColumnWidth = 17
    $sheet.Columns.Item('I').Hidden = $false
    $sheet.Columns.Item('I').ColumnWidth = 3
    $sheet.Columns.Item('J:P').Hidden = $false
    $sheet.Columns.Item('J:P').ColumnWidth = 12
    $sheet.Rows.Item('1:218').RowHeight = 22
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
    $sheet.Rows.Item('6:11').RowHeight = 34
    $sheet.Range('D6:D11').Locked = $false
    # K-SACCR input palette; preserve the template's typeface and borders.
    $sheet.Range('D6:D11').Interior.Color = 0xCCF2FF
    $sheet.Range('D6:D11').NumberFormat = '0.0000'
    $sheet.Range('D6').Validation.Add(3, 1, 1, 'Normal,Exponential,Poisson')
    $sheet.Range('D6').Validation.ErrorTitle = 'Choose a listed family'
    $sheet.Range('D6').Validation.ShowError = $true
    $limits = @(@('D7',-1000000,1000000),@('D8',0.001,10000),@('D9',0.001,1000),@('D10',0.01,100),@('D11',-10000000,10000000))
    foreach ($limit in $limits) {
        $inputCell = $sheet.Range($limit[0])
        $inputCell.Validation.Add(2, 1, 1, [string]$limit[1], [string]$limit[2])
        $inputCell.Validation.ShowError = $true
    }
    $donor.Range('D6').Copy($sheet.Range('F6:H15'))
    $sheet.Range('G5:H5').Merge()
    $sheet.Range('G5').Formula = '=IF(IFERROR(AND(ISNUMBER(D11),ABS(D11)<=10000000,OR(AND(D6="Normal",ISNUMBER(D7),ABS(D7)<=1000000,ISNUMBER(D8),D8>=0.001,D8<=10000),AND(D6="Exponential",ISNUMBER(D9),D9>=0.001,D9<=1000),AND(D6="Poisson",ISNUMBER(D10),D10>=0.01,D10<=100))),FALSE),"OK","Check inputs")'
    $sheet.Range('F5').Value2 = 'Input status'
    $sheet.Range('F7').Value2 = 'PDF / PMF at x'
    $sheet.Range('F8').Value2 = 'CDF at x'
    $sheet.Range('F9').Value2 = 'P(X > x)'
    $sheet.Range('F10').Value2 = 'Plot lower bound'
    $sheet.Range('F11').Value2 = 'Plot upper bound'
    $sheet.Range('F12').Value2 = 'Right tail omitted'
    $sheet.Range('G7').Formula = DistributionFormula 'density' '$D$11'
    $sheet.Range('G8').Formula = DistributionFormula 'cdf' '$D$11'
    $sheet.Range('G9').Formula = DistributionFormula 'survival' '$D$11'
    $sheet.Range('G10').Formula = '=IF($G$5="OK",IF(D6="Normal",D7-4*D8,0),NA())'
    $sheet.Range('G11').Formula = '=IF($G$5="OK",IF(D6="Normal",D7+4*D8,IF(D6="Exponential",8/D9,ROUNDUP(D10+8*SQRT(D10)+10,0))),NA())'
    $sheet.Range('G12').Formula = DistributionFormula 'survival' '$G$11'
    $sheet.Range('G7:G12').NumberFormat = '0.000000;[Red]-0.000000;0.000000'
    $sheet.Range('F14:H15').Merge()
    $sheet.Range('F14').Formula = '=IF(D6="Normal","Support: all real x. Parameters: mean, standard deviation.",IF(D6="Exponential","Support: x >= 0. Parameter: rate, not scale.","Support: integer k >= 0. Fractional probe k is truncated by the library."))'
    $sheet.Range('F14').WrapText = $true
    $sheet.Range('C14:D14').Merge()
    $sheet.Range('C14').Value2 = 'Read the results'
    $notes = @('Only the selected family uses its parameters. Other inputs are retained.','Yellow cells: editable inputs. Results and formulas are protected.','Select any result or grid cell to inspect its formula in the formula bar.','Continuous plots: 201 equally spaced points. Normal: mean +/- 4 standard deviations; Exponential: 0 to 8/rate.','Poisson: integer bars for PMF and unconnected CDF points. No interpolation between counts.','Limits here keep charts readable; they are narrower than the public API. No quantiles or other operations are exposed in this view.','Use Reset Explorer inputs on the Ribbon for Normal(0,1). Cancellation changes nothing.','Educational demo, not certified accuracy evidence. Independent reference checks are recorded by CI.','Macros disabled: navigation remains usable; formulas/charts show the last saved state. Do not edit inputs until macros are enabled.')
    for ($i=0; $i -lt $notes.Count; $i++) {
        $row = 15 + $i*3
        $area = $sheet.Range("C${row}:D$($row+2)")
        $donor.Range('D6').Copy($area)
        $area.Merge()
        $area.Value2 = $notes[$i]
        $area.WrapText = $true
    }
    $sheet.Range('F17').Value2 = 'x / integer k'
    $sheet.Range('G17').Value2 = 'PDF / PMF'
    $sheet.Range('H17').Value2 = 'CDF'
    $donor.Range('D6').Copy($sheet.Range('F18:H218'))
    $sheet.Range('F18:H218').NumberFormat = '0.000000'
    for ($row=18; $row -le 218; $row++) {
        $index = $row - 18
        $sheet.Range("F$row").Formula = '=IF($G$5="OK",IF($D$6="Poisson",IF(' + $index + '<=$G$11,' + $index + ',NA()),$G$10+($G$11-$G$10)*' + $index + '/200),NA())'
        $sheet.Range("G$row").Formula = DistributionFormula 'density' "F$row"
        $sheet.Range("H$row").Formula = DistributionFormula 'cdf' "F$row"
    }
    foreach ($spec in @(@('ExplorerDensity',5,'G'),@('ExplorerCDF',20,'H'))) {
        $anchor = $sheet.Range('J' + $spec[1])
        $chart = $sheet.ChartObjects().Add($anchor.Left,$anchor.Top,$sheet.Range('J:P').Width,300)
        $chart.Name = $spec[0]
        $chart.Chart.ChartType = 74
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
    $sheet.PageSetup.PrintArea = '$B$1:$P$44'
    $sheet.PageSetup.Zoom = $false
    $sheet.PageSetup.FitToPagesWide = 1
    $sheet.PageSetup.FitToPagesTall = 1
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
} finally {
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
