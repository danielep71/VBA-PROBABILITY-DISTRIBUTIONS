#==============================================================================
# ExcelGridExport.psm1
#------------------------------------------------------------------------------
# PURPOSE
#   Pure helpers for ci/Export-ExcelGrids.ps1: grid specifications, the
#   byte-level check of an exported grid, the regression-record gate, the Excel
#   environment file and the dialog log of the CI export bridge.
#
# WHY A SEPARATE MODULE
#   Nothing here touches Excel, COM, Win32 or the file system. Every function
#   takes values and returns values, so each rule can be exercised without
#   Excel. benchmark/apply_excel_export.py repeats the grid check in Python on
#   the maintainer's machine, so a defect in either copy is caught by the other.
#
# TWO KINDS OF GRID
#   observed   The exporter rewrites the observed_vba column of the rows it
#              owns in a committed grid (the rules below).
#   generated  A study probe writes its whole CSV afresh (#225): the
#              positive_ratio_subnormal probes. The export must keep the
#              committed header, line count and, row by row, the leading
#              KeyColumns fields that identify each point; the remaining
#              fields are the measurement and may change. Every row must have
#              the header's field count.
#
# GRID RULES (mirrors the exporter modules, which are not changed)
#   The exporters read the whole grid, normalize CRLF and CR to LF, split on
#   LF, rewrite one column of the rows they own, and Print every line back
#   (CRLF-terminated), except a final empty element. A row is owned when it is
#   not blank after trimming spaces and has at least MinimumColumns fields;
#   the main exporter also requires evidence_set = "main grid", and its
#   missing-only mode also requires a blank observed_vba cell. VBA compares
#   strings case-sensitively (Option Compare Binary) and Trim$ removes spaces
#   only, so the checks below do the same.
#==============================================================================

Set-StrictMode -Version Latest

# ISO-8859-1 maps every byte to one char and back, so text comparisons below
# are byte comparisons.
$script:Latin1 = [System.Text.Encoding]::GetEncoding(28591)

# The grids a dispatch selects by name through -Studies (the workflow input
# export_studies). The five grids with their own boolean inputs are not
# listed, so a grid cannot be selected twice.
$script:StudyGridNames = @(
    "prs_gamma_probe", "prs_chisquare_probe", "prs_gamma_holdout",
    "prs_chisquare_holdout", "prs_gamma_q_complement", "gss_small_shape",
    "glq_lower_q", "gbr_binary64_map", "prs_gamma_holdout2", "prs_chisquare_holdout2"
)

function Get-StudyGridNames {
    return , ([string[]]$script:StudyGridNames)
}

function New-GeneratedGridSpec {
    # A study CSV whose identity columns come first and whose measurement
    # follows: written whole, or filled after the key columns, by a
    # non-interactive Export_ entry point beside the export workbook.
    param(
        [string]$Folder = "positive_ratio_subnormal",
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$FileName,
        [Parameter(Mandatory = $true)][string]$Exporter,
        [Parameter(Mandatory = $true)][string]$Macro,
        [Parameter(Mandatory = $true)][int]$Columns,
        [Parameter(Mandatory = $true)][int]$KeyColumns,
        [Parameter(Mandatory = $true)][string]$SuccessPrefix
    )

    $folder = "benchmark/$Folder/"
    return [pscustomobject][ordered]@{
        Name           = $Name
        Path           = $folder + $FileName
        FileName       = $FileName
        Exporter       = $folder + $Exporter
        Macro          = $Macro
        Mode           = "full"
        ObservedIndex  = -1
        MinimumColumns = $Columns
        ScopeIndex     = -1
        ScopeValue     = $null
        SuccessPrefix  = $SuccessPrefix
        ReportsCount   = $false
        Kind           = "generated"
        KeyColumns     = $KeyColumns
    }
}

function ConvertTo-StudySelection {
    # The study grids named in a comma- or space-separated list, in the order
    # given. Empty selects none; an unknown or repeated name is refused.
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value)

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($item in ($Value -split '[,\s]+')) {
        $name = $item.Trim()
        if ($name.Length -eq 0) { continue }
        if ($script:StudyGridNames -cnotcontains $name) {
            throw "unknown study grid '$name'; known: $($script:StudyGridNames -join ', ')"
        }
        if ($names.Contains($name)) { throw "study grid '$name' is named twice" }
        $names.Add($name)
    }
    return , ($names.ToArray())
}

function Get-ExportGridSpec {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("main", "holdout", "holdout_older", "survival_boundary", "tchi_large_df",
                     "prs_gamma_probe", "prs_chisquare_probe", "prs_gamma_holdout",
                     "prs_chisquare_holdout", "prs_gamma_q_complement", "gss_small_shape",
                     "glq_lower_q", "gbr_binary64_map", "prs_gamma_holdout2",
                     "prs_chisquare_holdout2")]
        [string]$Name,

        [ValidateSet("full", "missing-only")]
        [string]$MainMode = "full"
    )

    switch ($Name) {
        "main" {
            $macro = "Export_Accuracy_Observations"
            if ($MainMode -eq "missing-only") { $macro = "Export_Accuracy_MissingOnly" }
            return [pscustomobject][ordered]@{
                Name           = "main"
                Path           = "benchmark/probability_accuracy_grid.csv"
                FileName       = "probability_accuracy_grid.csv"
                Exporter       = "benchmark/M_STATS_PROBDIST_ACCURACYEXPORT.bas"
                Macro          = $macro
                Mode           = $MainMode
                ObservedIndex  = 9
                MinimumColumns = 12
                ScopeIndex     = 11
                ScopeValue     = "main grid"
                SuccessPrefix  = "Accuracy observations written to:"
                ReportsCount   = $false
                Kind           = "observed"
                KeyColumns     = 0
            }
        }
        "holdout" {
            return [pscustomobject][ordered]@{
                Name           = "holdout"
                Path           = "benchmark/holdout/holdout_grid.csv"
                FileName       = "holdout_grid.csv"
                Exporter       = "benchmark/holdout/M_STATS_PROBDIST_HOLDOUT.bas"
                Macro          = "Export_Holdout"
                Mode           = "full"
                ObservedIndex  = 9
                MinimumColumns = 12
                ScopeIndex     = -1
                ScopeValue     = $null
                SuccessPrefix  = "Holdout complete: "
                ReportsCount   = $true
                Kind           = "observed"
                KeyColumns     = 0
            }
        }
        "holdout_older" {
            return [pscustomobject][ordered]@{
                Name           = "holdout_older"
                Path           = "benchmark/holdout_older/holdout_older_grid.csv"
                FileName       = "holdout_older_grid.csv"
                Exporter       = "benchmark/holdout_older/M_STATS_PROBDIST_HOLDOUTOLDER.bas"
                Macro          = "Export_HoldoutOlder"
                Mode           = "full"
                ObservedIndex  = 8
                MinimumColumns = 11
                ScopeIndex     = -1
                ScopeValue     = $null
                SuccessPrefix  = "Older-families holdout export complete: "
                ReportsCount   = $true
                Kind           = "observed"
                KeyColumns     = 0
            }
        }
        "survival_boundary" {
            return [pscustomobject][ordered]@{
                Name           = "survival_boundary"
                Path           = "benchmark/survival_boundary/survival_boundary_grid.csv"
                FileName       = "survival_boundary_grid.csv"
                Exporter       = "benchmark/survival_boundary/M_STATS_PROBDIST_SURVBOUND.bas"
                Macro          = "Export_SurvivalBoundary"
                Mode           = "full"
                ObservedIndex  = 8
                MinimumColumns = 11
                ScopeIndex     = -1
                ScopeValue     = $null
                SuccessPrefix  = "Survival-boundary export complete: "
                ReportsCount   = $true
                Kind           = "observed"
                KeyColumns     = 0
            }
        }
        "prs_gamma_probe" {
            return New-GeneratedGridSpec -Name "prs_gamma_probe" -FileName "gamma_probe.csv" `
                -Exporter "M_STATS_PROBDIST_GAMMAPROBE.bas" -Macro "Export_GammaPositiveRatio" `
                -Columns 15 -KeyColumns 9 -SuccessPrefix "Gamma positive-ratio probe complete: "
        }
        "prs_chisquare_probe" {
            return New-GeneratedGridSpec -Name "prs_chisquare_probe" -FileName "chisquare_probe.csv" `
                -Exporter "M_STATS_PROBDIST_CHISQPROBE.bas" -Macro "Export_ChiSquarePositiveRatio" `
                -Columns 14 -KeyColumns 8 -SuccessPrefix "Chi-square positive-ratio probe complete: "
        }
        "prs_gamma_holdout" {
            return New-GeneratedGridSpec -Name "prs_gamma_holdout" -FileName "gamma_holdout.csv" `
                -Exporter "M_STATS_PROBDIST_GAMMAHOLD.bas" -Macro "Export_GammaHoldout" `
                -Columns 15 -KeyColumns 9 -SuccessPrefix "Gamma positive-ratio HOLDOUT complete: "
        }
        "prs_chisquare_holdout" {
            return New-GeneratedGridSpec -Name "prs_chisquare_holdout" -FileName "chisquare_holdout.csv" `
                -Exporter "M_STATS_PROBDIST_CHISQHOLD.bas" -Macro "Export_ChiSquareHoldout" `
                -Columns 14 -KeyColumns 8 -SuccessPrefix "Chi-square positive-ratio HOLDOUT complete: "
        }
        # The second holdouts (#13, Amendment 1) key on the six text fields
        # only: their echoed operands are measurements, checked by
        # analyze_amendment1.py against the exact construction.
        "prs_gamma_holdout2" {
            return New-GeneratedGridSpec -Name "prs_gamma_holdout2" -FileName "gamma_holdout2.csv" `
                -Exporter "M_STATS_PROBDIST_GAMMAHOLD2.bas" -Macro "Export_GammaHoldout2" `
                -Columns 15 -KeyColumns 6 -SuccessPrefix "Gamma positive-ratio HOLDOUT 2 complete: "
        }
        "prs_chisquare_holdout2" {
            return New-GeneratedGridSpec -Name "prs_chisquare_holdout2" -FileName "chisquare_holdout2.csv" `
                -Exporter "M_STATS_PROBDIST_CHISQHOLD2.bas" -Macro "Export_ChiSquareHoldout2" `
                -Columns 14 -KeyColumns 6 -SuccessPrefix "Chi-square positive-ratio HOLDOUT 2 complete: "
        }
        "prs_gamma_q_complement" {
            return New-GeneratedGridSpec -Name "prs_gamma_q_complement" -FileName "gamma_q_complement.csv" `
                -Exporter "M_STATS_PROBDIST_GQPROBE.bas" -Macro "Export_GammaQComplement" `
                -Columns 12 -KeyColumns 3 -SuccessPrefix "Gamma Q complement probe complete: "
        }
        "gss_small_shape" {
            return New-GeneratedGridSpec -Folder "gamma_series_small_shape" -Name "gss_small_shape" `
                -FileName "gamma_series_small_shape_grid.csv" `
                -Exporter "M_STATS_PROBDIST_GSSS.bas" -Macro "Export_GammaSeriesSmallShape" `
                -Columns 13 -KeyColumns 10 -SuccessPrefix "Gamma small-shape study complete: "
        }
        "glq_lower_q" {
            return New-GeneratedGridSpec -Folder "gamma_lower_q" -Name "glq_lower_q" `
                -FileName "gamma_lower_q_grid.csv" `
                -Exporter "M_STATS_PROBDIST_GLQ.bas" -Macro "Export_GammaLowerQ" `
                -Columns 13 -KeyColumns 10 -SuccessPrefix "Gamma lower-Q holdout complete: "
        }
        "gbr_binary64_map" {
            return New-GeneratedGridSpec -Folder "gamma_route_study" -Name "gbr_binary64_map" `
                -FileName "binary64_map_grid.csv" `
                -Exporter "M_STATS_PROBDIST_GBR.bas" -Macro "Export_GammaBinary64Map" `
                -Columns 13 -KeyColumns 5 -SuccessPrefix "Gamma binary64 route map complete: "
        }
        "tchi_large_df" {
            return [pscustomobject][ordered]@{
                Name           = "tchi_large_df"
                Path           = "benchmark/tchi_large_df/tchi_large_df_grid.csv"
                FileName       = "tchi_large_df_grid.csv"
                Exporter       = "benchmark/tchi_large_df/M_STATS_PROBDIST_TCHI.bas"
                Macro          = "Export_TChiLargeDf"
                Mode           = "full"
                ObservedIndex  = 8
                MinimumColumns = 9
                ScopeIndex     = -1
                ScopeValue     = $null
                SuccessPrefix  = "Large-df t / Chi-square study complete: "
                ReportsCount   = $true
                Kind           = "observed"
                KeyColumns     = 0
            }
        }
    }
}

function ConvertTo-BooleanSwitch {
    # Workflow inputs arrive as the strings "true" and "false".
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value, [string]$Name = "value")

    switch ($Value.Trim().ToLowerInvariant()) {
        "true" { return $true }
        "false" { return $false }
        "" { return $false }
        default { throw "$Name must be true or false, not '$Value'" }
    }
}

function ConvertTo-GridLines {
    # The line list an exporter works on: CRLF and CR become LF, the text is
    # split on LF, and a final empty element (a trailing newline) is dropped,
    # because the exporters never Print it.
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$Bytes)

    $text = $script:Latin1.GetString($Bytes).Replace("`r`n", "`n").Replace("`r", "`n")
    $list = New-Object System.Collections.Generic.List[string]
    $list.AddRange([string[]]$text.Split([char[]]@([char]10)))
    if ($list.Count -gt 0 -and $list[$list.Count - 1].Length -eq 0) {
        $list.RemoveAt($list.Count - 1)
    }
    return , ($list.ToArray())
}

function ConvertTo-PrintedGridText {
    # The exact text VBA's Print # writes for these lines: each one followed by
    # CRLF.
    param([string[]]$Lines = @())

    if ($Lines.Length -eq 0) { return "" }
    return [string]::Join("`r`n", $Lines) + "`r`n"
}

function Test-GridRowInScope {
    param(
        [Parameter(Mandatory = $true)]$Spec,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Line
    )

    if ($Line.Trim(' ').Length -eq 0) { return $false }
    $cols = $Line.Split([char[]]@([char]','))
    if ($cols.Length -lt $Spec.MinimumColumns) { return $false }
    if ($Spec.ScopeIndex -ge 0 -and ($cols[$Spec.ScopeIndex].Trim(' ') -cne $Spec.ScopeValue)) {
        return $false
    }
    if ($Spec.Mode -eq "missing-only" -and $cols[$Spec.ObservedIndex].Trim(' ').Length -gt 0) {
        return $false
    }
    return $true
}

function Test-GeneratedGrid {
    # A probe CSV written afresh, against the committed file it replaces: the
    # bytes VBA's Print # writes, the same header and line count, every row
    # with the header's field count, and the first KeyColumns fields of each
    # row unchanged. Only the measurement fields after them may differ.
    param(
        [Parameter(Mandatory = $true)]$Spec,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$BaselineBytes,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$ExportedBytes,
        [int]$MaxListed = 20
    )

    $problems = New-Object System.Collections.Generic.List[string]
    $before = ConvertTo-GridLines -Bytes $BaselineBytes
    $after = ConvertTo-GridLines -Bytes $ExportedBytes
    $filled = 0
    $errors = 0
    $changed = 0
    $listed = 0

    if ($script:Latin1.GetString($ExportedBytes) -cne (ConvertTo-PrintedGridText -Lines $after)) {
        $problems.Add("the export is not CRLF-terminated lines as VBA Print # writes them")
    }
    if ($before.Length -eq 0) {
        $problems.Add("the checked-out grid is empty")
    }
    if ($before.Length -ne $after.Length) {
        $problems.Add("line count changed from $($before.Length) to $($after.Length)")
    }
    elseif ($before.Length -gt 0) {
        if ($before[0] -cne $after[0]) {
            $problems.Add("the header line changed")
        }
        if ($before[0].Split([char[]]@([char]',')).Length -ne $Spec.MinimumColumns) {
            $problems.Add("the committed header does not have $($Spec.MinimumColumns) fields")
        }
        for ($i = 1; $i -lt $before.Length; $i++) {
            $oldCols = $before[$i].Split([char[]]@([char]','))
            $newCols = $after[$i].Split([char[]]@([char]','))
            $lineNumber = $i + 1
            if ($newCols.Length -ne $Spec.MinimumColumns) {
                if ($listed -lt $MaxListed) { $problems.Add("line ${lineNumber}: $($newCols.Length) fields, not $($Spec.MinimumColumns)") }
                $listed += 1
                continue
            }
            $keyChanged = $false
            for ($j = 0; $j -lt $Spec.KeyColumns; $j++) {
                if ($j -ge $oldCols.Length -or $oldCols[$j] -cne $newCols[$j]) { $keyChanged = $true }
            }
            if ($keyChanged) {
                if ($listed -lt $MaxListed) { $problems.Add("line ${lineNumber}: the point's key fields changed") }
                $listed += 1
                continue
            }
            $filled += 1
            $measured = $newCols[$Spec.KeyColumns..($newCols.Length - 1)]
            if (@($measured -ceq "ERROR").Count -gt 0 -or @($measured -ceq "CVERR").Count -gt 0) { $errors += 1 }
            if ($before[$i] -cne $after[$i]) { $changed += 1 }
        }
        if ($listed -gt $MaxListed) {
            $problems.Add("... and $($listed - $MaxListed) more row problem(s)")
        }
        if ($before.Length -lt 2) {
            $problems.Add("no row is in the exporter's scope")
        }
    }

    return [pscustomobject][ordered]@{
        rows                 = [Math]::Max($after.Length - 1, 0)
        in_scope_rows        = [Math]::Max($before.Length - 1, 0)
        filled_rows          = $filled
        error_rows           = $errors
        out_of_scope_rows    = 0
        changed_observations = $changed
        problems             = $problems.ToArray()
    }
}

function Test-ExportedGrid {
    # Compare an exported grid with the checked-out grid it was made from.
    # Returns the counts and every problem found; an empty Problems array means
    # the export changed exactly the cells its exporter owns, filled every one
    # of them, and wrote the bytes VBA's Print # writes.
    param(
        [Parameter(Mandatory = $true)]$Spec,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$BaselineBytes,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$ExportedBytes,
        [int]$MaxListed = 20
    )

    if ($Spec.Kind -eq "generated") {
        return Test-GeneratedGrid -Spec $Spec -BaselineBytes $BaselineBytes -ExportedBytes $ExportedBytes -MaxListed $MaxListed
    }
    $problems = New-Object System.Collections.Generic.List[string]
    $before = ConvertTo-GridLines -Bytes $BaselineBytes
    $after = ConvertTo-GridLines -Bytes $ExportedBytes
    $inScope = 0
    $filled = 0
    $errors = 0
    $outOfScope = 0
    $changed = 0
    $listed = 0

    if ($script:Latin1.GetString($ExportedBytes) -cne (ConvertTo-PrintedGridText -Lines $after)) {
        $problems.Add("the export is not CRLF-terminated lines as VBA Print # writes them")
    }
    if ($before.Length -eq 0) {
        $problems.Add("the checked-out grid is empty")
    }
    if ($before.Length -ne $after.Length) {
        $problems.Add("line count changed from $($before.Length) to $($after.Length)")
    }
    else {
        if ($before.Length -gt 0 -and $before[0] -cne $after[0]) {
            $problems.Add("the header line changed")
        }
        for ($i = 1; $i -lt $before.Length; $i++) {
            $old = $before[$i]
            $new = $after[$i]
            $lineNumber = $i + 1
            if (-not (Test-GridRowInScope -Spec $Spec -Line $old)) {
                $outOfScope += 1
                if ($old -cne $new) {
                    if ($listed -lt $MaxListed) { $problems.Add("line ${lineNumber}: out-of-scope row changed") }
                    $listed += 1
                }
                continue
            }

            $inScope += 1
            $oldCols = $old.Split([char[]]@([char]','))
            $newCols = $new.Split([char[]]@([char]','))
            if ($oldCols.Length -ne $newCols.Length) {
                if ($listed -lt $MaxListed) { $problems.Add("line ${lineNumber}: field count changed") }
                $listed += 1
                continue
            }
            for ($j = 0; $j -lt $oldCols.Length; $j++) {
                if ($j -ne $Spec.ObservedIndex -and $oldCols[$j] -cne $newCols[$j]) {
                    if ($listed -lt $MaxListed) { $problems.Add("line ${lineNumber}: column $j changed outside observed_vba") }
                    $listed += 1
                }
            }
            $observed = $newCols[$Spec.ObservedIndex].Trim(' ')
            if ($observed.Length -eq 0) {
                if ($listed -lt $MaxListed) { $problems.Add("line ${lineNumber}: observed_vba left blank") }
                $listed += 1
            }
            else {
                $filled += 1
                if ($observed -ceq "ERROR") { $errors += 1 }
            }
            if ($oldCols[$Spec.ObservedIndex] -cne $newCols[$Spec.ObservedIndex]) {
                $changed += 1
            }
        }
        if ($listed -gt $MaxListed) {
            $problems.Add("... and $($listed - $MaxListed) more row problem(s)")
        }
        if ($inScope -eq 0) {
            if ($Spec.Mode -eq "missing-only") {
                $problems.Add("missing-only found no blank observed_vba cell in a main-grid row; nothing was exported (use full)")
            }
            else {
                $problems.Add("no row is in the exporter's scope")
            }
        }
    }

    return [pscustomobject][ordered]@{
        rows                 = [Math]::Max($after.Length - 1, 0)
        in_scope_rows        = $inScope
        filled_rows          = $filled
        error_rows           = $errors
        out_of_scope_rows    = $outOfScope
        changed_observations = $changed
        problems             = $problems.ToArray()
    }
}

function Get-Sha256Text {
    # "sha256:<hex>" of a byte array, the digest format of the evidence records.
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$Bytes)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash($Bytes)
    }
    finally {
        $sha.Dispose()
    }
    return "sha256:" + (($hash | ForEach-Object { $_.ToString("x2") }) -join "")
}

function Test-GreenRegressionRecord {
    # Reasons the regression record does not authorize an export of this
    # candidate in this run; empty when it does.
    param(
        [Parameter(Mandatory = $true)]$Record,
        [Parameter(Mandatory = $true)][string]$CandidateSha,
        [Parameter(Mandatory = $true)][int]$ExpectedAssertions,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$LogSha256,
        [Int64]$RunId = 0
    )

    $reasons = New-Object System.Collections.Generic.List[string]
    if ([string]$Record.candidate_sha -cne $CandidateSha) {
        $reasons.Add("record candidate $($Record.candidate_sha) is not the checked-out $CandidateSha")
    }
    foreach ($stage in @("import", "compile", "regression", "cleanup")) {
        $status = [string]$Record.stages.$stage.status
        if ($status -cne "PASS") {
            $reasons.Add("stage $stage is $status, not PASS")
        }
    }
    $assertions = [int]$Record.harness.assertions
    $passed = [int]$Record.harness.passed
    $failed = [int]$Record.harness.failed
    if ($assertions -ne $ExpectedAssertions -or $passed -ne $assertions -or $failed -ne 0) {
        $reasons.Add("harness reports $passed/$assertions passed and $failed failed; the policy requires $ExpectedAssertions/$ExpectedAssertions")
    }
    if ([string]$Record.log.sha256 -cne $LogSha256) {
        $reasons.Add("the regression log digest $LogSha256 differs from the record's $($Record.log.sha256)")
    }
    if ($RunId -gt 0) {
        $workflow = $Record.runner.workflow
        if ($null -eq $workflow -or [Int64]$workflow.run_id -ne $RunId) {
            $reasons.Add("the record does not name this workflow run ($RunId)")
        }
    }
    return , ($reasons.ToArray())
}

function ConvertTo-OfficeBitnessLabel {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value)

    switch ($Value.Trim().ToLowerInvariant().Replace(" ", "")) {
        { $_ -in @("64", "64bit", "64-bit", "x64") } { return "64-bit" }
        { $_ -in @("32", "32bit", "32-bit", "x86") } { return "32-bit" }
        default { return $Value.Trim() }
    }
}

function Read-ExcelEnvironmentFile {
    # Parse the three fields Export_ExcelEnvironment writes and refuse anything
    # that is not a plain version, build and bitness.
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$Bytes)

    $text = $script:Latin1.GetString($Bytes)
    try {
        $parsed = $text | ConvertFrom-Json
    }
    catch {
        throw "Export_ExcelEnvironment wrote invalid JSON: $($_.Exception.Message)"
    }
    $names = @($parsed.PSObject.Properties | ForEach-Object { $_.Name } | Sort-Object)
    if (($names -join ",") -cne "excel_build,excel_version,office_bitness") {
        throw "Export_ExcelEnvironment wrote unexpected fields: $($names -join ', ')"
    }
    $version = [string]$parsed.excel_version
    $build = [string]$parsed.excel_build
    $bitness = [string]$parsed.office_bitness
    if ($version -notmatch '^[0-9]+(\.[0-9]+)*$') { throw "excel_version '$version' is not a version number" }
    if ($build -notmatch '^[0-9]+(\.[0-9]+)*$') { throw "excel_build '$build' is not a build number" }
    if ($bitness -notmatch '^(32|64)$') { throw "office_bitness '$bitness' is not 32 or 64" }
    return [pscustomobject][ordered]@{
        excel_version  = $version
        excel_build    = $build
        office_bitness = $bitness
    }
}

function ConvertTo-CompletedEnvironmentJson {
    # The committed benchmark/excel_environment.json: the three fields Excel
    # wrote, unchanged, plus the commit SHA, which Export_ExcelEnvironment
    # cannot know. commit_sha_recorded_by says who added it. LF line endings,
    # as Git stores *.json.
    param(
        [Parameter(Mandatory = $true)]$Environment,
        [Parameter(Mandatory = $true)][string]$CommitSha,
        [Parameter(Mandatory = $true)][string]$RecordedBy
    )

    if ($CommitSha -notmatch '^[0-9a-f]{40}$') { throw "commit SHA '$CommitSha' is not a full lowercase SHA" }
    if ($RecordedBy -match '["\\]' -or $RecordedBy -match '[\x00-\x1f]') { throw "recorded-by text must not need JSON escaping" }
    $lines = @(
        "{",
        "  `"excel_version`": `"$($Environment.excel_version)`",",
        "  `"excel_build`": `"$($Environment.excel_build)`",",
        "  `"office_bitness`": `"$($Environment.office_bitness)`",",
        "  `"commit_sha`": `"$CommitSha`",",
        "  `"commit_sha_recorded_by`": `"$RecordedBy`"",
        "}"
    )
    return ($lines -join "`n") + "`n"
}

function Split-BridgeDialogLog {
    # The CI export bridge joins dialogs with Chr(30) and separates each
    # title from its prompt with Chr(31).
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Log)

    $dialogs = New-Object System.Collections.Generic.List[object]
    if ($Log.Length -eq 0) { return , ($dialogs.ToArray()) }
    foreach ($entry in $Log.Split([char[]]@([char]30))) {
        $parts = $entry.Split([char[]]@([char]31), 2)
        if ($parts.Length -eq 2) {
            $dialogs.Add([pscustomobject]@{ Title = $parts[0]; Prompt = $parts[1] })
        }
        else {
            $dialogs.Add([pscustomobject]@{ Title = ""; Prompt = $entry })
        }
    }
    return , ($dialogs.ToArray())
}

function Format-DialogText {
    param([Parameter(Mandatory = $true)]$Dialog)

    $prompt = ([string]$Dialog.Prompt) -replace "\r?\n", " / "
    return "[$($Dialog.Title)] $prompt"
}

function Test-ExportDialogs {
    # The exporters report through exactly one message box. A run passes only
    # when that one dialog is the success message for the expected file, and,
    # where the exporter reports a count, the count of rows it filled.
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Dialogs,
        [Parameter(Mandatory = $true)][string]$SuccessPrefix,
        [Parameter(Mandatory = $true)][string]$ExpectedPath,
        [int]$ExpectedCount = -1
    )

    $reasons = New-Object System.Collections.Generic.List[string]
    if ($Dialogs.Length -ne 1) {
        $reasons.Add("expected exactly one exporter message, got $($Dialogs.Length)")
        foreach ($dialog in $Dialogs) { $reasons.Add("  message: " + (Format-DialogText -Dialog $dialog)) }
        return , ($reasons.ToArray())
    }
    $prompt = [string]$Dialogs[0].Prompt
    if (-not $prompt.StartsWith($SuccessPrefix, [System.StringComparison]::Ordinal)) {
        $reasons.Add("the exporter did not report success: " + (Format-DialogText -Dialog $Dialogs[0]))
        return , ($reasons.ToArray())
    }
    $reportedPath = $null
    foreach ($line in ($prompt -split "\r?\n")) {
        if ($line.EndsWith(".csv", [System.StringComparison]::OrdinalIgnoreCase) -or
            $line.EndsWith(".json", [System.StringComparison]::OrdinalIgnoreCase)) {
            $reportedPath = $line.Trim()
        }
    }
    if ($null -eq $reportedPath -or -not [string]::Equals($reportedPath, $ExpectedPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        $reasons.Add("the exporter reports writing '$reportedPath', not '$ExpectedPath'")
    }
    if ($ExpectedCount -ge 0) {
        if ($prompt -match ':\s*(\d+)\s+observation') {
            if ([int]$Matches[1] -ne $ExpectedCount) {
                $reasons.Add("the exporter reports $($Matches[1]) observation(s); $ExpectedCount rows are in its scope")
            }
        }
        else {
            $reasons.Add("the exporter's message carries no observation count: " + (Format-DialogText -Dialog $Dialogs[0]))
        }
    }
    return , ($reasons.ToArray())
}

Export-ModuleMember -Function @(
    "Get-StudyGridNames",
    "ConvertTo-StudySelection",
    "Get-ExportGridSpec",
    "ConvertTo-BooleanSwitch",
    "ConvertTo-GridLines",
    "ConvertTo-PrintedGridText",
    "Test-GridRowInScope",
    "Test-GeneratedGrid",
    "Test-ExportedGrid",
    "Get-Sha256Text",
    "Test-GreenRegressionRecord",
    "ConvertTo-OfficeBitnessLabel",
    "Read-ExcelEnvironmentFile",
    "ConvertTo-CompletedEnvironmentJson",
    "Split-BridgeDialogLog",
    "Format-DialogText",
    "Test-ExportDialogs"
)
