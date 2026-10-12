# QuotaRecon.Excel.ps1
#
# ImportExcel wrapper. Handles workbook layout, freeze panes, autosize,
# conditional formatting on the Summary sheet.

function Assert-QRExcelModule {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        throw "The ImportExcel module is required. Install with: Install-Module ImportExcel -Scope CurrentUser"
    }
    Import-Module ImportExcel -ErrorAction Stop | Out-Null
}

function Write-QRWorkbook {
    <#
    .SYNOPSIS
        Writes all sheets in one pass. Any input array that's empty gets an
        empty sheet with just the headers so consumers can still filter.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $Path,
        [Parameter(Mandatory)] $Summary,
        [Parameter(Mandatory)] $Quota,
        [Parameter(Mandatory)] $Restrictions,
        [Parameter(Mandatory)] $ZoneMap,
        [Parameter(Mandatory)] $Inputs,
        [Parameter(Mandatory)] $Errors,
        $QuotaGroupLimits      = @(),
        $QuotaGroupAllocations = @(),
        $VmInventory           = @(),
        $ContainerAppsQuota    = @(),
        $ContainerAppsEnvQuota = @(),

        # When $false (quota groups weren't collected), the QuotaGroupLimits
        # and QuotaGroupAllocations sheets are omitted from the workbook
        # entirely rather than written out empty.
        [switch] $IncludeQuotaGroupSheets,

        # When non-empty, the workbook is trimmed down to ONLY these sheet
        # names plus the two always-present sheets (Inputs, Errors) -
        # Summary/Quota/Restrictions/ZoneMap/QuotaGroup*/VMInventory/
        # ContainerApps* are all omitted unless explicitly listed here.
        # This is the mechanism behind every -Only<Service> switch: pass
        # just that service's sheet name(s) to get a minimal, fast,
        # purpose-built workbook. Empty (default) means "include
        # everything", subject to the other per-sheet switches above.
        [string[]] $OnlySheets = @()
    )

    Assert-QRExcelModule

    $onlyModeActive = $OnlySheets -and $OnlySheets.Count -gt 0
    function Test-QRIncludeSheet {
        param([Parameter(Mandatory)][string] $Name)
        if (-not $onlyModeActive) { return $true }
        return $OnlySheets -contains $Name
    }

    # Wipe any previous run so ImportExcel doesn't try to merge sheets.
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force
    }
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $common = @{
        Path         = $Path
        AutoSize     = $true
        FreezeTopRow = $true
        AutoFilter   = $true
        BoldTopRow   = $true
    }

    # --- Summary (with conditional formatting) ---
    if (Test-QRIncludeSheet -Name 'Summary') {
    $summaryRows = if ($Summary -and $Summary.Count) { $Summary } else { @([pscustomobject]@{}) }
    $pkg = $summaryRows | Export-Excel @common -WorksheetName 'Summary' -PassThru
    $ws  = $pkg.Workbook.Worksheets['Summary']
    if ($Summary -and $Summary.Count) {
        $rowCount = $Summary.Count + 1
        # UsagePercent column: yellow >=80, red >=95
        $cols = @{}
        $headerRow = 1
        for ($c = 1; $c -le $ws.Dimension.Columns; $c++) {
            $cols[$ws.Cells[$headerRow, $c].Text] = $c
        }
        function _Col($name) { return $cols[$name] }

        $pctCol = _Col 'MaxUsagePercent'
        if ($pctCol) {
            $range = "$([OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($pctCol))2:" +
                     "$([OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($pctCol))$rowCount"
            Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType GreaterThanOrEqual -ConditionValue 95 -BackgroundColor LightCoral
            Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType GreaterThanOrEqual -ConditionValue 80 -BackgroundColor LightYellow
        }

        # Region-wide and per-zone restriction columns: green for 'None', red for 'Yes'.
        foreach ($zoneCol in @('Regional-Restrictions','Logical-AZ1-Restrictions','Logical-AZ2-Restrictions','Logical-AZ3-Restrictions')) {
            $c = _Col $zoneCol
            if ($c) {
                $range = "$([OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c))2:" +
                         "$([OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c))$rowCount"
                Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType Equal -ConditionValue '"None"' -BackgroundColor LightGreen
                Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType Equal -ConditionValue '"Yes"'  -BackgroundColor LightCoral
            }
        }

        # InQuotaGroup column: light-blue fill when 'Yes'.
        $qgCol = _Col 'InQuotaGroup'
        if ($qgCol) {
            $range = "$([OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($qgCol))2:" +
                     "$([OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($qgCol))$rowCount"
            Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType Equal -ConditionValue '"Yes"' -BackgroundColor LightBlue
        }
    }
    Close-ExcelPackage $pkg
    }

    # --- Remaining sheets ---
    # QuotaGroupLimits/QuotaGroupAllocations are only included when quota
    # groups were actually collected (-IncludeQuotaGroups / -QuotaGroupMgIds)
    # - otherwise they're omitted entirely rather than written out empty.
    # Every entry is also subject to -OnlySheets filtering via
    # Test-QRIncludeSheet (Inputs/Errors are added unconditionally below -
    # they're always present regardless of -Only<Service> mode).
    $rest = New-Object System.Collections.Generic.List[hashtable]
    if (Test-QRIncludeSheet -Name 'Quota')        { $rest.Add(@{ Name = 'Quota';        Data = $Quota }) }
    if (Test-QRIncludeSheet -Name 'Restrictions') { $rest.Add(@{ Name = 'Restrictions'; Data = $Restrictions }) }
    if (Test-QRIncludeSheet -Name 'ZoneMap')      { $rest.Add(@{ Name = 'ZoneMap';      Data = $ZoneMap }) }
    if ($IncludeQuotaGroupSheets) {
        if (Test-QRIncludeSheet -Name 'QuotaGroupLimits')      { $rest.Add(@{ Name = 'QuotaGroupLimits';      Data = $QuotaGroupLimits }) }
        if (Test-QRIncludeSheet -Name 'QuotaGroupAllocations') { $rest.Add(@{ Name = 'QuotaGroupAllocations'; Data = $QuotaGroupAllocations }) }
    }
    if (Test-QRIncludeSheet -Name 'VMInventory')           { $rest.Add(@{ Name = 'VMInventory';           Data = $VmInventory }) }
    if (Test-QRIncludeSheet -Name 'ContainerAppsQuota')    { $rest.Add(@{ Name = 'ContainerAppsQuota';    Data = $ContainerAppsQuota }) }
    if (Test-QRIncludeSheet -Name 'ContainerAppsEnvQuota') { $rest.Add(@{ Name = 'ContainerAppsEnvQuota'; Data = $ContainerAppsEnvQuota }) }
    # Inputs/Errors are always present, regardless of -OnlySheets.
    $rest.Add(@{ Name = 'Inputs'; Data = $Inputs })
    $rest.Add(@{ Name = 'Errors'; Data = $Errors })

    foreach ($sheet in $rest) {
        $rows = if ($sheet.Data -and $sheet.Data.Count) { $sheet.Data } else { @([pscustomobject]@{}) }
        $rows | Export-Excel `
            -Path         $Path `
            -WorksheetName $sheet.Name `
            -AutoSize `
            -FreezeTopRow `
            -AutoFilter `
            -BoldTopRow
    }
}
