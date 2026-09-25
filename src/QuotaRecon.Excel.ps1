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
        $VmInventory           = @()
    )

    Assert-QRExcelModule

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

    # --- Remaining sheets ---
    $rest = @(
        @{ Name = 'Quota';                 Data = $Quota },
        @{ Name = 'Restrictions';          Data = $Restrictions },
        @{ Name = 'ZoneMap';               Data = $ZoneMap },
        @{ Name = 'QuotaGroupLimits';      Data = $QuotaGroupLimits },
        @{ Name = 'QuotaGroupAllocations'; Data = $QuotaGroupAllocations },
        @{ Name = 'VMInventory';           Data = $VmInventory },
        @{ Name = 'Inputs';                Data = $Inputs },
        @{ Name = 'Errors';                Data = $Errors }
    )
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
