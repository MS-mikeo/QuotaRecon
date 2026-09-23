#Requires -Version 5.1
<#
.SYNOPSIS
    QuotaRecon — audits Azure compute quota, SKU restrictions, and
    availability-zone mappings across multiple subscriptions, regions, and
    VM SKUs, and produces a single .xlsx report.

.DESCRIPTION
    Reads three inputs (subscriptions, regions, SKUs) from CSVs or inline
    parameters, normalizes them against live Azure catalogs, then for every
    (Sub × Region × SKU) triple collects:

      * Regional vCPU + family vCPU quota (Microsoft.Compute/usages)
      * SKU availability + restrictions (Microsoft.Compute/skus)
      * Logical → physical availability-zone map
        (Microsoft.Resources/checkZonePeers)

    Output is one Excel workbook with sheets:
    Summary, Quota, Restrictions, ZoneMap, Inputs, Errors.

.PARAMETER SubscriptionsCsv
    Path to a single-column CSV of subscription GUIDs.

.PARAMETER RegionsCsv
    Path to a single-column CSV of Azure region names (ARM or display name).

.PARAMETER SkusCsv
    Path to a single-column CSV of VM SKU names.

.PARAMETER SubscriptionIds
    Inline alternative to -SubscriptionsCsv.

.PARAMETER Regions
    Inline alternative to -RegionsCsv.

.PARAMETER Skus
    Inline alternative to -SkusCsv.

.PARAMETER OutputPath
    Where the .xlsx should be written. Default: ./output/QuotaRecon-<UTCstamp>.xlsx

.EXAMPLE
    ./Invoke-QuotaRecon.ps1 `
        -SubscriptionsCsv ./templates/subscriptions.csv `
        -RegionsCsv       ./templates/regions.csv `
        -SkusCsv          ./templates/skus.csv

.EXAMPLE
    ./Invoke-QuotaRecon.ps1 `
        -SubscriptionIds '11111111-...','22222222-...' `
        -Regions         'eastus2','westus3' `
        -Skus            'Standard_D4s_v5','E8ds v5'

.NOTES
    Requires: Az.Accounts + ImportExcel modules. Run Connect-AzAccount first.
#>
[CmdletBinding()]
param(
    [string]   $SubscriptionsCsv,
    [string]   $RegionsCsv,
    [string]   $SkusCsv,

    [string[]] $SubscriptionIds,
    [string[]] $Regions,
    [string[]] $Skus,

    [string]   $OutputPath,

    # When set, writes every ARM request URL (and 4xx response bodies) to a
    # transcript log next to the output workbook. Useful for reproducing
    # unexpected results. Zero overhead when not set.
    [switch]   $DiagnosticLog
)

$ErrorActionPreference = 'Stop'

# ---------- Load modules ------------------------------------------------
$here = Split-Path -Parent $PSCommandPath
. (Join-Path $here 'src\QuotaRecon.Auth.ps1')
. (Join-Path $here 'src\QuotaRecon.Input.ps1')
. (Join-Path $here 'src\QuotaRecon.Skus.ps1')
. (Join-Path $here 'src\QuotaRecon.Quota.ps1')
. (Join-Path $here 'src\QuotaRecon.Zones.ps1')
. (Join-Path $here 'src\QuotaRecon.Excel.ps1')

if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
    throw "Az.Accounts is required. Install with: Install-Module Az.Accounts -Scope CurrentUser"
}
Import-Module Az.Accounts -ErrorAction Stop | Out-Null
Assert-QRExcelModule

if (-not (Get-AzContext -ErrorAction SilentlyContinue)) {
    throw "No Az context. Run Connect-AzAccount before invoking QuotaRecon."
}

# ---------- Output path -------------------------------------------------
if (-not $OutputPath) {
    $stamp = (Get-Date).ToUniversalTime().ToString('yyyyMMdd-HHmmssZ')
    $OutputPath = Join-Path $here "output\QuotaRecon-$stamp.xlsx"
}
$outDir = Split-Path -Parent $OutputPath
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

# ---------- Diagnostic log ---------------------------------------------
# Only touched when -DiagnosticLog is passed. Starts a transcript alongside
# the XLSX and enables Write-Verbose everywhere so each ARM URL and any 4xx
# body ends up in the log. Zero cost when the switch isn't used.
$transcriptStarted = $false
if ($DiagnosticLog) {
    $logPath = [System.IO.Path]::ChangeExtension($OutputPath, '.log')
    try {
        Start-Transcript -Path $logPath -Force | Out-Null
        $transcriptStarted = $true
        Write-Host ("[QuotaRecon] Diagnostic log: {0}" -f $logPath) -ForegroundColor DarkCyan
    } catch {
        Write-Warning "Could not start transcript: $($_.Exception.Message)"
    }
    $VerbosePreference = 'Continue'
}

# ---------- Read raw inputs --------------------------------------------
Write-Host "[QuotaRecon] Reading inputs..." -ForegroundColor Cyan

$rawSubs = @()
if ($SubscriptionIds) {
    $rawSubs = Read-QRList -Inline $SubscriptionIds -ColumnName 'SubscriptionId'
} elseif ($SubscriptionsCsv) {
    $rawSubs = Read-QRList -CsvPath $SubscriptionsCsv -ColumnName 'SubscriptionId'
} else {
    $answer = Read-Host "Subscription GUIDs (comma separated)"
    $rawSubs = Read-QRList -Inline @($answer) -ColumnName 'SubscriptionId'
}

$rawRegions = @()
if ($Regions) {
    $rawRegions = Read-QRList -Inline $Regions -ColumnName 'Region'
} elseif ($RegionsCsv) {
    $rawRegions = Read-QRList -CsvPath $RegionsCsv -ColumnName 'Region'
} else {
    $answer = Read-Host "Regions (comma separated)"
    $rawRegions = Read-QRList -Inline @($answer) -ColumnName 'Region'
}

$rawSkus = @()
if ($Skus) {
    $rawSkus = Read-QRList -Inline $Skus -ColumnName 'Sku'
} elseif ($SkusCsv) {
    $rawSkus = Read-QRList -CsvPath $SkusCsv -ColumnName 'Sku'
} else {
    $answer = Read-Host "SKUs (comma separated)"
    $rawSkus = Read-QRList -Inline @($answer) -ColumnName 'Sku'
}

if (-not $rawSubs)    { throw "No subscriptions supplied." }
if (-not $rawRegions) { throw "No regions supplied." }
if (-not $rawSkus)    { throw "No SKUs supplied." }

# ---------- Normalize ---------------------------------------------------
Write-Host "[QuotaRecon] Normalizing subscriptions/regions/SKUs..." -ForegroundColor Cyan
$subRes = Resolve-QRSubscriptions -Raw $rawSubs
$validSubs = @($subRes | Where-Object Valid | Select-Object -ExpandProperty Normalized)
if (-not $validSubs) { throw "No valid subscription GUIDs supplied." }

# Resolve display names once, up front. Used everywhere.
Write-Host "[QuotaRecon] Resolving subscription display names..." -ForegroundColor Cyan
$subNames = Get-QRSubscriptionNames -SubscriptionIds $validSubs
# Attach Name to each subscription input row for the Inputs sheet.
$subRes = foreach ($row in $subRes) {
    $name = if ($row.Valid -and $subNames.ContainsKey($row.Normalized)) { $subNames[$row.Normalized] } else { '' }
    $row | Add-Member -NotePropertyName SubscriptionName -NotePropertyValue $name -PassThru -Force
}

$probeSub = $validSubs[0]
$regRes = Resolve-QRRegions -Raw $rawRegions -ProbeSubscriptionId $probeSub
$validRegions = @($regRes | Where-Object Valid | Select-Object -ExpandProperty Normalized | Select-Object -Unique)
if (-not $validRegions) { throw "No valid regions supplied." }

$skuRes = Resolve-QRSkus -Raw $rawSkus -ProbeSubscriptionId $probeSub -Regions $validRegions

# The resolver now returns one row per user-supplied entry, already aggregated
# to family level. If the user supplied both a family and a specific size in
# the same family, prefer the size (SampleSizeSource='user') as the sample.
$validSkuRows = @($skuRes | Where-Object Valid)
if (-not $validSkuRows) { throw "No valid SKUs supplied. See Inputs report." }

# Family -> chosen row (SampleSize + SampleSizeSource + FamilyKey + Raw).
$familyRows = @{}
foreach ($row in $validSkuRows) {
    $key = $row.Family
    if (-not $key) { continue }
    if (-not $familyRows.ContainsKey($key)) {
        $familyRows[$key] = $row
    } elseif ($row.SampleSizeSource -eq 'user' -and $familyRows[$key].SampleSizeSource -ne 'user') {
        # Upgrade auto-picked sample to a user-specified one when both are given.
        $familyRows[$key] = $row
    }
}

$allInputs = @($subRes) + @($regRes) + @($skuRes)

# Output collections are initialized here so both the zone-map stage and the
# per-(sub,region,family) loop below can push into them.
$summary      = New-Object System.Collections.Generic.List[object]
$quotaRows    = New-Object System.Collections.Generic.List[object]
$restrictions = New-Object System.Collections.Generic.List[object]
$errors       = New-Object System.Collections.Generic.List[object]

# ---------- Zone map (first, so Summary can reference physical zones) --
Write-Host "[QuotaRecon] Building zone map from /locations..." -ForegroundColor Cyan
$zoneMap = @()
try {
    $zoneMap = Get-QRZoneMapAllSubs -SubscriptionIds $validSubs -SubscriptionNames $subNames -Regions $validRegions
}
catch {
    $errors.Add([pscustomobject]@{
        Stage = 'ZoneMap'; SubscriptionId = ''; SubscriptionName = ''; Region = ''; Family = ''
        Message = $_.Exception.Message
    })
}
# subId -> region -> (logicalZone -> physicalZone)
$zoneLookup = @{}
foreach ($region in $validRegions) {
    $sub2map = Get-QRPhysicalZonesForRegion -ZoneMapRows $zoneMap -Region $region
    foreach ($k in $sub2map.Keys) {
        if (-not $zoneLookup.ContainsKey($k)) { $zoneLookup[$k] = @{} }
        $zoneLookup[$k][$region] = $sub2map[$k]
    }
}

# ---------- Collect data -----------------------------------------------
Write-Host "[QuotaRecon] Collecting quota + restrictions per (Sub x Region x Family)..." -ForegroundColor Cyan

$summary      = New-Object System.Collections.Generic.List[object]
$quotaRows    = New-Object System.Collections.Generic.List[object]
$restrictions = New-Object System.Collections.Generic.List[object]

foreach ($sub in $validSubs) {
    $subName = $subNames[$sub]
    foreach ($region in $validRegions) {

        # Pull the raw usage rows once per (sub, region) for the Quota sheet.
        try {
            $allUsage = Get-QRAllUsageRows -SubscriptionId $sub -SubscriptionName $subName -Region $region
            foreach ($u in $allUsage) { $quotaRows.Add($u) }
        }
        catch {
            $errors.Add([pscustomobject]@{
                Stage = 'Usages'; SubscriptionId = $sub; SubscriptionName = $subName; Region = $region; Family = ''
                Message = $_.Exception.Message
            })
        }

        # Physical zone map for this (sub, region), used in the Summary column.
        $physicalZonesText = ''
        if ($zoneLookup.ContainsKey($sub) -and $zoneLookup[$sub].ContainsKey($region)) {
            $m = $zoneLookup[$sub][$region]
            $physicalZonesText = (($m.Keys | Sort-Object | ForEach-Object { "$_=$($m[$_])" }) -join '; ')
        }

        foreach ($fam in $familyRows.Keys) {
            $famRow    = $familyRows[$fam]
            $famKey    = $famRow.FamilyKey
            $sample    = $famRow.SampleSize
            $sampleSrc = $famRow.SampleSizeSource

            # Restriction / zone probe: use the sample size's resourceSkus row.
            $skuRow = $null
            try {
                $skuRow = Get-QRSkuRow -SubscriptionId $sub -Region $region -SkuName $sample
            }
            catch {
                $errors.Add([pscustomobject]@{
                    Stage = 'Skus'; SubscriptionId = $sub; SubscriptionName = $subName; Region = $region; Family = $fam
                    Message = $_.Exception.Message
                })
            }

            # Restrictions (from the sample size).
            $skuRestr = @()
            $notAvailForSub = $false
            if ($skuRow) {
                $skuRestr = Get-QRSkuRestrictions -SkuRow $skuRow -SubscriptionId $sub -SubscriptionName $subName -Region $region
                foreach ($r in $skuRestr) { $restrictions.Add($r) }
                $notAvailForSub = @($skuRestr | Where-Object { $_.ReasonCode -eq 'NotAvailableForSubscription' }).Count -gt 0
            }

            # Per-logical-zone restriction status ('None' or 'Yes').
            #   * Any Zone-type restriction flags the listed logical zones.
            #   * A Location-type NotAvailableForSubscription is blanket:
            #     it means the SKU is off for every zone in this region for
            #     this subscription, so we flag all three logical zones AND
            #     set the region-wide Regional-Restrictions cell to 'Yes'.
            $zoneRestricted = @{ '1' = $false; '2' = $false; '3' = $false }
            $regionalRestricted = $false
            $restrLines = New-Object System.Collections.Generic.List[string]
            foreach ($r in $skuRestr) {
                $restrLines.Add((Format-QRRestriction -Restriction $r)) | Out-Null
                if ($r.Zones) {
                    foreach ($z in ($r.Zones -split ',')) {
                        $t = $z.Trim()
                        if ($zoneRestricted.ContainsKey($t)) { $zoneRestricted[$t] = $true }
                    }
                }
                elseif ($r.Type -eq 'Location') {
                    # Location-scoped restriction: region-wide impact.
                    $regionalRestricted = $true
                    foreach ($z in @('1','2','3')) { $zoneRestricted[$z] = $true }
                }
            }
            $restrictionSummaryText = ($restrLines -join '; ')
            $regionalStatus = if ($regionalRestricted) { 'Yes' } else { 'None' }
            $azStatus = @{}
            foreach ($z in @('1','2','3')) {
                $azStatus[$z] = if ($zoneRestricted[$z]) { 'Yes' } else { 'None' }
            }

            # Quota is per-family, looked up directly by the FamilyKey against
            # /usages regardless of whether the sample size exists in this
            # region. Regional 'cores' quota is the umbrella limit.
            $familyUsageRow   = $null
            $regionalUsageRow = $null
            try {
                $usages = Get-QRComputeUsages -SubscriptionId $sub -Region $region
                $regionalUsageRow = $usages | Where-Object { $_.name.value -eq 'cores' } | Select-Object -First 1
                if ($fam) {
                    $familyUsageRow = $usages | Where-Object { $_.name.value -ieq $fam } | Select-Object -First 1
                }
            }
            catch {
                $errors.Add([pscustomobject]@{
                    Stage = 'QuotaMatch'; SubscriptionId = $sub; SubscriptionName = $subName; Region = $region; Family = $fam
                    Message = $_.Exception.Message
                })
            }

            $familyConv   = if ($familyUsageRow)   { Convert-QRUsageRow -Row $familyUsageRow   -SubscriptionId $sub -SubscriptionName $subName -Region $region } else { $null }
            $regionalConv = if ($regionalUsageRow) { Convert-QRUsageRow -Row $regionalUsageRow -SubscriptionId $sub -SubscriptionName $subName -Region $region } else { $null }
            $maxPct = @($familyConv.UsagePercent, $regionalConv.UsagePercent) |
                Where-Object { $_ -ne $null } | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum

            $summary.Add([pscustomobject]@{
                SubscriptionId              = $sub
                SubscriptionName            = $subName
                Region                      = $region
                Family                      = $fam
                FamilyKey                   = $famKey
                SampleSize                  = $sample
                SampleSizeSource            = $sampleSrc
                FamilyUsed                  = $familyConv.CurrentValue
                FamilyLimit                 = $familyConv.Limit
                FamilyPercent               = $familyConv.UsagePercent
                RegionalUsed                = $regionalConv.CurrentValue
                RegionalLimit               = $regionalConv.Limit
                RegionalPercent             = $regionalConv.UsagePercent
                MaxUsagePercent             = $maxPct
                PhysicalZones               = $physicalZonesText
                'Regional-Restrictions'     = $regionalStatus
                'Logical-AZ1-Restrictions'  = $azStatus['1']
                'Logical-AZ2-Restrictions'  = $azStatus['2']
                'Logical-AZ3-Restrictions'  = $azStatus['3']
                RestrictionSummary          = $restrictionSummaryText
            })
        }
    }
}

# ---------- Write workbook ---------------------------------------------
Write-Host "[QuotaRecon] Writing $OutputPath ..." -ForegroundColor Cyan
Write-QRWorkbook `
    -Path         $OutputPath `
    -Summary      $summary `
    -Quota        $quotaRows `
    -Restrictions $restrictions `
    -ZoneMap      $zoneMap `
    -Inputs       $allInputs `
    -Errors       $errors

Write-Host ""
Write-Host "[QuotaRecon] Done." -ForegroundColor Green
Write-Host ("  Subscriptions : {0} valid / {1} supplied" -f $validSubs.Count,    $rawSubs.Count)
Write-Host ("  Regions       : {0} valid / {1} supplied" -f $validRegions.Count, $rawRegions.Count)
Write-Host ("  Families      : {0} valid / {1} supplied" -f $familyRows.Count,   $rawSkus.Count)
Write-Host ("  Summary rows  : {0}" -f $summary.Count)
Write-Host ("  Errors        : {0}" -f $errors.Count)
Write-Host ("  Output        : {0}" -f $OutputPath)

if ($transcriptStarted) {
    try { Stop-Transcript | Out-Null } catch { }
}
