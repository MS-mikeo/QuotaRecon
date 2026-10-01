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
    Version is tracked in $script:QRVersion below and printed at the top of
    every run so support tickets reference the exact build the user ran.
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
    [switch]   $DiagnosticLog,

    # Group Quota discovery is OPT-IN as of v1.1.0. Pass -IncludeQuotaGroups
    # to enumerate every Management Group the caller can see and probe each
    # for group quotas. Narrow the probe to specific MGs with
    # -QuotaGroupMgIds (which also implies opt-in). -SkipQuotaGroups is
    # accepted for backward compatibility but is now a no-op because the
    # default is already "skipped".
    [switch]   $IncludeQuotaGroups,
    [string[]] $QuotaGroupMgIds,
    [switch]   $SkipQuotaGroups,

    # Opt-in: run an Azure Resource Graph query to list every deployed VM
    # (and VMSS) whose vmSize is in one of the requested SKU families and
    # whose location is in the target region list. Adds a VMInventory sheet
    # and a DeployedVmCount column to Summary. One extra ARM POST per page.
    [switch]   $IncludeVmInventory,

    # Opt-in: ignore the Skus / SkusCsv input and instead enumerate every
    # virtualMachine family present in the union of the target regions'
    # Microsoft.Compute/skus catalogs. User-supplied SKUs (if any) still
    # participate and 'win' on sample-size when there's overlap. Useful
    # for a wide-angle scan: 'show me everything Azure offers me in these
    # regions, with quota + restrictions.'
    [switch]   $IncludeAllSKUs,

    # Prints the QuotaRecon version and exits. Handy for support tickets.
    [switch]   $Version
)

$ErrorActionPreference = 'Stop'

# ---------- Version ----------------------------------------------------
# Bump on every behavioural change. Printed at the top of every run and
# shown by -Version. SemVer: MAJOR.MINOR.PATCH.
$script:QRVersion = '1.2.0'

if ($Version) {
    Write-Host ("QuotaRecon v{0}" -f $script:QRVersion)
    return
}

Write-Host ("")
Write-Host ("[QuotaRecon] v{0}" -f $script:QRVersion) -ForegroundColor Green

# ---------- Load modules ------------------------------------------------
$here = Split-Path -Parent $PSCommandPath
. (Join-Path $here 'src\QuotaRecon.Auth.ps1')
. (Join-Path $here 'src\QuotaRecon.Input.ps1')
. (Join-Path $here 'src\QuotaRecon.Skus.ps1')
. (Join-Path $here 'src\QuotaRecon.Quota.ps1')
. (Join-Path $here 'src\QuotaRecon.Zones.ps1')
. (Join-Path $here 'src\QuotaRecon.QuotaGroups.ps1')
. (Join-Path $here 'src\QuotaRecon.Inventory.ps1')
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
} elseif (-not $IncludeAllSKUs) {
    # Skip the interactive prompt when -IncludeAllSKUs is set; the catalog
    # itself provides the SKU list.
    $answer = Read-Host "SKUs (comma separated)"
    $rawSkus = Read-QRList -Inline @($answer) -ColumnName 'Sku'
}

if (-not $rawSubs)    { throw "No subscriptions supplied." }
if (-not $rawRegions) { throw "No regions supplied." }
if (-not $rawSkus -and -not $IncludeAllSKUs) { throw "No SKUs supplied. (Pass -IncludeAllSKUs to enumerate every family in the target regions.)" }

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

# Filter out subs the current identity/tenant can't read. Get-QRSubscriptionNames
# stamps '(inaccessible: ...)' when GET /subscriptions/{id} returns 401/403.
# Keeping them in the loop only produces one duplicate Errors row per
# (region x family x stage), which drowns real problems in noise. We drop
# them from the working set here and log one Errors row per skipped sub
# down below (after the $errors list exists).
$skippedSubs = @{}
foreach ($sub in $validSubs) {
    $n = $subNames[$sub]
    if ($n -and $n.StartsWith('(inaccessible')) {
        $skippedSubs[$sub] = $n
    }
}
if ($skippedSubs.Count -gt 0) {
    $validSubs = @($validSubs | Where-Object { -not $skippedSubs.ContainsKey($_) })
    Write-Host ("[QuotaRecon] Skipping {0} inaccessible sub(s) - see Errors sheet." -f $skippedSubs.Count) -ForegroundColor DarkYellow
    if (-not $validSubs) { throw "No accessible subscriptions remain after filtering." }
}

# Check Microsoft.Compute provider registration for each surviving sub. When
# a sub has never provisioned any Azure Compute, the /usages and /skus
# endpoints return zero rows with no error, which shows up as mysteriously
# blank quota columns in Summary. Capture the state so the Summary loop can
# annotate rows with a helpful Note and we can emit one Errors row per sub.
Write-Host "[QuotaRecon] Checking Microsoft.Compute provider registration..." -ForegroundColor Cyan
$computeProviderState = @{}   # subId -> 'Registered' | 'NotRegistered' | 'Unknown'
foreach ($sub in $validSubs) {
    try {
        $reg = Invoke-QRArm -Path "/subscriptions/$sub/providers/Microsoft.Compute`?api-version=2022-12-01"
        $state = if ($reg.registrationState) { [string]$reg.registrationState } else { 'Unknown' }
        $computeProviderState[$sub] = $state
    } catch {
        # If the check itself fails, assume Registered so we don't
        # silently skip real data. Log a verbose note for the diagnostic log.
        Write-Verbose "Provider check failed for $sub`: $($_.Exception.Message)"
        $computeProviderState[$sub] = 'Unknown'
    }
}
$notRegisteredSubs = @($validSubs | Where-Object { $computeProviderState[$_] -eq 'NotRegistered' })
if ($notRegisteredSubs.Count -gt 0) {
    Write-Host ("[QuotaRecon] Microsoft.Compute NotRegistered on {0} sub(s) - Summary rows will be blank with a Note." -f $notRegisteredSubs.Count) -ForegroundColor DarkYellow
}

# For region/SKU catalog probing, prefer a sub with Microsoft.Compute
# actually registered - otherwise Get-QRComputeSkuCatalog and
# Get-QRLocationCatalog will use an empty catalog and produce misleading
# "unknown region/SKU" errors for the whole run.
$registeredSubs = @($validSubs | Where-Object { $computeProviderState[$_] -eq 'Registered' })
$probeSub = if ($registeredSubs.Count -gt 0) { $registeredSubs[0] } else { $validSubs[0] }
$regRes = Resolve-QRRegions -Raw $rawRegions -ProbeSubscriptionId $probeSub
$validRegions = @($regRes | Where-Object Valid | Select-Object -ExpandProperty Normalized | Select-Object -Unique)
if (-not $validRegions) { throw "No valid regions supplied." }

$skuRes = @()
if ($rawSkus -and $rawSkus.Count -gt 0) {
    $skuRes = Resolve-QRSkus -Raw $rawSkus -ProbeSubscriptionId $probeSub -Regions $validRegions
}

# -IncludeAllSKUs: enumerate every virtualMachine family present in the
# union of the target regions' resourceSkus catalogs and append them to
# $skuRes. User-supplied families (if any) stay first in the list so they
# win on sample-size when there's overlap (see dedup loop below).
if ($IncludeAllSKUs) {
    Write-Host "[QuotaRecon] -IncludeAllSKUs: enumerating every family from region catalogs..." -ForegroundColor Cyan
    $catalogSkuRes = Get-QRAllFamiliesFromRegions -ProbeSubscriptionId $probeSub -Regions $validRegions
    Write-Host ("[QuotaRecon]   catalog families found: {0}" -f @($catalogSkuRes).Count) -ForegroundColor DarkCyan
    $skuRes = @($skuRes) + @($catalogSkuRes)
}

# ---------- Call out invalid inputs to the console --------------------
# Users were missing invalid SKUs because they only showed up on the Inputs
# sheet. Repeat them on stdout so a quick skim of the terminal catches
# typos before anyone opens the workbook.
$invalidSubRows    = @($subRes | Where-Object { -not $_.Valid })
$invalidRegionRows = @($regRes | Where-Object { -not $_.Valid })
$invalidSkuRows    = @($skuRes | Where-Object { -not $_.Valid })
if ($invalidSubRows.Count -gt 0) {
    Write-Warning ("[QuotaRecon] {0} invalid subscription(s) ignored (see Inputs sheet):" -f $invalidSubRows.Count)
    foreach ($row in $invalidSubRows) {
        Write-Host ("    - '{0}' -> {1}" -f $row.Raw, $row.Reason) -ForegroundColor Yellow
    }
}
if ($invalidRegionRows.Count -gt 0) {
    Write-Warning ("[QuotaRecon] {0} invalid region(s) ignored (see Inputs sheet):" -f $invalidRegionRows.Count)
    foreach ($row in $invalidRegionRows) {
        Write-Host ("    - '{0}' -> {1}" -f $row.Raw, $row.Reason) -ForegroundColor Yellow
    }
}
if ($invalidSkuRows.Count -gt 0) {
    Write-Warning ("[QuotaRecon] {0} invalid SKU/family entry(ies) ignored (see Inputs sheet):" -f $invalidSkuRows.Count)
    foreach ($row in $invalidSkuRows) {
        Write-Host ("    - '{0}' -> {1}" -f $row.Raw, $row.Reason) -ForegroundColor Yellow
    }
}

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

# One Errors row per skipped sub - much cleaner than one per operation.
foreach ($skippedId in $skippedSubs.Keys) {
    $errors.Add([pscustomobject]@{
        Stage            = 'SkippedInaccessibleSub'
        SubscriptionId   = $skippedId
        SubscriptionName = $skippedSubs[$skippedId]
        Region           = ''
        Family           = ''
        Message          = 'Subscription is inaccessible to the current identity; all downstream calls skipped.'
    })
}

# One Errors row per sub whose Microsoft.Compute provider is NotRegistered.
# These subs still appear in Summary/ZoneMap/QuotaGroups (they may be members
# of a group), but every /usages and /skus probe returns empty.
foreach ($notRegId in $notRegisteredSubs) {
    $errors.Add([pscustomobject]@{
        Stage            = 'ComputeProviderNotRegistered'
        SubscriptionId   = $notRegId
        SubscriptionName = $subNames[$notRegId]
        Region           = ''
        Family           = ''
        Message          = ("Microsoft.Compute provider state is 'NotRegistered' on this subscription; usage/SKU probes return empty and Summary rows are blank. Fix: 'Register-AzResourceProvider -ProviderNamespace Microsoft.Compute' (needs Contributor or higher on the sub).")
    })
}

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

# ---------- Group Quotas (opt-in) --------------------------------------
# As of v1.1.0 Group Quota discovery is OPT-IN. Pass -IncludeQuotaGroups to
# enumerate every MG the caller can see and probe each, or -QuotaGroupMgIds
# to narrow the probe to specific MGs (which also implies opt-in). After
# collection, every unique sub ID referenced by any group is added to the
# name resolver so allocation rows carry real names even for subs the user
# didn't put in the input list. -SkipQuotaGroups is kept as a no-op for
# backward compat.
$quotaGroupGroups      = @()
$quotaGroupLimits      = @()
$quotaGroupAllocations = @()
$quotaGroupIndex       = @{}
$runQuotaGroups = ($IncludeQuotaGroups -or ($QuotaGroupMgIds -and $QuotaGroupMgIds.Count -gt 0)) -and (-not $SkipQuotaGroups)
if ($SkipQuotaGroups -and $IncludeQuotaGroups) {
    Write-Warning "[QuotaRecon] -SkipQuotaGroups overrides -IncludeQuotaGroups; group quotas will NOT be collected."
}
if ($runQuotaGroups) {
    $mgProbeList = @()
    if ($QuotaGroupMgIds -and $QuotaGroupMgIds.Count -gt 0) {
        $mgProbeList = @($QuotaGroupMgIds)
        Write-Host ("[QuotaRecon] Group Quotas: probing {0} MG(s) from -QuotaGroupMgIds" -f $mgProbeList.Count) -ForegroundColor Cyan
    } else {
        Write-Host "[QuotaRecon] Group Quotas: auto-discovering Management Groups..." -ForegroundColor Cyan
        $mgProbeList = @(Get-QRAllManagementGroups)
        # Ensure the tenant-root MG is in the list too (some contexts omit it).
        $tenantId = (Get-AzContext).Tenant.Id
        if ($tenantId -and ($mgProbeList -notcontains $tenantId)) {
            $mgProbeList = @($mgProbeList) + @($tenantId)
        }
        Write-Host ("[QuotaRecon]   found {0} MG(s) to probe" -f $mgProbeList.Count) -ForegroundColor DarkCyan
    }

    if ($mgProbeList.Count -gt 0) {
        try {
            $qgResult = Get-QRAllGroupQuotaData `
                -ManagementGroupIds $mgProbeList `
                -Regions            $validRegions

            # Phase 2: resolve names for every sub referenced by any group
            # that we didn't already resolve for the input set.
            $extraSubIds = Get-QRSubIdsInGroupData -QuotaGroupData $qgResult |
                Where-Object { $_ -and -not $subNames.ContainsKey($_) }
            if ($extraSubIds -and @($extraSubIds).Count -gt 0) {
                Write-Host ("[QuotaRecon]   resolving names for {0} extra sub(s) referenced by groups" -f @($extraSubIds).Count) -ForegroundColor DarkCyan
                $extraNames = Get-QRSubscriptionNames -SubscriptionIds @($extraSubIds)
                foreach ($k in $extraNames.Keys) { $subNames[$k] = $extraNames[$k] }
            }
            Set-QRSubscriptionNamesOnGroupData -QuotaGroupData $qgResult -SubscriptionNames $subNames

            $quotaGroupGroups      = $qgResult.Groups
            $quotaGroupLimits      = $qgResult.Limits
            $quotaGroupAllocations = $qgResult.Allocations
            $quotaGroupIndex       = Build-QRQuotaGroupSummaryIndex -Allocations $quotaGroupAllocations
            Write-Host ("[QuotaRecon]   groups={0}  limits={1}  allocations={2}" -f `
                $quotaGroupGroups.Count, $quotaGroupLimits.Count, $quotaGroupAllocations.Count) -ForegroundColor DarkCyan
        }
        catch {
            $errors.Add([pscustomobject]@{
                Stage = 'QuotaGroups'; SubscriptionId = ''; SubscriptionName = ''; Region = ''; Family = ''
                Message = $_.Exception.Message
            })
        }
    }
}

# ---------- VM Inventory (opt-in) --------------------------------------
# Runs an Azure Resource Graph query to list every VM / VMSS whose vmSize
# is in one of the requested families and whose location is in the target
# region list. Also builds a (sub,region,familyKey) -> instance-count index
# so the Summary sheet can carry a DeployedVmCount column.
$vmInventory = @()
$vmCountIndex = @{}
if ($IncludeVmInventory) {
    Write-Host "[QuotaRecon] Collecting VM inventory via Resource Graph..." -ForegroundColor Cyan
    try {
        $familyKeys = @($familyRows.Values | ForEach-Object { $_.FamilyKey } | Where-Object { $_ } | Select-Object -Unique)
        $famSizeMap = Get-QRFamilySizeMap `
            -ProbeSubscriptionId $probeSub `
            -Regions             $validRegions `
            -FamilyKeys          $familyKeys
        $vmInventory = Get-QRVmInventory `
            -SubscriptionIds     $validSubs `
            -Regions             $validRegions `
            -FamilySizeMap       $famSizeMap `
            -SubscriptionNames   $subNames
        $vmCountIndex = Build-QRVmInventoryCountIndex -Inventory $vmInventory
        Write-Host ("[QuotaRecon]   inventory rows={0}" -f @($vmInventory).Count) -ForegroundColor DarkCyan
    }
    catch {
        $errors.Add([pscustomobject]@{
            Stage = 'VmInventory'; SubscriptionId = ''; SubscriptionName = ''; Region = ''; Family = ''
            Message = $_.Exception.Message
        })
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

            # Family-aware probe:
            #   1. Pull the full resourceSkus catalog for this (sub, region).
            #   2. Filter to every size whose family key matches $famKey.
            #   3. If zero sizes offered -> family not available in this
            #      region at all (common for Confidential Compute, M-series,
            #      HPC families in secondary regions).
            #   4. Otherwise prefer the pre-picked sample; if that specific
            #      size isn't in this region's catalog but a sibling size
            #      is, fall back to the smallest-vCPU sibling as the probe
            #      and record the swap in the Note column.
            $skuRow           = $null
            $effectiveSample  = $sample
            $sampleFallback   = $false
            $offeredCount     = 0
            $offeredNames     = @()
            try {
                $catalog = Get-QRComputeSkuCatalog -SubscriptionId $sub -Region $region
                $famOfferings = @($catalog | Where-Object {
                    $_.resourceType -eq 'virtualMachines' -and
                    $_.family -and
                    (Get-QRFamilyLookupKey -Family $_.family) -eq $famKey
                })
                $offeredCount = $famOfferings.Count
                $offeredNames = @($famOfferings | ForEach-Object { $_.name } | Sort-Object)

                if ($offeredCount -gt 0) {
                    $skuRow = $famOfferings | Where-Object { $_.name -eq $sample } | Select-Object -First 1
                    if (-not $skuRow) {
                        $skuRow = $famOfferings | Sort-Object `
                            @{Expression={
                                $v = @($_.capabilities | Where-Object { $_.name -eq 'vCPUs' } | Select-Object -First 1)
                                if ($v) { [int]$v[0].value } else { [int]::MaxValue }
                            }}, @{Expression={ $_.name }} | Select-Object -First 1
                        if ($skuRow) {
                            $effectiveSample = $skuRow.name
                            $sampleFallback  = $true
                        }
                    }
                }
            }
            catch {
                $errors.Add([pscustomobject]@{
                    Stage = 'Skus'; SubscriptionId = $sub; SubscriptionName = $subName; Region = $region; Family = $fam
                    Message = $_.Exception.Message
                })
            }

            # Restrictions (from the effective sample size).
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
            # Decide restriction cell values:
            #   * $offeredCount == 0 -> family not offered in this region
            #     at all. All restriction cells = 'N/A' so the Summary CF
            #     leaves them uncoloured; the Note explains why.
            #   * Otherwise walk each logical zone:
            #       - zone not in locationInfo[].zones -> 'N/A' (SKU is not
            #         deployable in that zone at all; Azure doesn't emit a
            #         restrictions[] row for it, it's just absent).
            #       - zone offered and listed in a Zone restriction -> 'Yes'
            #       - zone offered and unrestricted -> 'None'
            #     Regional status is 'Yes' if ANY Location-type restriction
            #     fires, 'None' otherwise.
            $azStatus = @{}
            if ($offeredCount -eq 0) {
                $regionalStatus = 'N/A'
                foreach ($z in @('1','2','3')) { $azStatus[$z] = 'N/A' }
            } else {
                $regionalStatus = if ($regionalRestricted) { 'Yes' } else { 'None' }
                # Build the set of logical zones this probe row is actually
                # offered in. For non-AZ regions or non-zonal SKUs this will
                # be empty and all three zone cells read 'N/A'.
                $offeredZones = @{}
                foreach ($li in @($skuRow.locationInfo)) {
                    foreach ($z in @($li.zones)) {
                        if ($z) { $offeredZones[[string]$z] = $true }
                    }
                }
                foreach ($z in @('1','2','3')) {
                    if (-not $offeredZones.ContainsKey($z)) {
                        $azStatus[$z] = 'N/A'
                    } elseif ($zoneRestricted[$z]) {
                        $azStatus[$z] = 'Yes'
                    } else {
                        $azStatus[$z] = 'None'
                    }
                }
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

            # Look up group-quota participation for this (sub, region, family).
            $qgKey = ('{0}|{1}|{2}' -f $sub, $region.ToLowerInvariant(), $famKey.ToLowerInvariant())
            $qgRec = $quotaGroupIndex[$qgKey]
            $inGroup            = if ($qgRec) { 'Yes' } else { 'No' }
            $groupNames         = if ($qgRec) { ($qgRec.GroupNames -join ';') } else { '' }
            $groupPoolLimit     = if ($qgRec) { $qgRec.PoolLimit }     else { $null }
            $groupPoolAvailable = if ($qgRec) { $qgRec.PoolAvailable } else { $null }
            $subAllocated       = if ($qgRec) { $qgRec.SubAllocated }  else { $null }

            # Deployed VM count for this (sub, region, family), 0 if the
            # -IncludeVmInventory switch wasn't passed or nothing deployed.
            $vmCount = 0
            if ($vmCountIndex.ContainsKey($qgKey)) { $vmCount = $vmCountIndex[$qgKey] }

            # Compose a Note explaining any missing data on this row.
            $noteParts = New-Object System.Collections.Generic.List[string]
            if ($computeProviderState[$sub] -eq 'NotRegistered') {
                $noteParts.Add('Microsoft.Compute provider NotRegistered on this subscription') | Out-Null
            }
            if ($offeredCount -eq 0 -and $computeProviderState[$sub] -ne 'NotRegistered') {
                $noteParts.Add(("Family '{0}' has no sizes offered in {1}" -f $fam, $region)) | Out-Null
            }
            elseif ($sampleFallback) {
                $noteParts.Add(("Sample size '{0}' not offered in {1}; probed '{2}' instead" -f $sample, $region, $effectiveSample)) | Out-Null
            }
            if ($notAvailForSub) {
                $noteParts.Add('NotAvailableForSubscription for this SKU') | Out-Null
            }
            $noteText = ($noteParts -join '; ')

            $summary.Add([pscustomobject]@{
                SubscriptionId              = $sub
                SubscriptionName            = $subName
                Region                      = $region
                Family                      = $fam
                FamilyKey                   = $famKey
                SampleSize                  = $sample
                SampleSizeSource            = $sampleSrc
                ProbeSampleUsed             = $effectiveSample
                OfferedSizesInRegion        = $offeredCount
                Note                        = $noteText
                FamilyUsed                  = $familyConv.CurrentValue
                FamilyLimit                 = $familyConv.Limit
                FamilyPercent               = $familyConv.UsagePercent
                RegionalUsed                = $regionalConv.CurrentValue
                RegionalLimit               = $regionalConv.Limit
                RegionalPercent             = $regionalConv.UsagePercent
                MaxUsagePercent             = $maxPct
                InQuotaGroup                = $inGroup
                QuotaGroupNames             = $groupNames
                QuotaGroupPoolLimit         = $groupPoolLimit
                QuotaGroupPoolAvailable     = $groupPoolAvailable
                SubAllocationInGroup        = $subAllocated
                DeployedVmCount             = $vmCount
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
    -Path                  $OutputPath `
    -Summary               $summary `
    -Quota                 $quotaRows `
    -Restrictions          $restrictions `
    -ZoneMap               $zoneMap `
    -QuotaGroupLimits      $quotaGroupLimits `
    -QuotaGroupAllocations $quotaGroupAllocations `
    -VmInventory           $vmInventory `
    -Inputs                $allInputs `
    -Errors                $errors

Write-Host ""
Write-Host "[QuotaRecon] Done." -ForegroundColor Green
Write-Host ("  Subscriptions      : {0} valid / {1} supplied" -f $validSubs.Count,    $rawSubs.Count)
Write-Host ("  Regions            : {0} valid / {1} supplied" -f $validRegions.Count, $rawRegions.Count)
if ($IncludeAllSKUs) {
    Write-Host ("  Families           : {0} valid ({1} user-supplied + catalog-enumerated)" -f $familyRows.Count, @($rawSkus).Count)
} else {
    Write-Host ("  Families           : {0} valid / {1} supplied" -f $familyRows.Count,   $rawSkus.Count)
}
Write-Host ("  Summary rows       : {0}" -f $summary.Count)
if ($runQuotaGroups) {
    Write-Host ("  Quota groups       : {0}" -f $quotaGroupGroups.Count)
    Write-Host ("  Group limit rows   : {0}" -f $quotaGroupLimits.Count)
    Write-Host ("  Group alloc rows   : {0}" -f $quotaGroupAllocations.Count)
} else {
    Write-Host ("  Quota groups       : skipped (pass -IncludeQuotaGroups to collect)") -ForegroundColor DarkGray
}
if ($IncludeVmInventory) {
    Write-Host ("  VM inventory rows  : {0}" -f @($vmInventory).Count)
}
Write-Host ("  Errors             : {0}" -f $errors.Count)
Write-Host ("  Output             : {0}" -f $OutputPath)

if ($transcriptStarted) {
    try { Stop-Transcript | Out-Null } catch { }
}
