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

    Output is one Excel workbook. Core sheets (always present): Summary,
    Quota, Restrictions, ZoneMap, Inputs, Errors. Additional sheets appear
    only when their opt-in switch is passed: QuotaGroupLimits /
    QuotaGroupAllocations (-IncludeQuotaGroups / -QuotaGroupMgIds),
    VMInventory (-IncludeVmInventory), ContainerAppsQuota /
    ContainerAppsEnvQuota (-IncludeContainerAppsQuota),
    CognitiveServicesQuota (-IncludeCognitiveServicesQuota),
    AppServiceQuota (-IncludeAppServiceQuota). Passing any -Only<Service>
    switch (-OnlyContainerAppsQuota / -OnlyCognitiveServicesQuota /
    -OnlyAppServiceQuota) trims the workbook down to just that
    service's sheet(s) plus Inputs and Errors.

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

.PARAMETER IncludeContainerAppsQuota
    Opt-in: adds BOTH Container Apps quota sheets for every target Sub x
    Region - ContainerAppsQuota (subscription/region-scoped quota: Managed
    Environment Count, Session Pools, Subscription NCA100 GPUs, read via
    the Microsoft.Quota provider under the Microsoft.App namespace) AND
    ContainerAppsEnvQuota (per-managed-environment workload-profile core
    quota: Consumption, General Purpose/D-series, Memory Optimized/
    E-series, GPU - discovered via Azure Resource Graph, one call per
    environment found). If no managed environments are found for the
    specified subscriptions/regions, that's noted on the Errors sheet and
    the console - no error is raised, ContainerAppsQuota is still
    populated regardless.

.PARAMETER OnlyContainerAppsQuota
    Fast-path mode: skips everything except the Container Apps quota check
    (implies -IncludeContainerAppsQuota; you don't need to pass both). SKUs
    are not required and are ignored - SKU resolution, Microsoft.Compute
    provider checks, the zone map, Group Quotas, VM Inventory, and the
    Summary/Quota/Restrictions sheets are all skipped, AND the output
    workbook itself is trimmed down to just four sheets: ContainerAppsQuota,
    ContainerAppsEnvQuota, Inputs, Errors. Only Subscriptions and Regions
    are needed. This -Only<Service> pattern (minimal, purpose-driven, fast)
    is the template for future per-service opt-in switches.

.PARAMETER IncludeCognitiveServicesQuota
    Opt-in: adds a CognitiveServicesQuota sheet with Azure Cognitive
    Services / Azure OpenAI / AI Foundry quota for every target Sub x
    Region - account limits plus every per-model-family TPM/RPM/deployment
    quota item (e.g. OpenAI.ProvisionedManaged,
    OpenAI.GlobalProvisionedManaged, OpenAI.FineTuned.Deployments). Read
    via the native Microsoft.CognitiveServices/locations/{region}/usages
    endpoint - no new module or CLI dependency.

.PARAMETER OnlyCognitiveServicesQuota
    Fast-path mode: skips everything except the Cognitive Services quota
    check (implies -IncludeCognitiveServicesQuota; you don't need to pass
    both). Same skip list as -OnlyContainerAppsQuota, and trims the output
    workbook down to just CognitiveServicesQuota, Inputs, Errors.

.PARAMETER IncludeAppServiceQuota
    Opt-in: adds an AppServiceQuota sheet with Azure App Service quota for
    every target Sub x Region - one row per App Service plan pricing tier
    (F1, D1, B1, B2, B3, S1, P1v3, etc) with its instance-count limit and
    current usage. Read via the Microsoft.Quota provider under the
    Microsoft.Web namespace (App Service has no native
    locations/{region}/usages endpoint) - no new module or CLI dependency.

.PARAMETER OnlyAppServiceQuota
    Fast-path mode: skips everything except the App Service quota check
    (implies -IncludeAppServiceQuota; you don't need to pass both). Same
    skip list as -OnlyContainerAppsQuota, and trims the output workbook
    down to just AppServiceQuota, Inputs, Errors.

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

.EXAMPLE
    # Container Apps quota only, minimal + fast - no SKUs needed.
    ./Invoke-QuotaRecon.ps1 `
        -SubscriptionsCsv ./templates/subscriptions.csv `
        -RegionsCsv       ./templates/regions.csv `
        -OnlyContainerAppsQuota

.EXAMPLE
    # Azure OpenAI / Cognitive Services quota only, minimal + fast.
    ./Invoke-QuotaRecon.ps1 `
        -SubscriptionsCsv ./templates/subscriptions.csv `
        -RegionsCsv       ./templates/regions.csv `
        -OnlyCognitiveServicesQuota

.EXAMPLE
    # App Service quota only, minimal + fast.
    ./Invoke-QuotaRecon.ps1 `
        -SubscriptionsCsv ./templates/subscriptions.csv `
        -RegionsCsv       ./templates/regions.csv `
        -OnlyAppServiceQuota

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

    # Opt-in: adds BOTH Container Apps quota sheets - ContainerAppsQuota
    # (subscription/region-scoped quota via the Microsoft.Quota provider
    # under the Microsoft.App namespace) and ContainerAppsEnvQuota
    # (per-managed-environment workload-profile core quota, discovered via
    # Azure Resource Graph). Same auth, no new module/CLI dependency.
    [switch]   $IncludeContainerAppsQuota,

    # Fast-path mode: skips everything except the Container Apps quota
    # check (implies -IncludeContainerAppsQuota). SKUs aren't required;
    # Microsoft.Compute checks, the zone map, Group Quotas, VM Inventory,
    # and Summary/Quota/Restrictions are all skipped - AND the output
    # workbook is trimmed to just ContainerAppsQuota, ContainerAppsEnvQuota,
    # Inputs, Errors. Template pattern for future -Only<Service> switches:
    # minimal, purpose-driven, fast.
    [switch]   $OnlyContainerAppsQuota,

    # Opt-in: adds a CognitiveServicesQuota sheet with Azure Cognitive
    # Services / Azure OpenAI / AI Foundry quota, read via the native
    # Microsoft.CognitiveServices/locations/{region}/usages endpoint - same
    # row shape as Microsoft.Compute/usages, no new module/CLI dependency.
    [switch]   $IncludeCognitiveServicesQuota,

    # Fast-path mode: skips everything except the Cognitive Services quota
    # check (implies -IncludeCognitiveServicesQuota). Same skip list as
    # -OnlyContainerAppsQuota; trims the workbook to CognitiveServicesQuota,
    # Inputs, Errors.
    [switch]   $OnlyCognitiveServicesQuota,

    # Opt-in: adds an AppServiceQuota sheet with Azure App Service quota
    # (one row per App Service plan pricing tier - F1/D1/B1/S1/P1v3/etc),
    # read via the Microsoft.Quota provider under the Microsoft.Web
    # namespace (App Service has no native locations/{region}/usages
    # endpoint - confirmed via live probe, returns 405). No new
    # module/CLI dependency.
    [switch]   $IncludeAppServiceQuota,

    # Fast-path mode: skips everything except the App Service quota check
    # (implies -IncludeAppServiceQuota). Same skip list as
    # -OnlyContainerAppsQuota; trims the workbook to AppServiceQuota,
    # Inputs, Errors.
    [switch]   $OnlyAppServiceQuota,

    # Prints the QuotaRecon version and exits. Handy for support tickets.
    [switch]   $Version
)

$ErrorActionPreference = 'Stop'

# ---------- Version ----------------------------------------------------
# Bump on every behavioural change. Printed at the top of every run and
# shown by -Version. SemVer: MAJOR.MINOR.PATCH.
$script:QRVersion = '1.7.1'

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
. (Join-Path $here 'src\QuotaRecon.ContainerApps.ps1')
. (Join-Path $here 'src\QuotaRecon.CognitiveServices.ps1')
. (Join-Path $here 'src\QuotaRecon.AppService.ps1')
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

# A single generic flag drives every "skip everything unrelated" gate
# below, regardless of which -Only<Service> switch was passed. Add new
# -Only<Service> switches to this OR-list as they're introduced - see
# .github/copilot-instructions.md for the full pattern.
$anyOnlyModeActive = $OnlyContainerAppsQuota -or $OnlyCognitiveServicesQuota -or $OnlyAppServiceQuota

if ($anyOnlyModeActive) {
    $activeOnlySwitches = New-Object System.Collections.Generic.List[string]
    if ($OnlyContainerAppsQuota)     { $activeOnlySwitches.Add('-OnlyContainerAppsQuota') }
    if ($OnlyCognitiveServicesQuota) { $activeOnlySwitches.Add('-OnlyCognitiveServicesQuota') }
    if ($OnlyAppServiceQuota)        { $activeOnlySwitches.Add('-OnlyAppServiceQuota') }
    Write-Host ("[QuotaRecon] {0}: skipping everything except the specified quota check(s)." -f ($activeOnlySwitches -join ', ')) -ForegroundColor DarkYellow
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
} elseif (-not $IncludeAllSKUs -and -not $anyOnlyModeActive) {
    # Skip the interactive prompt when -IncludeAllSKUs is set; the catalog
    # itself provides the SKU list. Also skipped under any -Only<Service>
    # mode, which doesn't use SKUs at all.
    $answer = Read-Host "SKUs (comma separated)"
    $rawSkus = Read-QRList -Inline @($answer) -ColumnName 'Sku'
}

if (-not $rawSubs)    { throw "No subscriptions supplied." }
if (-not $rawRegions) { throw "No regions supplied." }
if (-not $rawSkus -and -not $IncludeAllSKUs -and -not $anyOnlyModeActive) { throw "No SKUs supplied. (Pass -IncludeAllSKUs to enumerate every family in the target regions.)" }

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
# Skipped entirely under any -Only<Service> mode - compute
# quota isn't collected in that mode, so the check is pure overhead.
$computeProviderState = @{}   # subId -> 'Registered' | 'NotRegistered' | 'Unknown'
$notRegisteredSubs = @()
if (-not $anyOnlyModeActive) {
    Write-Host "[QuotaRecon] Checking Microsoft.Compute provider registration..." -ForegroundColor Cyan
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
if ($IncludeAllSKUs -and -not $anyOnlyModeActive) {
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
if (-not $validSkuRows -and -not $anyOnlyModeActive) { throw "No valid SKUs supplied. See Inputs report." }

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
# Skipped under any -Only<Service> mode (no compute quota is collected in
# that mode, so there's nothing for a zone map to annotate).
$zoneMap = @()
$zoneLookup = @{}
if (-not $anyOnlyModeActive) {
    Write-Host "[QuotaRecon] Building zone map from /locations..." -ForegroundColor Cyan
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
    foreach ($region in $validRegions) {
        $sub2map = Get-QRPhysicalZonesForRegion -ZoneMapRows $zoneMap -Region $region
        foreach ($k in $sub2map.Keys) {
            if (-not $zoneLookup.ContainsKey($k)) { $zoneLookup[$k] = @{} }
            $zoneLookup[$k][$region] = $sub2map[$k]
        }
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
$runQuotaGroups = ($IncludeQuotaGroups -or ($QuotaGroupMgIds -and $QuotaGroupMgIds.Count -gt 0)) -and (-not $SkipQuotaGroups) -and (-not $anyOnlyModeActive)
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
if ($IncludeVmInventory -and -not $anyOnlyModeActive) {
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

# ---------- Container Apps quota (opt-in) -------------------------------
# -IncludeContainerAppsQuota collects BOTH halves of Container Apps quota
# as a matched pair, not two separate opt-ins:
#   * ContainerAppsQuota    - subscription/region-scoped (Microsoft.Quota
#                             provider under the Microsoft.App namespace).
#   * ContainerAppsEnvQuota - per-managed-environment workload-profile core
#                             quota, found by discovering environments via
#                             Resource Graph first, then one usages call
#                             per environment.
# Neither depends on SKUs/families, so this runs standalone rather than
# inside the per-family loop below. If no managed environments are found,
# that's noted on the Errors sheet and the console - not an error.
$containerAppsQuota    = New-Object System.Collections.Generic.List[object]
$containerAppsEnvQuota = New-Object System.Collections.Generic.List[object]
if ($IncludeContainerAppsQuota -or $OnlyContainerAppsQuota) {
    Write-Host "[QuotaRecon] Collecting Container Apps subscription/region quota..." -ForegroundColor Cyan
    foreach ($sub in $validSubs) {
        $subName = $subNames[$sub]
        foreach ($region in $validRegions) {
            try {
                $caRows = Get-QRContainerAppsQuota -SubscriptionId $sub -SubscriptionName $subName -Region $region
                foreach ($r in $caRows) { $containerAppsQuota.Add($r) }
            }
            catch {
                $errors.Add([pscustomobject]@{
                    Stage = 'ContainerAppsQuota'; SubscriptionId = $sub; SubscriptionName = $subName; Region = $region; Family = ''
                    Message = $_.Exception.Message
                })
            }
        }
    }
    Write-Host ("[QuotaRecon]   ContainerAppsQuota rows={0}" -f $containerAppsQuota.Count) -ForegroundColor DarkCyan

    Write-Host "[QuotaRecon] Discovering Container Apps managed environments via Resource Graph..." -ForegroundColor Cyan
    try {
        $caEnvs = Get-QRContainerAppsEnvironments `
            -SubscriptionIds   $validSubs `
            -Regions           $validRegions `
            -SubscriptionNames $subNames

        if (@($caEnvs).Count -eq 0) {
            Write-Host "[QuotaRecon]   no Container Apps managed environments found for the specified subscriptions/regions." -ForegroundColor DarkYellow
            $errors.Add([pscustomobject]@{
                Stage = 'ContainerAppsEnvQuota'; SubscriptionId = ''; SubscriptionName = ''; Region = ''; Family = ''
                Message = 'No Container Apps managed environments were found for the specified subscriptions/regions.'
            })
        } else {
            Write-Host ("[QuotaRecon]   found {0} environment(s); collecting per-environment quota..." -f @($caEnvs).Count) -ForegroundColor DarkCyan
            foreach ($env in $caEnvs) {
                try {
                    $envRows = Get-QRContainerAppsEnvironmentUsage `
                        -SubscriptionId   $env.SubscriptionId `
                        -SubscriptionName $env.SubscriptionName `
                        -ResourceGroup    $env.ResourceGroup `
                        -EnvironmentName  $env.EnvironmentName `
                        -Region           $env.Region
                    foreach ($r in $envRows) { $containerAppsEnvQuota.Add($r) }
                }
                catch {
                    $errors.Add([pscustomobject]@{
                        Stage = 'ContainerAppsEnvQuota'; SubscriptionId = $env.SubscriptionId; SubscriptionName = $env.SubscriptionName
                        Region = $env.Region; Family = ''
                        Message = ("Environment '{0}' (rg '{1}'): {2}" -f $env.EnvironmentName, $env.ResourceGroup, $_.Exception.Message)
                    })
                }
            }
        }
    }
    catch {
        $errors.Add([pscustomobject]@{
            Stage = 'ContainerAppsEnvQuota'; SubscriptionId = ''; SubscriptionName = ''; Region = ''; Family = ''
            Message = $_.Exception.Message
        })
    }
}

# ---------- Cognitive Services quota (opt-in) ---------------------------
# Standalone sub/region loop, same rationale as Container Apps above: this
# doesn't depend on SKUs/families either.
$cognitiveServicesQuota = New-Object System.Collections.Generic.List[object]
if ($IncludeCognitiveServicesQuota -or $OnlyCognitiveServicesQuota) {
    Write-Host "[QuotaRecon] Collecting Cognitive Services / Azure OpenAI quota..." -ForegroundColor Cyan
    foreach ($sub in $validSubs) {
        $subName = $subNames[$sub]
        foreach ($region in $validRegions) {
            try {
                $csRows = Get-QRCognitiveServicesQuota -SubscriptionId $sub -SubscriptionName $subName -Region $region
                foreach ($r in $csRows) { $cognitiveServicesQuota.Add($r) }
            }
            catch {
                $errors.Add([pscustomobject]@{
                    Stage = 'CognitiveServicesQuota'; SubscriptionId = $sub; SubscriptionName = $subName; Region = $region; Family = ''
                    Message = $_.Exception.Message
                })
            }
        }
    }
    Write-Host ("[QuotaRecon]   CognitiveServicesQuota rows={0}" -f $cognitiveServicesQuota.Count) -ForegroundColor DarkCyan
}

# ---------- App Service quota (opt-in) -----------------------------------
# Standalone sub/region loop, same rationale as Container Apps/Cognitive
# Services above: this doesn't depend on SKUs/families either. Quota is
# purely subscription/region scoped (one row per pricing tier), so unlike
# Container Apps there's no secondary per-resource discovery step.
$appServiceQuota = New-Object System.Collections.Generic.List[object]
if ($IncludeAppServiceQuota -or $OnlyAppServiceQuota) {
    Write-Host "[QuotaRecon] Collecting App Service quota..." -ForegroundColor Cyan
    foreach ($sub in $validSubs) {
        $subName = $subNames[$sub]
        foreach ($region in $validRegions) {
            try {
                $asRows = Get-QRAppServiceQuota -SubscriptionId $sub -SubscriptionName $subName -Region $region
                foreach ($r in $asRows) { $appServiceQuota.Add($r) }
            }
            catch {
                $errors.Add([pscustomobject]@{
                    Stage = 'AppServiceQuota'; SubscriptionId = $sub; SubscriptionName = $subName; Region = $region; Family = ''
                    Message = $_.Exception.Message
                })
            }
        }
    }
    Write-Host ("[QuotaRecon]   AppServiceQuota rows={0}" -f $appServiceQuota.Count) -ForegroundColor DarkCyan
}

# ---------- Collect data -----------------------------------------------
# Skipped entirely under any -Only<Service> mode: no SKUs were resolved in
# that mode (so $familyRows is empty anyway), but we also skip the
# Microsoft.Compute /usages pull up front to avoid the wasted ARM calls.
$summary      = New-Object System.Collections.Generic.List[object]
$quotaRows    = New-Object System.Collections.Generic.List[object]
$restrictions = New-Object System.Collections.Generic.List[object]

if (-not $anyOnlyModeActive) {
Write-Host "[QuotaRecon] Collecting quota + restrictions per (Sub x Region x Family)..." -ForegroundColor Cyan

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
}

# ---------- Write workbook ---------------------------------------------
Write-Host "[QuotaRecon] Writing $OutputPath ..." -ForegroundColor Cyan
# Union of sheet names for every -Only<Service> switch that was passed.
# Empty when none were passed, meaning "include everything" (subject to
# the other per-sheet switches, e.g. -IncludeQuotaGroups).
$onlySheetsList = New-Object System.Collections.Generic.List[string]
if ($OnlyContainerAppsQuota)     { $onlySheetsList.AddRange([string[]]@('ContainerAppsQuota', 'ContainerAppsEnvQuota')) }
if ($OnlyCognitiveServicesQuota) { $onlySheetsList.Add('CognitiveServicesQuota') }
if ($OnlyAppServiceQuota)        { $onlySheetsList.Add('AppServiceQuota') }
Write-QRWorkbook `
    -Path                   $OutputPath `
    -Summary                $summary `
    -Quota                  $quotaRows `
    -Restrictions           $restrictions `
    -ZoneMap                $zoneMap `
    -QuotaGroupLimits       $quotaGroupLimits `
    -QuotaGroupAllocations  $quotaGroupAllocations `
    -VmInventory            $vmInventory `
    -ContainerAppsQuota     $containerAppsQuota `
    -ContainerAppsEnvQuota  $containerAppsEnvQuota `
    -CognitiveServicesQuota $cognitiveServicesQuota `
    -AppServiceQuota        $appServiceQuota `
    -Inputs                 $allInputs `
    -Errors                 $errors `
    -IncludeQuotaGroupSheets:$runQuotaGroups `
    -IncludeVmInventorySheet:$IncludeVmInventory `
    -IncludeContainerAppsSheets:($IncludeContainerAppsQuota -or $OnlyContainerAppsQuota) `
    -IncludeCognitiveServicesSheet:($IncludeCognitiveServicesQuota -or $OnlyCognitiveServicesQuota) `
    -IncludeAppServiceSheet:($IncludeAppServiceQuota -or $OnlyAppServiceQuota) `
    -OnlySheets             $onlySheetsList

Write-Host ""
Write-Host "[QuotaRecon] Done." -ForegroundColor Green
Write-Host ("  Subscriptions      : {0} valid / {1} supplied" -f $validSubs.Count,    $rawSubs.Count)
Write-Host ("  Regions            : {0} valid / {1} supplied" -f $validRegions.Count, $rawRegions.Count)
if ($anyOnlyModeActive) {
    Write-Host ("  Mode               : {0} (everything else skipped)" -f ($activeOnlySwitches -join ', ')) -ForegroundColor DarkGray
} else {
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
}
if ($IncludeContainerAppsQuota -or $OnlyContainerAppsQuota) {
    Write-Host ("  Container Apps rows: {0}" -f $containerAppsQuota.Count)
    Write-Host ("  Container Apps env quota rows: {0}" -f $containerAppsEnvQuota.Count)
}
if ($IncludeCognitiveServicesQuota -or $OnlyCognitiveServicesQuota) {
    Write-Host ("  Cognitive Services rows: {0}" -f $cognitiveServicesQuota.Count)
}
if ($IncludeAppServiceQuota -or $OnlyAppServiceQuota) {
    Write-Host ("  App Service rows   : {0}" -f $appServiceQuota.Count)
}
Write-Host ("  Errors             : {0}" -f $errors.Count)
Write-Host ("  Output             : {0}" -f $OutputPath)

if ($transcriptStarted) {
    try { Stop-Transcript | Out-Null } catch { }
}
