#Requires -Version 5.1
<#
.SYNOPSIS
    Regenerates the two QuotaRecon reference CSVs from live Azure catalogs.

.DESCRIPTION
    Produces:
      references/regions-reference.csv
          ArmName, DisplayName, PhysicalZonesSupported
      references/vm-skus-reference.csv
          Family, FamilyKey, Size, vCPUs, MemoryGB, RegionsOffered

    Both files carry a "# REFERENCE ONLY — not consumed by Invoke-QuotaRecon.ps1"
    header comment so nobody accidentally points the tool at them.

    Uses the current Az context. Any subscription in the target tenant is
    fine — the catalog is per-tenant, not per-sub.

.PARAMETER SubscriptionId
    Optional. The subscription to query. Defaults to the current Az context.

.EXAMPLE
    ./tools/New-QRReferences.ps1
#>
[CmdletBinding()]
param(
    [string] $SubscriptionId
)

$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $PSCommandPath
$repo = Split-Path -Parent $here
. (Join-Path $repo 'src\QuotaRecon.Auth.ps1')

if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
    throw "Az.Accounts is required. Install with: Install-Module Az.Accounts -Scope CurrentUser"
}
Import-Module Az.Accounts -ErrorAction Stop | Out-Null

if (-not $SubscriptionId) {
    $ctx = Get-AzContext -ErrorAction Stop
    if (-not $ctx) { throw "No Az context. Run Connect-AzAccount first." }
    $SubscriptionId = $ctx.Subscription.Id
}
Write-Host "[Refs] Using subscription: $SubscriptionId" -ForegroundColor Cyan

$refDir = Join-Path $repo 'references'
if (-not (Test-Path -LiteralPath $refDir)) {
    New-Item -ItemType Directory -Path $refDir -Force | Out-Null
}

# ---------- Regions ----------------------------------------------------
Write-Host "[Refs] Fetching /locations ..." -ForegroundColor Cyan
$locResp   = Invoke-QRArm -Path "/subscriptions/$SubscriptionId/locations?api-version=2022-12-01"
$regionOut = foreach ($l in ($locResp.value | Sort-Object -Property name)) {
    $zones = if ($l.availabilityZoneMappings) {
        (($l.availabilityZoneMappings | ForEach-Object { "$($_.logicalZone)=$($_.physicalZone)" }) -join ';')
    } else { '' }
    [pscustomobject]@{
        ArmName                = $l.name
        DisplayName            = $l.displayName
        PhysicalZonesSupported = $zones
    }
}

$regionCsv = Join-Path $refDir 'regions-reference.csv'
$header = @(
    '# REFERENCE ONLY - do NOT pass this file to Invoke-QuotaRecon.ps1.',
    '# Regenerate with tools/New-QRReferences.ps1 whenever Azure adds regions.',
    '# Source: GET /subscriptions/{id}/locations?api-version=2022-12-01'
)
$header  | Set-Content -LiteralPath $regionCsv -Encoding UTF8
$regionOut | ConvertTo-Csv -NoTypeInformation | Add-Content -LiteralPath $regionCsv -Encoding UTF8
Write-Host ("[Refs] Wrote {0} regions -> {1}" -f $regionOut.Count, $regionCsv) -ForegroundColor Green

# ---------- SKUs -------------------------------------------------------
Write-Host "[Refs] Fetching Microsoft.Compute/skus (this can take a minute) ..." -ForegroundColor Cyan
# No location filter -> catalog for ALL regions in one paged call. locations[]
# on each row tells us which regions offer that size.
$skuPath = "/subscriptions/$SubscriptionId/providers/Microsoft.Compute/skus?api-version=2021-07-01"
$skus    = Invoke-QRArmPaged -Path $skuPath

$vmSkus = $skus | Where-Object { $_.resourceType -eq 'virtualMachines' }
Write-Host ("[Refs] {0} (Sku x Region) rows -> consolidating by size" -f $vmSkus.Count) -ForegroundColor Cyan

function _CapValue([object]$capabilities, [string]$name) {
    $c = @($capabilities | Where-Object { $_.name -eq $name } | Select-Object -First 1)
    if ($c) { return $c[0].value }
    return ''
}

# Group by SKU name: one row per size, RegionsOffered = every region that lists it.
$bySize = $vmSkus | Group-Object -Property name
$rows = New-Object System.Collections.Generic.List[object]
foreach ($grp in $bySize) {
    $first = $grp.Group[0]
    $familyKey = ''
    if ($first.family) {
        $familyKey = (($first.family -replace '(?i)^standard','' -replace '(?i)family$','') -replace '[\s_\-]','').ToLowerInvariant()
    }
    $regions = New-Object System.Collections.Generic.List[string]
    foreach ($row in $grp.Group) {
        foreach ($loc in @($row.locations)) {
            if ($loc -and -not $regions.Contains([string]$loc)) { $regions.Add([string]$loc) | Out-Null }
        }
    }
    $rows.Add([pscustomobject]@{
        Family         = $first.family
        FamilyKey      = $familyKey
        Size           = $first.name
        vCPUs          = _CapValue $first.capabilities 'vCPUs'
        MemoryGB       = _CapValue $first.capabilities 'MemoryGB'
        RegionsOffered = ((@($regions) | Sort-Object) -join ',')
    }) | Out-Null
}
$skuOut = $rows | Sort-Object -Property Family, Size

$skuCsv = Join-Path $refDir 'vm-skus-reference.csv'
$skuHeader = @(
    '# REFERENCE ONLY - do NOT pass this file to Invoke-QuotaRecon.ps1.',
    '# Regenerate with tools/New-QRReferences.ps1 whenever Azure ships new SKUs.',
    '# Source: GET /providers/Microsoft.Compute/skus?api-version=2021-07-01',
    '# FamilyKey is what you would type in templates/skus.csv to hit a family;',
    '# Size is what you would type to hit a specific size.'
)
$skuHeader | Set-Content -LiteralPath $skuCsv -Encoding UTF8
$skuOut    | ConvertTo-Csv -NoTypeInformation | Add-Content -LiteralPath $skuCsv -Encoding UTF8
Write-Host ("[Refs] Wrote {0} VM SKUs -> {1}" -f $skuOut.Count, $skuCsv) -ForegroundColor Green
