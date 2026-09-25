#Requires -Version 5.1
<#
.SYNOPSIS
    Discovery pass for Azure Group Quotas (Microsoft.Quota) — finds the
    Management Groups above each target subscription and dumps every group
    quota it can read into a local JSON file.

.DESCRIPTION
    Group Quotas live at MG scope, not subscription scope. This tool:
      1. Reads the subscriptions from -SubscriptionIds or
         templates/subscriptions.local.csv (or the current Az context).
      2. Calls Microsoft.Management/getEntities once to learn the full
         MG/subscription hierarchy the current identity can see.
      3. For each subscription, walks up its MG ancestor chain.
      4. For each unique MG in every chain, tries the Group Quotas endpoint
         with a couple of api-versions and captures raw responses.
      5. Writes everything to output/quotagroups-discovery-<stamp>.local.json
         so we can look at the real shape before designing the workbook
         columns.

    Nothing here writes to Azure. Read-only reconnaissance.

.PARAMETER SubscriptionIds
    Optional inline list of subscription GUIDs. Falls back to
    templates/subscriptions.local.csv, then to the current Az context.

.PARAMETER OutputPath
    Optional. Defaults to output/quotagroups-discovery-<UTCstamp>.local.json.

.EXAMPLE
    Connect-AzAccount -TenantId <contoso-tenant-id>
    Set-AzContext  -Subscription <a-sub-in-that-tenant>
    ./tools/Discover-QuotaGroups.ps1
#>
[CmdletBinding()]
param(
    [string[]] $SubscriptionIds,
    [string]   $OutputPath,

    # Optional: hard-code the MG(s) to query. If not provided, the tool
    # auto-discovers via getEntities, Get-AzManagementGroup, and (as a
    # last resort) the tenant-root MG (which has the same GUID as the
    # tenant).
    [string[]] $ManagementGroupIds
)

$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $PSCommandPath
$repo = Split-Path -Parent $here
. (Join-Path $repo 'src\QuotaRecon.Auth.ps1')
. (Join-Path $repo 'src\QuotaRecon.Input.ps1')

if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
    throw "Az.Accounts is required. Install with: Install-Module Az.Accounts -Scope CurrentUser"
}
Import-Module Az.Accounts -ErrorAction Stop | Out-Null

if (-not (Get-AzContext -ErrorAction SilentlyContinue)) {
    throw "No Az context. Run Connect-AzAccount before this script."
}
$ctx = Get-AzContext
Write-Host "[Discover] Az context tenant: $($ctx.Tenant.Id)"
Write-Host "[Discover] Az context subscription: $($ctx.Subscription.Id) ($($ctx.Subscription.Name))"

# ---------- Resolve subscription list ----------------------------------
if (-not $SubscriptionIds) {
    $localCsv = Join-Path $repo 'templates\subscriptions.local.csv'
    $publicCsv = Join-Path $repo 'templates\subscriptions.csv'
    if (Test-Path -LiteralPath $localCsv) {
        Write-Host "[Discover] Reading subs from $localCsv"
        $SubscriptionIds = Read-QRList -CsvPath $localCsv -ColumnName 'SubscriptionId'
    } elseif (Test-Path -LiteralPath $publicCsv) {
        Write-Host "[Discover] Reading subs from $publicCsv"
        $SubscriptionIds = Read-QRList -CsvPath $publicCsv -ColumnName 'SubscriptionId'
    }
    if (-not $SubscriptionIds -or $SubscriptionIds.Count -eq 0) {
        $SubscriptionIds = @($ctx.Subscription.Id)
        Write-Host "[Discover] Falling back to current-context subscription."
    }
}
Write-Host ("[Discover] Testing {0} subscription(s)" -f $SubscriptionIds.Count)

# ---------- Pass 1: entity graph via Microsoft.Management/getEntities --
Write-Host "[Discover] Calling /providers/Microsoft.Management/getEntities ..."
$entities = @()
$entitiesRaw = @()   # keep the raw pages for diagnostics
try {
    # getEntities is a paged POST-with-empty-body endpoint but ARM also
    # accepts it as GET for the initial page. Use POST for completeness.
    $entResp = Invoke-QRArm `
        -Path  "/providers/Microsoft.Management/getEntities?api-version=2020-05-01" `
        -Method POST `
        -Body   @{}
    $entitiesRaw += $entResp
    if ($entResp.value) { $entities += $entResp.value }
    $next = $entResp.nextLink
    while ($next) {
        if ($next.StartsWith('https://')) { $next = $next -replace '^https://[^/]+','' }
        $entResp = Invoke-QRArm -Path $next -Method POST -Body @{}
        $entitiesRaw += $entResp
        if ($entResp.value) { $entities += $entResp.value }
        $next = $entResp.nextLink
    }
} catch {
    Write-Warning "getEntities failed: $($_.Exception.Message)"
}
Write-Host ("[Discover] getEntities returned {0} entries" -f $entities.Count)

# Build sub -> ancestor MG chain
$byName = @{}
foreach ($e in $entities) { $byName[$e.name] = $e }

function _AncestorChain([string]$startName) {
    $chain = @()
    $cur   = $byName[$startName]
    $guard = 0
    while ($cur -and $cur.parent -and $cur.parent.id -and $guard -lt 32) {
        # parent.id looks like "/providers/Microsoft.Management/managementGroups/<name>"
        $parentName = ($cur.parent.id -split '/')[-1]
        $chain += $parentName
        $cur = $byName[$parentName]
        $guard++
    }
    return ,$chain
}

$subToMgChain = @{}
foreach ($sub in $SubscriptionIds) {
    $subToMgChain[$sub] = _AncestorChain -startName $sub
    Write-Host ("  {0} -> ancestors: {1}" -f $sub, ($subToMgChain[$sub] -join ' > '))
}

# ---------- Pass 1b: fallback MG enumeration --------------------------
# If getEntities gave us no MG ancestors (common in sandbox/MCAP tenants
# where the caller has sub-level RBAC but not MG-Reader), try two more
# discovery mechanisms so we always have SOMETHING to probe:
#   1. Get-AzManagementGroup (Az PowerShell) - returns MGs visible via
#      Az context caches, sometimes different from raw ARM calls.
#   2. The tenant-root MG - in Azure, the root MG has the same GUID as
#      the tenant. Default group quotas often live here.

$discoveredMgs = New-Object System.Collections.Generic.List[string]

# From ancestor walk
foreach ($chain in $subToMgChain.Values) {
    foreach ($mg in $chain) {
        if ($mg -and -not $discoveredMgs.Contains($mg)) { $discoveredMgs.Add($mg) | Out-Null }
    }
}

# From Az PowerShell
try {
    $azMgs = Get-AzManagementGroup -ErrorAction Stop
    foreach ($mg in $azMgs) {
        if ($mg.Name -and -not $discoveredMgs.Contains($mg.Name)) {
            $discoveredMgs.Add($mg.Name) | Out-Null
        }
    }
    Write-Host ("[Discover] Get-AzManagementGroup returned {0} MG(s)" -f $azMgs.Count)
} catch {
    Write-Host ("[Discover] Get-AzManagementGroup failed: {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow
}

# Tenant-root MG (name == tenantId in Azure)
if ($ctx.Tenant.Id -and -not $discoveredMgs.Contains($ctx.Tenant.Id)) {
    $discoveredMgs.Add($ctx.Tenant.Id) | Out-Null
    Write-Host ("[Discover] Adding tenant-root MG '{0}' to probe list" -f $ctx.Tenant.Id)
}

# Explicit MG IDs from parameter
if ($ManagementGroupIds) {
    foreach ($mg in $ManagementGroupIds) {
        if ($mg -and -not $discoveredMgs.Contains($mg)) { $discoveredMgs.Add($mg) | Out-Null }
    }
}

# ---------- Pass 2: query Group Quotas at each unique MG ---------------
$mgs = @($discoveredMgs)
Write-Host ("[Discover] Unique MGs to probe: {0}" -f $mgs.Count)
foreach ($m in $mgs) { Write-Host "  * $m" }

$apiVersionsToTry = @(
    '2024-10-15-preview',
    '2023-06-01-preview',
    '2023-06-01'
)

$mgResults = @()
foreach ($mg in $mgs) {
    Write-Host "[Discover] MG '$mg' — trying Microsoft.Quota/groupQuotas"
    $mgResult = [ordered]@{
        MgName        = $mg
        Attempts      = @()
        GroupQuotas   = @()
    }
    foreach ($apiver in $apiVersionsToTry) {
        $path = "/providers/Microsoft.Management/managementGroups/$mg/providers/Microsoft.Quota/groupQuotas?api-version=$apiver"
        try {
            # Use paged fetch - the top-level groupQuotas list DOES paginate,
            # and a single-page GET can miss real groups.
            $items = Invoke-QRArmPaged -Path $path
            $count = @($items).Count
            $mgResult.Attempts += @{
                ApiVersion = $apiver
                Path       = $path
                Result     = "OK ($count groupQuotas)"
            }
            $resp = @{ value = $items }
            Write-Host ("  {0} -> {1} groupQuotas" -f $apiver, $count) -ForegroundColor Green
            if ($resp.value) {
                # For each group, fetch details + several sub-resources.
                # The GroupQuotas API is nested under
                #   .../groupQuotas/{name}/resourceProviders/{rp}/...
                # for the interesting bits (limits, allocations).
                foreach ($grp in $resp.value) {
                    $grpDetail = [ordered]@{
                        Name       = $grp.name
                        Id         = $grp.id
                        Type       = $grp.type
                        Properties = $grp.properties
                        SubResources = @()
                    }
                    $grpPath = '/' + $grp.id.TrimStart('/')

                    # Sub-resource URL patterns to probe. Each is optional -
                    # some are RBAC-gated, some region-scoped, some api-version-
                    # specific. We record whichever return data.
                    $subResourcePaths = @(
                        # Member subscriptions in the group
                        ('{0}/subscriptions?api-version={1}' -f $grpPath, $apiver),
                        # Group-level quotas per resource provider (aggregate pool sizes)
                        ('{0}/resourceProviders/Microsoft.Compute/groupQuotaLimits?api-version={1}' -f $grpPath, $apiver),
                        # Same, filtered to a single common region (probe: eastus)
                        ('{0}/resourceProviders/Microsoft.Compute/groupQuotaLimits?api-version={1}&$filter=location eq ''eastus''' -f $grpPath, $apiver),
                        # Older 'quotas' shape (some tenants still respond here)
                        ('{0}/resourceProviders/Microsoft.Compute/quotas?api-version={1}' -f $grpPath, $apiver),
                        # Historical requests against this group
                        ('{0}/resourceProviders/Microsoft.Compute/quotaAllocationRequests?api-version={1}' -f $grpPath, $apiver)
                    )
                    foreach ($srp in $subResourcePaths) {
                        try {
                            $sr = Invoke-QRArm -Path $srp
                            $grpDetail.SubResources += [ordered]@{
                                Path         = $srp
                                Status       = 'OK'
                                ValueCount   = if ($sr.value) { @($sr.value).Count } else { $null }
                                Response     = $sr
                            }
                        } catch {
                            $grpDetail.SubResources += [ordered]@{
                                Path   = $srp
                                Status = 'ERROR'
                                Error  = $_.Exception.Message
                            }
                        }
                    }
                    $mgResult.GroupQuotas += $grpDetail
                }
                break  # first api-version that returned data wins
            }
        } catch {
            $msg = $_.Exception.Message
            $mgResult.Attempts += @{
                ApiVersion = $apiver
                Path       = $path
                Error      = $msg
            }
            Write-Host ("  {0} -> ERROR: {1}" -f $apiver, $msg) -ForegroundColor DarkYellow
        }
    }
    $mgResults += $mgResult
}

# ---------- Persist for review -----------------------------------------
if (-not $OutputPath) {
    $stamp = (Get-Date).ToUniversalTime().ToString('yyyyMMdd-HHmmssZ')
    $OutputPath = Join-Path $repo "output\quotagroups-discovery-$stamp.local.json"
}
$outDir = Split-Path -Parent $OutputPath
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

$payload = [ordered]@{
    Timestamp        = (Get-Date).ToUniversalTime().ToString('u')
    Tenant           = $ctx.Tenant.Id
    SubscriptionIds  = $SubscriptionIds
    SubToMgChain     = $subToMgChain
    DiscoveredMgs    = @($discoveredMgs)
    EntitiesReturned = $entities.Count
    EntitiesSample   = @($entities | Select-Object -First 5)   # for diagnosis
    ManagementGroups = $mgResults
}
$payload | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $OutputPath -Encoding UTF8

Write-Host ""
Write-Host "[Discover] Wrote $OutputPath" -ForegroundColor Cyan
Write-Host "[Discover] Open that file and paste back the interesting bits so we can design the workbook sheets."
