# QuotaRecon.ContainerApps.ps1
#
# Azure Container Apps subscription+region quota (OPT-IN via
# -IncludeContainerAppsQuota). Unlike Microsoft.Compute, Container Apps does
# not expose a single `Microsoft.App/locations/{region}/usages` endpoint.
# Subscription/region-scoped quota (Managed Environment Count, Session
# Pools, Subscription NCA100 GPUs) is instead surfaced through the generic
# Microsoft.Quota resource provider, scoped under the Microsoft.App
# namespace:
#
#   GET .../subscriptions/{sub}/providers/Microsoft.App/locations/{region}
#           /providers/Microsoft.Quota/quotas?api-version=2025-09-01   (limits)
#   GET .../subscriptions/{sub}/providers/Microsoft.App/locations/{region}
#           /providers/Microsoft.Quota/usages?api-version=2025-09-01   (current value)
#
# The two paged calls are merged by `name` into one row per quota item,
# mirroring Convert-QRUsageRow's shape so it reads like the Quota sheet.
#
# Per-workload-profile core quotas (Consumption, General Purpose aka
# D-series, Memory Optimized aka E-series, GPU) live on each Managed
# Environment resource instead (OPT-IN via
# -IncludeContainerAppsEnvironmentQuota, see the bottom of this file):
#
#   GET .../subscriptions/{sub}/resourceGroups/{rg}/providers/Microsoft.App
#           /managedEnvironments/{name}/usages?api-version=2024-03-01
#
# Finding the environments to call requires discovery first (there's no
# sub/region-wide listing endpoint), done via one Azure Resource Graph POST
# across every target subscription - same helper VM Inventory uses
# (Invoke-QRResourceGraph in QuotaRecon.Inventory.ps1).

$script:QRContainerAppsQuotaApiVersion = '2025-09-01'
$script:QRContainerAppsQuotaCache = @{}   # "$sub|$region" -> rows[]

function Get-QRContainerAppsQuota {
    <#
    .SYNOPSIS
        Returns merged limit+usage rows for Azure Container Apps
        subscription/region-scoped quota items (Microsoft.Quota provider,
        Microsoft.App namespace).
    .OUTPUTS
        pscustomobject rows: SubscriptionId, SubscriptionName, Region, Name,
        LocalizedName, Unit, CurrentValue, Limit, UsagePercent.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SubscriptionId,
        [string] $SubscriptionName,
        [Parameter(Mandatory)][string] $Region
    )
    $key = "$SubscriptionId|$Region"
    if ($script:QRContainerAppsQuotaCache.ContainsKey($key)) {
        return $script:QRContainerAppsQuotaCache[$key]
    }

    $basePath   = "/subscriptions/$SubscriptionId/providers/Microsoft.App" +
                  "/locations/$Region/providers/Microsoft.Quota"
    $limitsPath = "$basePath/quotas?api-version=$($script:QRContainerAppsQuotaApiVersion)"
    $usagesPath = "$basePath/usages?api-version=$($script:QRContainerAppsQuotaApiVersion)"

    # NOTE: assign directly, do not wrap in @(...). Invoke-QRArmPaged
    # already returns a comma-protected flat array; re-wrapping the call
    # result in @() double-nests it into a 1-element array-of-arrays, which
    # makes `foreach` iterate once over the *whole* array and turns every
    # `.properties.limit.value` access into a multi-value collection
    # instead of a scalar (surfaces as "Cannot convert System.Object[] to
    # System.Double").
    $limits = Invoke-QRArmPaged -Path $limitsPath
    $usages = Invoke-QRArmPaged -Path $usagesPath

    # Index usages by name for an O(1) merge against the limits list.
    $usageByName = @{}
    foreach ($u in $usages) {
        if ($u.name) { $usageByName[[string]$u.name] = $u }
    }

    $rows = foreach ($l in $limits) {
        $name = [string]$l.name
        $u    = $usageByName[$name]

        $limitValue = $null
        if ($l.properties -and $l.properties.limit -and ($null -ne $l.properties.limit.value)) {
            $limitValue = [double]$l.properties.limit.value
        }
        $usedValue = $null
        if ($u -and $u.properties -and $u.properties.usages -and ($null -ne $u.properties.usages.value)) {
            $usedValue = [double]$u.properties.usages.value
        }
        $pct = if ($limitValue -gt 0 -and $null -ne $usedValue) {
            [math]::Round(($usedValue / $limitValue) * 100, 2)
        } else { $null }

        [pscustomobject]@{
            SubscriptionId   = $SubscriptionId
            SubscriptionName = $SubscriptionName
            Region           = $Region
            Name             = $name
            LocalizedName    = if ($l.properties -and $l.properties.name) { $l.properties.name.localizedValue } else { $name }
            Unit             = if ($l.properties) { $l.properties.unit } else { $null }
            CurrentValue     = $usedValue
            Limit            = $limitValue
            UsagePercent     = $pct
        }
    }

    $rows = @($rows)
    $script:QRContainerAppsQuotaCache[$key] = $rows
    return ,$rows
}

# ---------- Per-environment workload-profile core quota (opt-in) -------

$script:QRContainerAppsEnvUsageApiVersion = '2024-03-01'

function Get-QRContainerAppsEnvironments {
    <#
    .SYNOPSIS
        Discovers every Container Apps managed environment across the given
        subscriptions, filtered to the target regions, via one Azure
        Resource Graph POST (reuses Invoke-QRResourceGraph from
        QuotaRecon.Inventory.ps1 - same pattern as -IncludeVmInventory).
    .OUTPUTS
        pscustomobject rows: SubscriptionId, SubscriptionName, ResourceGroup,
        EnvironmentName, Region, ResourceId. Empty array (not an error) when
        no environments exist in the target regions.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]  $SubscriptionIds,
        [Parameter(Mandatory)][string[]]  $Regions,
        [hashtable] $SubscriptionNames = @{}
    )
    $regionList = @($Regions | ForEach-Object { "'$_'" }) -join ','
    $query = @"
resources
| where type =~ 'microsoft.app/managedenvironments'
| where location in~ ($regionList)
| project SubscriptionId=subscriptionId,
          ResourceGroup=resourceGroup,
          EnvironmentName=name,
          Region=location,
          ResourceId=id
"@
    $rows = Invoke-QRResourceGraph -Query $query -SubscriptionIds $SubscriptionIds

    $out = foreach ($r in $rows) {
        $subName = $SubscriptionNames[$r.SubscriptionId]
        if (-not $subName) { $subName = '(name unknown)' }
        [pscustomobject]@{
            SubscriptionId   = $r.SubscriptionId
            SubscriptionName = $subName
            ResourceGroup    = $r.ResourceGroup
            EnvironmentName  = $r.EnvironmentName
            Region           = $r.Region
            ResourceId       = $r.ResourceId
        }
    }
    return ,@($out)
}

function Get-QRContainerAppsEnvironmentUsage {
    <#
    .SYNOPSIS
        Returns per-workload-profile core quota (Consumption, General
        Purpose aka D-series, Memory Optimized aka E-series, GPU) for one
        Container Apps managed environment.
    .OUTPUTS
        pscustomobject rows: SubscriptionId, SubscriptionName, ResourceGroup,
        EnvironmentName, Region, Name, LocalizedName, Unit, CurrentValue,
        Limit, UsagePercent.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SubscriptionId,
        [string] $SubscriptionName,
        [Parameter(Mandatory)][string] $ResourceGroup,
        [Parameter(Mandatory)][string] $EnvironmentName,
        [Parameter(Mandatory)][string] $Region
    )
    $path = "/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup" +
            "/providers/Microsoft.App/managedEnvironments/$EnvironmentName/usages" +
            "?api-version=$($script:QRContainerAppsEnvUsageApiVersion)"

    # Native RP shape (name/currentValue/limit/unit at the top level), not
    # the generic Microsoft.Quota wrapper used for the sub/region quotas
    # above - same shape Convert-QRUsageRow handles for Microsoft.Compute.
    $rows = Invoke-QRArmPaged -Path $path

    $out = foreach ($u in $rows) {
        $limitValue = if ($null -ne $u.limit)        { [double]$u.limit }        else { $null }
        $usedValue  = if ($null -ne $u.currentValue) { [double]$u.currentValue } else { $null }
        $pct = if ($limitValue -gt 0 -and $null -ne $usedValue) {
            [math]::Round(($usedValue / $limitValue) * 100, 2)
        } else { $null }

        [pscustomobject]@{
            SubscriptionId   = $SubscriptionId
            SubscriptionName = $SubscriptionName
            ResourceGroup    = $ResourceGroup
            EnvironmentName  = $EnvironmentName
            Region           = $Region
            Name             = if ($u.name) { $u.name.value } else { $null }
            LocalizedName    = if ($u.name) { $u.name.localizedValue } else { $null }
            Unit             = $u.unit
            CurrentValue     = $usedValue
            Limit            = $limitValue
            UsagePercent     = $pct
        }
    }
    return ,@($out)
}
