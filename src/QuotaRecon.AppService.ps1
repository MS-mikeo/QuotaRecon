# QuotaRecon.AppService.ps1
#
# Azure App Service quota (OPT-IN via -IncludeAppServiceQuota). Like
# Container Apps, App Service has no native
# `Microsoft.Web/locations/{region}/usages` endpoint (confirmed live: that
# path returns 405 Method Not Allowed) - quota is surfaced only through the
# generic Microsoft.Quota resource provider, scoped under the Microsoft.Web
# namespace:
#
#   GET .../subscriptions/{sub}/providers/Microsoft.Web/locations/{region}
#           /providers/Microsoft.Quota/quotas?api-version=2025-09-01   (limits)
#   GET .../subscriptions/{sub}/providers/Microsoft.Web/locations/{region}
#           /providers/Microsoft.Quota/usages?api-version=2025-09-01   (current value)
#
# Unlike Container Apps, App Service quota is purely subscription/region
# scoped - one row per App Service plan pricing tier (F1, D1, B1, B2, B3,
# S1, P1v3, etc - each tier's instance-count limit), not per-resource. So
# there's no environment-discovery step and no second sheet - this is the
# simplest service onboarded so far: one standalone collection loop,
# one merged limit+usage call per (Sub, Region), one sheet.

$script:QRAppServiceQuotaApiVersion = '2025-09-01'
$script:QRAppServiceQuotaCache = @{}   # "$sub|$region" -> rows[]

function Get-QRAppServiceQuota {
    <#
    .SYNOPSIS
        Returns merged limit+usage rows for Azure App Service quota items
        (Microsoft.Quota provider, Microsoft.Web namespace) - one row per
        App Service plan pricing tier (F1, D1, B1, S1, P1v3, etc).
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
    if ($script:QRAppServiceQuotaCache.ContainsKey($key)) {
        return $script:QRAppServiceQuotaCache[$key]
    }

    $basePath   = "/subscriptions/$SubscriptionId/providers/Microsoft.Web" +
                  "/locations/$Region/providers/Microsoft.Quota"
    $limitsPath = "$basePath/quotas?api-version=$($script:QRAppServiceQuotaApiVersion)"
    $usagesPath = "$basePath/usages?api-version=$($script:QRAppServiceQuotaApiVersion)"

    # NOTE: assign directly, do not wrap in @(...) - see
    # QuotaRecon.ContainerApps.ps1 for why re-wrapping a comma-protected
    # array return breaks scalar property access further down.
    $limits = Invoke-QRArmPaged -Path $limitsPath
    $usages = Invoke-QRArmPaged -Path $usagesPath

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
    $script:QRAppServiceQuotaCache[$key] = $rows
    return ,$rows
}
