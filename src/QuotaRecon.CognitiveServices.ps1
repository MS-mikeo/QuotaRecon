# QuotaRecon.CognitiveServices.ps1
#
# Azure Cognitive Services / Azure OpenAI / AI Foundry quota (OPT-IN via
# -IncludeCognitiveServicesQuota). Unlike Container Apps, this service has a
# native per-region usages endpoint - the exact same shape
# Microsoft.Compute/usages uses (top-level name.value/localizedValue,
# currentValue, limit, unit) - so it reuses Convert-QRUsageRow directly
# instead of needing its own row converter or a Microsoft.Quota merge:
#
#   GET .../subscriptions/{sub}/providers/Microsoft.CognitiveServices
#           /locations/{region}/usages?api-version=2024-10-01
#
# This single sub/region-scoped call covers everything: account count
# limits plus every deployed model's TPM/RPM-style quota (e.g.
# OpenAI.ProvisionedManaged, OpenAI.GlobalProvisionedManaged,
# OpenAI.FineTuned.Deployments, and per-model-family token/request quotas).
# No SKU/family dependency, so this runs standalone, not nested in the
# per-family loop.

$script:QRCognitiveServicesApiVersion = '2024-10-01'
$script:QRCognitiveServicesUsageCache = @{}   # "$sub|$region" -> usages[]

function Get-QRCognitiveServicesUsages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SubscriptionId,
        [Parameter(Mandatory)][string] $Region
    )
    $key = "$SubscriptionId|$Region"
    if ($script:QRCognitiveServicesUsageCache.ContainsKey($key)) {
        return $script:QRCognitiveServicesUsageCache[$key]
    }
    $path = "/subscriptions/$SubscriptionId/providers/Microsoft.CognitiveServices" +
            "/locations/$Region/usages?api-version=$($script:QRCognitiveServicesApiVersion)"
    $rows = Invoke-QRArmPaged -Path $path
    $script:QRCognitiveServicesUsageCache[$key] = $rows
    return $rows
}

function Get-QRCognitiveServicesQuota {
    <#
    .SYNOPSIS
        Returns every Cognitive Services / Azure OpenAI quota item for one
        (Sub, Region) - account limits plus every per-model-family
        TPM/RPM/deployment quota.
    .OUTPUTS
        pscustomobject rows: SubscriptionId, SubscriptionName, Region, Name,
        LocalizedName, Unit, CurrentValue, Limit, UsagePercent (same shape
        as the Quota sheet, via Convert-QRUsageRow).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SubscriptionId,
        [string] $SubscriptionName,
        [Parameter(Mandatory)][string] $Region
    )
    $usages = Get-QRCognitiveServicesUsages -SubscriptionId $SubscriptionId -Region $Region
    return ,@($usages | ForEach-Object {
        Convert-QRUsageRow -Row $_ -SubscriptionId $SubscriptionId -SubscriptionName $SubscriptionName -Region $Region
    })
}
