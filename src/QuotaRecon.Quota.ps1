# QuotaRecon.Quota.ps1
#
# Reads Microsoft.Compute/locations/{region}/usages and produces both raw
# rows for the Quota sheet and a matched pair (family counter + regional
# vCPU counter) for the Summary sheet.

$script:QRUsageCache = @{}   # "$sub|$region" -> usages[]

function Get-QRComputeUsages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SubscriptionId,
        [Parameter(Mandatory)][string] $Region
    )
    $key = "$SubscriptionId|$Region"
    if ($script:QRUsageCache.ContainsKey($key)) {
        return $script:QRUsageCache[$key]
    }
    $path = "/subscriptions/$SubscriptionId/providers/Microsoft.Compute" +
            "/locations/$Region/usages?api-version=2023-07-01"
    $rows = Invoke-QRArmPaged -Path $path
    $script:QRUsageCache[$key] = $rows
    return $rows
}

function Convert-QRUsageRow {
    param(
        [Parameter(Mandatory)] $Row,
        [string] $SubscriptionId,
        [string] $SubscriptionName,
        [string] $Region
    )
    $limit = [double]$Row.limit
    $used  = [double]$Row.currentValue
    $pct   = if ($limit -gt 0) { [math]::Round(($used / $limit) * 100, 2) } else { $null }
    [pscustomobject]@{
        SubscriptionId   = $SubscriptionId
        SubscriptionName = $SubscriptionName
        Region           = $Region
        Name             = $Row.name.value
        LocalizedName    = $Row.name.localizedValue
        Unit             = $Row.unit
        CurrentValue     = $used
        Limit            = $limit
        UsagePercent     = $pct
    }
}

function Get-QRQuotaForSku {
    <#
    .SYNOPSIS
        Given a subscription, region, and a resourceSkus row, returns the
        family-vCPU and regional-vCPU usage rows that constrain that SKU.
    .OUTPUTS
        A hashtable with keys `Family` and `Regional` (each is a converted
        row or $null) and `AllFamilyCandidates` for troubleshooting.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SubscriptionId,
        [string] $SubscriptionName,
        [Parameter(Mandatory)][string] $Region,
        [Parameter(Mandatory)] $SkuRow
    )

    $usages = Get-QRComputeUsages -SubscriptionId $SubscriptionId -Region $Region
    $familyKey = Get-QRSkuFamilyKey -SkuObject $SkuRow

    $regional = $usages |
        Where-Object { $_.name.value -eq 'cores' } |
        Select-Object -First 1

    $family = $null
    if ($familyKey) {
        $family = $usages | Where-Object { $_.name.value -ieq $familyKey } | Select-Object -First 1
    }

    return @{
        FamilyKey  = $familyKey
        Family     = if ($family)   { Convert-QRUsageRow -Row $family   -SubscriptionId $SubscriptionId -SubscriptionName $SubscriptionName -Region $Region } else { $null }
        Regional   = if ($regional) { Convert-QRUsageRow -Row $regional -SubscriptionId $SubscriptionId -SubscriptionName $SubscriptionName -Region $Region } else { $null }
    }
}

function Get-QRAllUsageRows {
    <#
    .SYNOPSIS
        Returns every usage row for a sub/region as flattened objects,
        for the raw `Quota` worksheet.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SubscriptionId,
        [string] $SubscriptionName,
        [Parameter(Mandatory)][string] $Region
    )
    $usages = Get-QRComputeUsages -SubscriptionId $SubscriptionId -Region $Region
    return ,($usages | ForEach-Object {
        Convert-QRUsageRow -Row $_ -SubscriptionId $SubscriptionId -SubscriptionName $SubscriptionName -Region $Region
    })
}
