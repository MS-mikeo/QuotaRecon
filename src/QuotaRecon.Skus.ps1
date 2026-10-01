# QuotaRecon.Skus.ps1
#
# Wraps Microsoft.Compute/skus. Caches the full list per (SubscriptionId,
# Region) because the endpoint is region-scoped and slow.

$script:QRSkuCatalogCache = @{}   # "$sub|$region" -> resourceSkus[]

function Get-QRComputeSkuCatalog {
    <#
    .SYNOPSIS
        Returns the raw resourceSkus[] for a subscription and region.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SubscriptionId,
        [Parameter(Mandatory)][string] $Region
    )
    $key = "$SubscriptionId|$Region"
    if ($script:QRSkuCatalogCache.ContainsKey($key)) {
        return $script:QRSkuCatalogCache[$key]
    }

    # Filter server-side by location to keep responses small.
    $path = "/subscriptions/$SubscriptionId/providers/Microsoft.Compute/skus" +
            "?api-version=2021-07-01&`$filter=location eq '$Region'"
    $rows = Invoke-QRArmPaged -Path $path
    $script:QRSkuCatalogCache[$key] = $rows
    return $rows
}

function Get-QRSkuFamilyKey {
    <#
    .SYNOPSIS
        Derives the compute usage-family key ("standardDSv5Family") from a
        SKU name. Uses the capabilities the SKU catalog returns when possible,
        with a regex fallback.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $SkuObject   # a row from resourceSkus
    )
    # resourceSkus has a `family` field like "standardDSv5Family".
    if ($SkuObject.family) { return $SkuObject.family }

    # Fallback: parse name.
    if ($SkuObject.name -match '^Standard_([A-Z]+)(\d+)(-\d+)?([a-z]*)_v(\d+)$') {
        $letters = $matches[1].ToUpperInvariant()
        $ver     = $matches[5]
        # This will not match every family (M, HB, HC, NDv4 etc.) so callers
        # should still cross-check against the usages list.
        return "standard${letters}v${ver}Family"
    }
    return $null
}

function Get-QRSkuRow {
    <#
    .SYNOPSIS
        Returns the resourceSkus row for a specific SKU + region + sub, or
        $null if the SKU isn't offered there at all.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SubscriptionId,
        [Parameter(Mandatory)][string] $Region,
        [Parameter(Mandatory)][string] $SkuName
    )
    $catalog = Get-QRComputeSkuCatalog -SubscriptionId $SubscriptionId -Region $Region
    return $catalog | Where-Object {
        $_.resourceType -eq 'virtualMachines' -and $_.name -eq $SkuName
    } | Select-Object -First 1
}

function Get-QRSkuRestrictions {
    <#
    .SYNOPSIS
        Flattens the resourceSkus[].restrictions[] array into one row per
        restriction. Locations and Zones are captured separately so callers
        can render an az-CLI-style summary line:
            "<ReasonCode>, type: <Type>, locations: <locs>, zones: <zones>"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $SkuRow,
        [Parameter(Mandatory)][string] $SubscriptionId,
        [string] $SubscriptionName,
        [Parameter(Mandatory)][string] $Region
    )
    if (-not $SkuRow -or -not $SkuRow.restrictions) { return @() }

    $out = foreach ($r in $SkuRow.restrictions) {
        $locs  = if ($r.restrictionInfo.locations) { ($r.restrictionInfo.locations -join ',') } else { '' }
        $zones = if ($r.restrictionInfo.zones)     { ($r.restrictionInfo.zones     -join ',') } else { '' }
        [pscustomobject]@{
            SubscriptionId   = $SubscriptionId
            SubscriptionName = $SubscriptionName
            Region           = $Region
            Sku              = $SkuRow.name
            Type             = $r.type
            Locations        = $locs
            Zones            = $zones
            ReasonCode       = $r.reasonCode
        }
    }
    return ,$out
}

function Format-QRRestriction {
    <#
    .SYNOPSIS
        Renders one restriction row as an az-CLI-style single line:
            "NotAvailableForSubscription, type: Zone, locations: eastus2, zones: 3,1,2"
    #>
    param([Parameter(Mandatory)] $Restriction)

    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add($Restriction.ReasonCode) | Out-Null
    $parts.Add("type: $($Restriction.Type)") | Out-Null
    if ($Restriction.Locations) { $parts.Add("locations: $($Restriction.Locations)") | Out-Null }
    if ($Restriction.Zones)     { $parts.Add("zones: $($Restriction.Zones)")         | Out-Null }
    return ($parts -join ', ')
}

function Get-QRSkuZonesAvailable {
    <#
    .SYNOPSIS
        Returns the set of logical zones ('1','2','3') the SKU is *actually*
        usable in for this subscription: offered zones minus zone restrictions.
        Always returns a flat [string[]] (never nested, never a scalar) so
        callers can safely `-join` it.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] $SkuRow)

    if (-not $SkuRow) { return ,([string[]]@()) }

    $offered = New-Object System.Collections.Generic.List[string]
    foreach ($li in @($SkuRow.locationInfo)) {
        foreach ($z in @($li.zones)) {
            if ($z -and -not $offered.Contains([string]$z)) { $offered.Add([string]$z) | Out-Null }
        }
    }

    $restrictedZones = New-Object System.Collections.Generic.List[string]
    $notAvailable    = $false
    foreach ($r in @($SkuRow.restrictions)) {
        if ($r.type -eq 'Zone' -and $r.restrictionInfo.zones) {
            foreach ($z in @($r.restrictionInfo.zones)) {
                if ($z) { $restrictedZones.Add([string]$z) | Out-Null }
            }
        }
        if ($r.reasonCode -eq 'NotAvailableForSubscription' -and $r.type -eq 'Location') {
            $notAvailable = $true
        }
    }
    if ($notAvailable) { return ,([string[]]@()) }

    $result = [string[]]@($offered | Where-Object { $restrictedZones -notcontains $_ } | Sort-Object)
    return ,$result
}

function Get-QRAllFamiliesFromRegions {
    <#
    .SYNOPSIS
        Enumerates every virtualMachine family present in the union of the
        resourceSkus catalogs for the supplied regions, and returns one row
        per unique family in the same shape Resolve-QRSkus produces.
        Used by Invoke-QuotaRecon's -IncludeAllSKUs switch to report on
        everything Azure offers in the target regions without a user list.

    .OUTPUTS
        pscustomobject rows with:
            Kind='Sku', Raw='(all-skus)', Family, FamilyKey, SampleSize,
            SampleSizeSource='catalog', Valid=$true, Reason
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]   $ProbeSubscriptionId,
        [Parameter(Mandatory)][string[]] $Regions
    )

    $familyRaw     = @{}  # familyKey -> raw family string
    $familyByKey   = @{}  # familyKey -> List[string] of size names
    $vcpuOfSize    = @{}  # size name  -> vCPU count

    foreach ($region in $Regions) {
        $catalog = Get-QRComputeSkuCatalog -SubscriptionId $ProbeSubscriptionId -Region $region
        foreach ($sku in $catalog) {
            if ($sku.resourceType -ne 'virtualMachines') { continue }
            if (-not $sku.family) { continue }
            $fk = Get-QRFamilyLookupKey -Family $sku.family
            if (-not $familyRaw.ContainsKey($fk))   { $familyRaw[$fk]   = $sku.family }
            if (-not $familyByKey.ContainsKey($fk)) {
                $familyByKey[$fk] = New-Object System.Collections.Generic.List[string]
            }
            if ($familyByKey[$fk] -notcontains $sku.name) {
                $familyByKey[$fk].Add($sku.name) | Out-Null
            }
            if (-not $vcpuOfSize.ContainsKey($sku.name)) {
                $vCap = @($sku.capabilities | Where-Object { $_.name -eq 'vCPUs' } | Select-Object -First 1)
                $vcpuOfSize[$sku.name] = if ($vCap) { [int]$vCap[0].value } else { [int]::MaxValue }
            }
        }
    }

    $out = New-Object System.Collections.Generic.List[object]
    foreach ($fk in ($familyByKey.Keys | Sort-Object)) {
        $sizes = @($familyByKey[$fk])
        $sample = $sizes | Sort-Object `
            @{Expression={ $vcpuOfSize[$_] }}, `
            @{Expression={ $_ }} | Select-Object -First 1
        $out.Add([pscustomobject]@{
            Kind             = 'Sku'
            Raw              = '(all-skus)'
            Family           = $familyRaw[$fk]
            FamilyKey        = $fk
            SampleSize       = $sample
            SampleSizeSource = 'catalog'
            Valid            = $true
            Reason           = ("catalog-enumerated; smallest of {0} sibling(s)" -f $sizes.Count)
        }) | Out-Null
    }
    return ,$out.ToArray()
}

