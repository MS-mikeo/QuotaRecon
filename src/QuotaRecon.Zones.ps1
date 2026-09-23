# QuotaRecon.Zones.ps1
#
# Logical -> physical availability-zone map. Uses the per-subscription
# /locations endpoint (api-version 2022-12-01+), whose `availabilityZoneMappings`
# array carries both `logicalZone` ("1"/"2"/"3") and `physicalZone`
# (e.g. "uksouth-az1"). This is the same data source shown in the Microsoft
# Learn page "What are Azure Availability Zones?" (Azure PowerShell tab):
#     GET /subscriptions/{id}/locations?api-version=2022-12-01
#
# No peer subscription is needed. Two subscriptions in the same tenant that
# report the same `physicalZone` string for their logical zone 1 are actually
# co-located.

$script:QRLocationDetailsCache = @{}   # subId -> locations[]

function Get-QRLocationDetails {
    <#
    .SYNOPSIS
        Returns the raw `value` array from
        GET /subscriptions/{id}/locations?api-version=2022-12-01
        cached per subscription.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string] $SubscriptionId)

    if ($script:QRLocationDetailsCache.ContainsKey($SubscriptionId)) {
        return $script:QRLocationDetailsCache[$SubscriptionId]
    }
    $resp = Invoke-QRArm -Path "/subscriptions/$SubscriptionId/locations?api-version=2022-12-01"
    $script:QRLocationDetailsCache[$SubscriptionId] = $resp.value
    return $resp.value
}

function Get-QRZoneMap {
    <#
    .SYNOPSIS
        For a single (subscription, region), returns one row per logical
        zone with the physical-zone label the platform maps it to.
    .OUTPUTS
        pscustomobject rows:
            SubscriptionId, SubscriptionName, Region,
            LogicalZone, PhysicalZone, Note
        If the region has no `availabilityZoneMappings` (non-AZ region), one
        row is emitted with LogicalZone/PhysicalZone empty and a Note.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SubscriptionId,
        [string] $SubscriptionName,
        [Parameter(Mandatory)][string] $Region
    )
    $locs = Get-QRLocationDetails -SubscriptionId $SubscriptionId
    $loc = $locs | Where-Object { $_.name -eq $Region } | Select-Object -First 1
    if (-not $loc) {
        return ,@([pscustomobject]@{
            SubscriptionId   = $SubscriptionId
            SubscriptionName = $SubscriptionName
            Region           = $Region
            LogicalZone      = ''
            PhysicalZone     = ''
            Note             = 'Region not visible to this subscription.'
        })
    }
    if (-not $loc.availabilityZoneMappings) {
        return ,@([pscustomobject]@{
            SubscriptionId   = $SubscriptionId
            SubscriptionName = $SubscriptionName
            Region           = $Region
            LogicalZone      = ''
            PhysicalZone     = ''
            Note             = 'Region does not expose availability zones.'
        })
    }
    $out = foreach ($m in @($loc.availabilityZoneMappings)) {
        [pscustomobject]@{
            SubscriptionId   = $SubscriptionId
            SubscriptionName = $SubscriptionName
            Region           = $Region
            LogicalZone      = $m.logicalZone
            PhysicalZone     = $m.physicalZone
            Note             = ''
        }
    }
    return ,$out
}

function Get-QRZoneMapAllSubs {
    <#
    .SYNOPSIS
        Convenience wrapper: emits Get-QRZoneMap rows for every subscription
        and region. Errors on a single sub/region become one row with the
        message in Note; they never abort the run.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]  $SubscriptionIds,
        [Parameter(Mandatory)][hashtable] $SubscriptionNames,
        [Parameter(Mandatory)][string[]]  $Regions
    )
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($sub in $SubscriptionIds) {
        $name = $SubscriptionNames[$sub]
        foreach ($region in $Regions) {
            try {
                foreach ($row in (Get-QRZoneMap -SubscriptionId $sub -SubscriptionName $name -Region $region)) {
                    $rows.Add($row)
                }
            }
            catch {
                $rows.Add([pscustomobject]@{
                    SubscriptionId   = $sub
                    SubscriptionName = $name
                    Region           = $region
                    LogicalZone      = ''
                    PhysicalZone     = ''
                    Note             = "ERROR: $($_.Exception.Message)"
                })
            }
        }
    }
    return ,$rows.ToArray()
}

function Get-QRPhysicalZonesForRegion {
    <#
    .SYNOPSIS
        Returns a hashtable subscriptionId -> (hashtable logicalZone -> physicalZone)
        for a given region, using the collected zone-map rows. Used by the
        Summary sheet to attach a "PhysicalZones" string per row.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $ZoneMapRows,
        [Parameter(Mandatory)][string] $Region
    )
    $result = @{}
    foreach ($row in $ZoneMapRows) {
        if ($row.Region -ne $Region -or -not $row.LogicalZone) { continue }
        if (-not $result.ContainsKey($row.SubscriptionId)) { $result[$row.SubscriptionId] = @{} }
        $result[$row.SubscriptionId][$row.LogicalZone] = $row.PhysicalZone
    }
    return $result
}
