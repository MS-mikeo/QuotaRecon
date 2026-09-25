# QuotaRecon.Inventory.ps1
#
# Deployed VM & VMSS enumeration via Azure Resource Graph. One POST fans out
# across every subscription in the query, filtered to the target regions and
# to the VM sizes that belong to the families the user asked about.
#
# The Resource Graph "resources" table exposes:
#   - properties.hardwareProfile.vmSize     (for VMs)
#   - sku.name / sku.capacity                (for VMSS)
#   - properties.extended.instanceView.powerState.code  (for VMs; may be blank
#     on new subs or non-primary regions where instanceView isn't collected)
#
# The endpoint is /providers/Microsoft.ResourceGraph/resources at
# api-version=2022-10-01. Paginated with $skipToken past 1000 rows.

$script:QRResourceGraphApiVersion = '2022-10-01'

function Invoke-QRResourceGraph {
    <#
    .SYNOPSIS
        POSTs a KQL query to Azure Resource Graph and returns every row,
        following $skipToken pagination past the 1000-row page size.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]   $Query,
        [Parameter(Mandatory)][string[]] $SubscriptionIds
    )
    $all = New-Object System.Collections.Generic.List[object]
    $skipToken = $null
    do {
        $options = [ordered]@{
            '$top'         = 1000
            resultFormat   = 'objectArray'
        }
        if ($skipToken) { $options['$skipToken'] = $skipToken }
        $body = @{
            subscriptions = $SubscriptionIds
            query         = $Query
            options       = $options
        }
        $resp = Invoke-QRArm `
            -Path   "/providers/Microsoft.ResourceGraph/resources?api-version=$($script:QRResourceGraphApiVersion)" `
            -Method POST `
            -Body   $body
        if ($resp.data) {
            foreach ($row in @($resp.data)) { $all.Add($row) | Out-Null }
        }
        $skipToken = $resp.'$skipToken'
    } while ($skipToken)
    return ,$all.ToArray()
}

function Get-QRFamilySizeMap {
    <#
    .SYNOPSIS
        Builds a hashtable of familyKey -> [size names] by walking the union
        of Microsoft.Compute/skus catalogs for the target regions. This is
        the same key shape ConvertTo-QRSkuLookupKey / Get-QRFamilyLookupKey
        produce, so it's directly comparable to user input.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]   $ProbeSubscriptionId,
        [Parameter(Mandatory)][string[]] $Regions,
        [Parameter(Mandatory)][string[]] $FamilyKeys
    )
    $wantedKeys = @{}
    foreach ($k in $FamilyKeys) { $wantedKeys[$k.ToLowerInvariant()] = $true }

    $map = @{}
    foreach ($region in $Regions) {
        try {
            $catalog = Get-QRComputeSkuCatalog -SubscriptionId $ProbeSubscriptionId -Region $region
        } catch {
            continue
        }
        foreach ($sku in $catalog) {
            if ($sku.resourceType -ne 'virtualMachines' -or -not $sku.family) { continue }
            $fk = Get-QRFamilyLookupKey -Family $sku.family
            if (-not $wantedKeys.ContainsKey($fk)) { continue }
            if (-not $map.ContainsKey($fk)) {
                $map[$fk] = New-Object System.Collections.Generic.List[string]
            }
            if ($map[$fk] -notcontains $sku.name) { $map[$fk].Add($sku.name) | Out-Null }
        }
    }
    return $map
}

function Get-QRVmInventory {
    <#
    .SYNOPSIS
        Returns one row per deployed VM (and one row per VMSS) whose vmSize
        belongs to a requested family and whose location is in the target
        region list.

    .OUTPUTS
        pscustomobject rows with:
            SubscriptionId, SubscriptionName, ResourceGroup, VmName, Region,
            VmSize, Family, FamilyKey, Zone ('1'/'2'/'3'/'Regional'),
            PowerState, Deployment ('VM'|'VMSS'), InstanceCount, ResourceId

    .PARAMETER FamilySizeMap
        familyKey -> [size names] (as built by Get-QRFamilySizeMap).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]  $SubscriptionIds,
        [Parameter(Mandatory)][string[]]  $Regions,
        [Parameter(Mandatory)][hashtable] $FamilySizeMap,
        [Parameter(Mandatory)][hashtable] $SubscriptionNames
    )

    # Reverse lookup: sizeName (lower) -> familyKey
    $sizeToFamilyKey = @{}
    foreach ($fk in $FamilySizeMap.Keys) {
        foreach ($size in $FamilySizeMap[$fk]) {
            $sizeToFamilyKey[$size.ToLowerInvariant()] = $fk
        }
    }
    if ($sizeToFamilyKey.Count -eq 0) { return @() }

    $sizeList = @($sizeToFamilyKey.Keys | ForEach-Object { "'$_'" }) -join ','
    $regionList = @($Regions | ForEach-Object { "'$_'" }) -join ','

    # One KQL that unions VMs and VMSS resources, filters to target sizes/
    # regions, and projects a common column set. tolower() around vmSize so
    # the `in~` casing doesn't miss.
    $query = @"
resources
| where (type =~ 'microsoft.compute/virtualmachines') or (type =~ 'microsoft.compute/virtualmachinescalesets')
| extend Deployment = iff(type =~ 'microsoft.compute/virtualmachines', 'VM', 'VMSS')
| extend VmSize = iff(type =~ 'microsoft.compute/virtualmachines',
                     tostring(properties.hardwareProfile.vmSize),
                     tostring(sku.name))
| extend InstanceCount = iff(type =~ 'microsoft.compute/virtualmachines', 1, toint(sku.capacity))
| extend PowerState = iff(type =~ 'microsoft.compute/virtualmachines',
                         tostring(properties.extended.instanceView.powerState.code),
                         '(scale-set)')
| where tolower(VmSize) in~ ($sizeList)
| where location in~ ($regionList)
| project SubscriptionId=subscriptionId,
          ResourceGroup=resourceGroup,
          VmName=name,
          Region=location,
          VmSize,
          Zones=zones,
          Deployment,
          InstanceCount,
          PowerState,
          ResourceId=id
"@

    $rows = @()
    try {
        $rows = Invoke-QRResourceGraph -Query $query -SubscriptionIds $SubscriptionIds
    } catch {
        Write-Warning "Resource Graph query failed: $($_.Exception.Message)"
        return @()
    }

    $out = foreach ($r in $rows) {
        $sizeLower = ([string]$r.VmSize).ToLowerInvariant()
        $fk = $sizeToFamilyKey[$sizeLower]
        if (-not $fk) { continue }
        $zone =
            if ($r.Zones -is [System.Array] -and $r.Zones.Count -gt 0) { [string]$r.Zones[0] }
            elseif ($r.Zones)                                          { [string]$r.Zones }
            else                                                        { 'Regional' }
        $subName = $SubscriptionNames[$r.SubscriptionId]
        if (-not $subName) { $subName = '(name unknown)' }
        $power   =
            if ([string]::IsNullOrWhiteSpace($r.PowerState)) { '(unknown)' }
            else {
                # 'PowerState/running' -> 'running'
                (([string]$r.PowerState) -replace '^PowerState/','')
            }
        [pscustomobject]@{
            SubscriptionId   = $r.SubscriptionId
            SubscriptionName = $subName
            ResourceGroup    = $r.ResourceGroup
            VmName           = $r.VmName
            Region           = $r.Region
            VmSize           = $r.VmSize
            Family           = ''      # optional display; can be filled from FamilySizeMap keys if needed
            FamilyKey        = $fk
            Zone             = $zone
            PowerState       = $power
            Deployment       = $r.Deployment
            InstanceCount    = if ($r.InstanceCount) { [int]$r.InstanceCount } else { 1 }
            ResourceId       = $r.ResourceId
        }
    }
    return ,@($out)
}

function Build-QRVmInventoryCountIndex {
    <#
    .SYNOPSIS
        Builds a hashtable keyed 'subId|region|familyKey' -> total deployed
        instance count, for the Summary sheet to join against.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Inventory)

    $idx = @{}
    foreach ($row in $Inventory) {
        if (-not $row.SubscriptionId -or -not $row.Region -or -not $row.FamilyKey) { continue }
        $key = ('{0}|{1}|{2}' -f $row.SubscriptionId, $row.Region.ToLowerInvariant(), $row.FamilyKey.ToLowerInvariant())
        if (-not $idx.ContainsKey($key)) { $idx[$key] = 0 }
        $idx[$key] += [int]$row.InstanceCount
    }
    return $idx
}
