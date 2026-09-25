# QuotaRecon.QuotaGroups.ps1
#
# Group Quotas ("quota pools") live at the Management Group scope, not the
# subscription scope. The Microsoft.Quota provider surfaces them via a small
# set of paged endpoints that we know work (verified against Microsoft.Quota
# 2023-06-01-preview):
#
#   GET  .../managementGroups/{mg}/providers/Microsoft.Quota/groupQuotas
#            ?api-version=2023-06-01-preview
#   GET  .../groupQuotas/{groupName}
#   GET  .../groupQuotas/{groupName}/subscriptions
#   GET  .../groupQuotas/{groupName}/resourceProviders/Microsoft.Compute
#            /groupQuotaLimits?api-version=...&$filter=location eq '{region}'
#
# The limits endpoint REQUIRES the location filter — an unfiltered call
# returns 400 Bad Request.

$script:QRGroupQuotaApiVersion = '2023-06-01-preview'
$script:QRGroupQuotaCache      = @{}   # "$mg" -> group[]

function Get-QRAllManagementGroups {
    <#
    .SYNOPSIS
        Returns every Management Group ID visible to the current identity,
        via GET /providers/Microsoft.Management/managementGroups. Follows
        pagination. Falls back to an empty list on failure (rather than
        aborting the run) so a caller without MG-Reader still gets a
        workbook.
    #>
    [CmdletBinding()]
    param()
    try {
        $items = Invoke-QRArmPaged -Path "/providers/Microsoft.Management/managementGroups?api-version=2020-05-01"
        return @($items | ForEach-Object { $_.name } | Where-Object { $_ } | Select-Object -Unique)
    } catch {
        Write-Warning "Get-QRAllManagementGroups failed: $($_.Exception.Message)"
        return @()
    }
}

function Get-QRGroupQuotas {
    <#
    .SYNOPSIS
        Returns every group quota resource at the given Management Group,
        following nextLink pagination.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $ManagementGroupId
    )
    if ($script:QRGroupQuotaCache.ContainsKey($ManagementGroupId)) {
        return $script:QRGroupQuotaCache[$ManagementGroupId]
    }
    $path  = "/providers/Microsoft.Management/managementGroups/$ManagementGroupId" +
             "/providers/Microsoft.Quota/groupQuotas" +
             "?api-version=$($script:QRGroupQuotaApiVersion)"
    $items = Invoke-QRArmPaged -Path $path
    $script:QRGroupQuotaCache[$ManagementGroupId] = $items
    return $items
}

function Get-QRGroupQuotaSubscriptions {
    <#
    .SYNOPSIS
        Returns the list of member subscription IDs for a group.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $ManagementGroupId,
        [Parameter(Mandatory)][string] $GroupName
    )
    $path = "/providers/Microsoft.Management/managementGroups/$ManagementGroupId" +
            "/providers/Microsoft.Quota/groupQuotas/$GroupName/subscriptions" +
            "?api-version=$($script:QRGroupQuotaApiVersion)"
    $rows = Invoke-QRArmPaged -Path $path
    # Normalize: the response uses `Properties` with a capital P for this
    # sub-resource. Return a plain [pscustomobject] with SubscriptionId.
    return ,@($rows | ForEach-Object {
        $subId = $null
        if ($_.Properties -and $_.Properties.subscriptionId) { $subId = $_.Properties.subscriptionId }
        elseif ($_.properties -and $_.properties.subscriptionId) { $subId = $_.properties.subscriptionId }
        elseif ($_.name) { $subId = $_.name }
        [pscustomobject]@{
            SubscriptionId    = $subId
            ProvisioningState =
                if ($_.Properties)   { $_.Properties.provisioningState }
                elseif ($_.properties) { $_.properties.provisioningState }
                else { '' }
        }
    })
}

function Get-QRGroupQuotaLimits {
    <#
    .SYNOPSIS
        Returns the compute quota limits for one group in one region.
        Empty array when no limits are set in that region for that group.
    .OUTPUTS
        pscustomobject rows:
            FamilyKey, FamilyDisplay, Region, Limit, AvailableLimit,
            Unit, Allocations (array of { SubscriptionId, QuotaAllocated })
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $ManagementGroupId,
        [Parameter(Mandatory)][string] $GroupName,
        [Parameter(Mandatory)][string] $Region,
        [string] $ResourceProvider = 'Microsoft.Compute'
    )
    $filter = "location eq '$Region'"
    $path = "/providers/Microsoft.Management/managementGroups/$ManagementGroupId" +
            "/providers/Microsoft.Quota/groupQuotas/$GroupName" +
            "/resourceProviders/$ResourceProvider/groupQuotaLimits" +
            "?api-version=$($script:QRGroupQuotaApiVersion)&`$filter=$filter"
    try {
        $rows = Invoke-QRArmPaged -Path $path
    } catch {
        # A 404 on this endpoint means "no limits configured" in most cases.
        # Swallow so the workbook shows an empty limits list rather than
        # failing the whole run. Other errors should propagate.
        if ($_.Exception.Message -match '404') { return @() }
        throw
    }
    $out = foreach ($row in $rows) {
        $p = $row.properties
        $allocs = @()
        if ($p.allocatedToSubscriptions -and $p.allocatedToSubscriptions.value) {
            foreach ($a in @($p.allocatedToSubscriptions.value)) {
                $allocs += [pscustomobject]@{
                    SubscriptionId  = $a.subscriptionId
                    QuotaAllocated  = $a.quotaAllocated
                }
            }
        }
        [pscustomobject]@{
            FamilyKey       = $p.name.value
            FamilyDisplay   = $p.name.localizedValue
            Region          = $p.region
            Limit           = $p.limit
            AvailableLimit  = $p.availableLimit
            Unit            = $p.unit
            Allocations     = $allocs
        }
    }
    return ,$out
}

function Get-QRAllGroupQuotaData {
    <#
    .SYNOPSIS
        Top-level roll-up: given a set of Management Groups and a region list,
        returns three flat collections ready to be written to worksheets:

            .Groups       - one row per (Mg × Group) that exists
            .Limits       - one row per (Mg × Group × Region × Family)
            .Allocations  - one row per (Mg × Group × Region × Family × Sub)

        Subscription-name fields are emitted as empty strings; the caller is
        expected to enrich them AFTER running Get-QRSubscriptionNames on the
        union of input subs + every sub ID referenced by any group. This
        two-phase pattern lets us resolve names for subs the user never
        explicitly asked about (e.g. group members outside the input set).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]] $ManagementGroupIds,
        [Parameter(Mandatory)][string[]] $Regions
    )

    $groups      = New-Object System.Collections.Generic.List[object]
    $limits      = New-Object System.Collections.Generic.List[object]
    $allocations = New-Object System.Collections.Generic.List[object]

    foreach ($mg in $ManagementGroupIds) {
        Write-Verbose "[QuotaGroups] MG $mg"
        $mgGroups = @()
        try {
            $mgGroups = Get-QRGroupQuotas -ManagementGroupId $mg
        } catch {
            Write-Verbose ("MG '{0}' groupQuotas list failed (probably 404 = no groups): {1}" -f $mg, $_.Exception.Message)
            continue
        }
        foreach ($grp in $mgGroups) {
            $grpName    = $grp.name
            $grpDisplay = $grp.properties.displayName
            $grpType    = $grp.properties.groupType
            $addlAttrs  = $grp.properties.additionalAttributes

            # Members
            $members = @()
            try {
                $members = Get-QRGroupQuotaSubscriptions -ManagementGroupId $mg -GroupName $grpName
            } catch {
                Write-Warning "MG '$mg' group '$grpName' members list failed: $($_.Exception.Message)"
            }
            $memberCount = @($members).Count
            $memberIds   = @($members | ForEach-Object { $_.SubscriptionId } | Where-Object { $_ } | Select-Object -Unique)

            $groups.Add([pscustomobject]@{
                MgName            = $mg
                GroupName         = $grpName
                GroupDisplayName  = $grpDisplay
                GroupType         = $grpType
                GroupingIdType    = $addlAttrs.groupId.groupingIdType
                GroupingIdValue   = $addlAttrs.groupId.value
                MemberSubCount    = $memberCount
                MemberSubIds      = ($memberIds -join ',')
                MemberSubNames    = ''    # filled in later by main script
            })

            # Limits per region
            foreach ($region in $Regions) {
                $regLimits = @()
                try {
                    $regLimits = Get-QRGroupQuotaLimits -ManagementGroupId $mg -GroupName $grpName -Region $region
                } catch {
                    Write-Warning "MG '$mg' group '$grpName' limits ($region) failed: $($_.Exception.Message)"
                    continue
                }
                foreach ($lim in $regLimits) {
                    $limits.Add([pscustomobject]@{
                        MgName          = $mg
                        GroupName       = $grpName
                        GroupDisplayName= $grpDisplay
                        Region          = $lim.Region
                        FamilyKey       = $lim.FamilyKey
                        FamilyDisplay   = $lim.FamilyDisplay
                        PoolLimit       = $lim.Limit
                        AvailableLimit  = $lim.AvailableLimit
                        Unit            = $lim.Unit
                        AllocationCount = @($lim.Allocations).Count
                    })

                    # Per-sub allocations
                    foreach ($alloc in @($lim.Allocations)) {
                        $direction = if ($alloc.QuotaAllocated -lt 0) { 'Contributed' }
                                     elseif ($alloc.QuotaAllocated -gt 0) { 'Drawn' }
                                     else { 'None' }
                        $allocations.Add([pscustomobject]@{
                            MgName            = $mg
                            GroupName         = $grpName
                            GroupDisplayName  = $grpDisplay
                            Region            = $lim.Region
                            FamilyKey         = $lim.FamilyKey
                            SubscriptionId    = $alloc.SubscriptionId
                            SubscriptionName  = ''      # filled in later by main script
                            QuotaAllocated    = $alloc.QuotaAllocated
                            Direction         = $direction
                            PoolLimit         = $lim.Limit
                            PoolAvailableLimit= $lim.AvailableLimit
                        })
                    }
                }
            }
        }
    }

    return @{
        Groups      = $groups.ToArray()
        Limits      = $limits.ToArray()
        Allocations = $allocations.ToArray()
    }
}

function Get-QRSubIdsInGroupData {
    <#
    .SYNOPSIS
        Extracts every unique subscription ID referenced by a QuotaGroup data
        payload (from Get-QRAllGroupQuotaData). Includes member subs AND
        allocation subs.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] $QuotaGroupData)

    $set = New-Object System.Collections.Generic.HashSet[string]
    foreach ($g in $QuotaGroupData.Groups) {
        if ($g.MemberSubIds) {
            foreach ($id in ($g.MemberSubIds -split ',')) {
                $t = $id.Trim()
                if ($t) { [void]$set.Add($t) }
            }
        }
    }
    foreach ($a in $QuotaGroupData.Allocations) {
        if ($a.SubscriptionId) { [void]$set.Add($a.SubscriptionId) }
    }
    return ,@($set)
}

function Set-QRSubscriptionNamesOnGroupData {
    <#
    .SYNOPSIS
        Given a QuotaGroup data payload and a fully-resolved subscription
        names hashtable, mutates the .Groups and .Allocations rows to fill in
        SubscriptionName / MemberSubNames.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $QuotaGroupData,
        [Parameter(Mandatory)][hashtable] $SubscriptionNames
    )

    foreach ($g in $QuotaGroupData.Groups) {
        if ($g.MemberSubIds) {
            $names = foreach ($id in ($g.MemberSubIds -split ',')) {
                $t = $id.Trim()
                if (-not $t) { continue }
                if ($SubscriptionNames.ContainsKey($t)) { $SubscriptionNames[$t] } else { '(name unknown)' }
            }
            $g.MemberSubNames = (@($names) -join ',')
        }
    }
    foreach ($a in $QuotaGroupData.Allocations) {
        if ($a.SubscriptionId -and $SubscriptionNames.ContainsKey($a.SubscriptionId)) {
            $a.SubscriptionName = $SubscriptionNames[$a.SubscriptionId]
        } elseif (-not $a.SubscriptionName) {
            $a.SubscriptionName = '(name unknown)'
        }
    }
}

function Build-QRQuotaGroupSummaryIndex {
    <#
    .SYNOPSIS
        Builds a hashtable keyed by "subId|region|familyKey" -> aggregated
        group participation record, for the Summary sheet to join against.

    .OUTPUTS
        hashtable value shape:
            @{
                InQuotaGroup = 'Yes'
                GroupNames   = 'GroupA;GroupB'
                PoolLimit    = 800
                PoolAvailable= 1200
                SubAllocated = -100
                SubDirection = 'Contributed'
            }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Allocations)

    $idx = @{}
    foreach ($a in $Allocations) {
        if (-not $a.SubscriptionId -or -not $a.Region -or -not $a.FamilyKey) { continue }
        $key = ('{0}|{1}|{2}' -f $a.SubscriptionId, $a.Region.ToLowerInvariant(), $a.FamilyKey.ToLowerInvariant())
        if (-not $idx.ContainsKey($key)) {
            $idx[$key] = [ordered]@{
                InQuotaGroup   = 'Yes'
                GroupNames     = New-Object System.Collections.Generic.List[string]
                PoolLimit      = 0
                PoolAvailable  = 0
                SubAllocated   = 0
                SubDirection   = ''
            }
        }
        $rec = $idx[$key]
        if (-not $rec.GroupNames.Contains([string]$a.GroupDisplayName)) {
            $rec.GroupNames.Add([string]$a.GroupDisplayName) | Out-Null
        }
        $rec.PoolLimit     += [int]$a.PoolLimit
        $rec.PoolAvailable += [int]$a.PoolAvailableLimit
        $rec.SubAllocated  += [int]$a.QuotaAllocated
        $rec.SubDirection   = $a.Direction
    }
    return $idx
}
