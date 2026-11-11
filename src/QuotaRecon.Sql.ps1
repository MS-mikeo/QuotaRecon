# QuotaRecon.Sql.ps1
#
# Azure SQL quota + regional availability (OPT-IN via -IncludeSqlQuota).
# Covers both Azure SQL Database and Azure SQL Managed Instance. Two native
# sub/region-scoped calls, no per-resource discovery:
#
#   GET .../subscriptions/{sub}/providers/Microsoft.Sql/locations/{region}
#           /usages?api-version=2023-08-01          (quota -> SqlQuota)
#   GET .../subscriptions/{sub}/providers/Microsoft.Sql/locations/{region}
#           /capabilities?api-version=2023-08-01    (availability ->
#                                                    SqlDbCapabilities +
#                                                    SqlMiCapabilities)
#
# Microsoft.Sql is NOT onboarded to the generic Microsoft.Quota provider
# (confirmed live: .../providers/Microsoft.Quota/quotas returns 400), so the
# native endpoint is the only option.
#
# The usages row shape differs from Microsoft.Compute/usages - `name` is a
# plain string and the label lives in `properties.displayName` - so it can't
# reuse Convert-QRUsageRow. Free-tier / free-offer counter rows are dropped.
#
# Capabilities: a service level objective (SLO) or MI vCore option that
# can't be provisioned in the region for this subscription carries a
# non-empty `reason` (e.g. "Provisioning is restricted in this region...")
# while its `status` still reads "Visible". Presence of `reason` is
# therefore the restriction signal. Rows are rolled up to
# Edition x Family x ComputeModel (DB) / Edition x Family (MI) with a
# three-way Restrictions value: None (green), Partial (amber), Yes (red).
# The internal `System` edition is always restricted and is excluded.

$script:QRSqlApiVersion       = '2023-08-01'
$script:QRSqlUsageCache       = @{}   # "$sub|$region" -> usages[]
$script:QRSqlCapabilityCache  = @{}   # "$sub|$region" -> capabilities object

function Get-QRSqlCapabilities {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SubscriptionId,
        [Parameter(Mandatory)][string] $Region
    )
    $key = "$SubscriptionId|$Region"
    if ($script:QRSqlCapabilityCache.ContainsKey($key)) {
        return $script:QRSqlCapabilityCache[$key]
    }
    $path = "/subscriptions/$SubscriptionId/providers/Microsoft.Sql" +
            "/locations/$Region/capabilities?api-version=$($script:QRSqlApiVersion)"
    $caps = Invoke-QRArm -Path $path
    $script:QRSqlCapabilityCache[$key] = $caps
    return $caps
}

function Get-QRSqlRestrictionStatus {
    param([int] $Offered, [int] $Restricted)
    if ($Offered -le 0)           { return 'N/A' }
    if ($Restricted -le 0)        { return 'None' }
    if ($Restricted -ge $Offered) { return 'Yes' }
    return 'Partial'
}

function Get-QRSqlTriState {
    param([int] $Total, [int] $Count)
    if ($Total -le 0 -or $Count -le 0) { return 'No' }
    if ($Count -ge $Total)             { return 'Yes' }
    return 'Partial'
}

function Get-QRSqlQuota {
    <#
    .SYNOPSIS
        Returns Azure SQL quota rows (SQL Database + Managed Instance) for
        one (Sub, Region), with free-tier counters removed.
    .OUTPUTS
        pscustomobject rows: SubscriptionId, SubscriptionName, Region, Scope,
        Name, DisplayName, Unit, CurrentValue, Limit, UsagePercent, Note.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SubscriptionId,
        [string] $SubscriptionName,
        [Parameter(Mandatory)][string] $Region
    )
    $key = "$SubscriptionId|$Region"
    if (-not $script:QRSqlUsageCache.ContainsKey($key)) {
        $path = "/subscriptions/$SubscriptionId/providers/Microsoft.Sql" +
                "/locations/$Region/usages?api-version=$($script:QRSqlApiVersion)"
        $script:QRSqlUsageCache[$key] = Invoke-QRArmPaged -Path $path
    }
    $usages = $script:QRSqlUsageCache[$key]

    $rows = foreach ($u in $usages) {
        $name = [string]$u.name
        if ($name -match 'Free') { continue }
        $p = $u.properties

        $scope = if ($name -match 'ManagedInstance') { 'ManagedInstance' } else { 'Database' }

        # MI rows echo the quota name back as the unit, and the regional
        # SQL DB/DW vCore row reports "Count"; normalize both.
        $unit = if ($p) { [string]$p.unit } else { '' }
        if ($name -match 'VCore') {
            $unit = 'VCores'
        } elseif ($unit -eq $name -or [string]::IsNullOrEmpty($unit)) {
            $unit = 'Count'
        }

        $current = if ($p -and $null -ne $p.currentValue) { [double]$p.currentValue } else { $null }
        $limit   = if ($p -and $null -ne $p.limit)        { [double]$p.limit }        else { $null }
        $note    = ''
        if ($null -ne $limit -and $limit -lt 0) {
            $note  = 'Azure reported -1: no explicit limit is set for this subscription/region.'
            $limit = $null
        }
        $pct = if ($limit -gt 0 -and $null -ne $current) {
            [math]::Round(($current / $limit) * 100, 2)
        } else { $null }

        [pscustomobject]@{
            SubscriptionId   = $SubscriptionId
            SubscriptionName = $SubscriptionName
            Region           = $Region
            Scope            = $scope
            Name             = $name
            DisplayName      = if ($p -and $p.displayName) { $p.displayName } else { $name }
            Unit             = $unit
            CurrentValue     = $current
            Limit            = $limit
            UsagePercent     = $pct
            Note             = $note
        }
    }
    return ,@($rows)
}

function Get-QRSqlDbCapabilityRows {
    <#
    .SYNOPSIS
        Rolls up Azure SQL Database capabilities for one (Sub, Region) to
        one row per Edition x Family x ComputeModel.
    .OUTPUTS
        pscustomobject rows: SubscriptionId, SubscriptionName, Region,
        Edition, Family, ComputeModel, SizeRange, Offered, RestrictedCount,
        Restrictions (None/Partial/Yes), ZoneRedundantCapable
        (Yes/Partial/No), Reason.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SubscriptionId,
        [string] $SubscriptionName,
        [Parameter(Mandatory)][string] $Region
    )
    $caps = Get-QRSqlCapabilities -SubscriptionId $SubscriptionId -Region $Region

    $groups = [ordered]@{}
    foreach ($sv in @($caps.supportedServerVersions)) {
        foreach ($ed in @($sv.supportedEditions)) {
            if (-not $ed -or $ed.name -eq 'System') { continue }
            foreach ($slo in @($ed.supportedServiceLevelObjectives)) {
                if (-not $slo) { continue }
                $family = if ($slo.sku -and $slo.sku.family) { [string]$slo.sku.family } else { '' }
                $model  = if ($slo.computeModel) { [string]$slo.computeModel } else { 'Provisioned' }
                $gKey   = "$($ed.name)|$family|$model"
                if (-not $groups.Contains($gKey)) {
                    $groups[$gKey] = [pscustomobject]@{
                        Edition = [string]$ed.name; Family = $family; Model = $model
                        EditionReason = [string]$ed.reason
                        Items = New-Object System.Collections.Generic.List[object]
                    }
                }
                $groups[$gKey].Items.Add($slo)
            }
        }
    }

    $rows = foreach ($g in $groups.Values) {
        $items      = $g.Items
        $offered    = $items.Count
        $restricted = @($items | Where-Object { $_.reason -or $g.EditionReason })
        $zr         = @($items | Where-Object { $_.zoneRedundant }).Count

        $sizes = @($items | Where-Object { $_.performanceLevel } | ForEach-Object { [double]$_.performanceLevel.value })
        $unit  = @($items | Where-Object { $_.performanceLevel } | Select-Object -First 1 | ForEach-Object { $_.performanceLevel.unit })
        $range = if ($sizes.Count) {
            $m = $sizes | Measure-Object -Minimum -Maximum
            "{0}-{1} {2}" -f $m.Minimum, $m.Maximum, ($unit -join '')
        } else { '' }

        $reasons = @($restricted | ForEach-Object { if ($_.reason) { [string]$_.reason } else { $g.EditionReason } } |
                     Where-Object { $_ } | Sort-Object -Unique)

        [pscustomobject]@{
            SubscriptionId       = $SubscriptionId
            SubscriptionName     = $SubscriptionName
            Region               = $Region
            Edition              = $g.Edition
            Family               = $g.Family
            ComputeModel         = $g.Model
            SizeRange            = $range
            Offered              = $offered
            RestrictedCount      = $restricted.Count
            Restrictions         = Get-QRSqlRestrictionStatus -Offered $offered -Restricted $restricted.Count
            ZoneRedundantCapable = Get-QRSqlTriState -Total $offered -Count $zr
            Reason               = $reasons -join ' | '
        }
    }
    return ,@($rows)
}

function Get-QRSqlMiCapabilityRows {
    <#
    .SYNOPSIS
        Rolls up Azure SQL Managed Instance capabilities for one
        (Sub, Region) to one row per Edition (+ GPv2 flag) x Family.
    .OUTPUTS
        pscustomobject rows: SubscriptionId, SubscriptionName, Region,
        Edition, GeneralPurposeV2, Family, Status, vCoreRange, Offered,
        RestrictedCount, Restrictions (None/Partial/Yes),
        ZoneRedundantCapable (Yes/No), Reason.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SubscriptionId,
        [string] $SubscriptionName,
        [Parameter(Mandatory)][string] $Region
    )
    $caps = Get-QRSqlCapabilities -SubscriptionId $SubscriptionId -Region $Region

    $rows = foreach ($mv in @($caps.supportedManagedInstanceVersions)) {
        foreach ($ed in @($mv.supportedEditions)) {
            if (-not $ed) { continue }
            foreach ($fam in @($ed.supportedFamilies)) {
                if (-not $fam) { continue }
                $vcores     = @($fam.supportedVcoresValues | Where-Object { $_ })
                $parentWhy  = @($mv.reason, $ed.reason, $fam.reason | Where-Object { $_ })
                $restricted = if ($parentWhy.Count) { $vcores } else { @($vcores | Where-Object { $_.reason }) }

                $range = ''
                if ($vcores.Count) {
                    $m = $vcores | ForEach-Object { [double]$_.value } | Measure-Object -Minimum -Maximum
                    $range = "{0}-{1}" -f $m.Minimum, $m.Maximum
                }
                $reasons = @(@($parentWhy) + @($restricted | ForEach-Object { $_.reason }) |
                             Where-Object { $_ } | ForEach-Object { [string]$_ } | Sort-Object -Unique)

                [pscustomobject]@{
                    SubscriptionId       = $SubscriptionId
                    SubscriptionName     = $SubscriptionName
                    Region               = $Region
                    Edition              = [string]$ed.name
                    GeneralPurposeV2     = if ($ed.isGeneralPurposeV2) { 'Yes' } else { 'No' }
                    Family               = [string]$fam.name
                    Status               = [string]$fam.status
                    vCoreRange           = $range
                    Offered              = $vcores.Count
                    RestrictedCount      = @($restricted).Count
                    Restrictions         = Get-QRSqlRestrictionStatus -Offered $vcores.Count -Restricted @($restricted).Count
                    ZoneRedundantCapable = if ($fam.zoneRedundant) { 'Yes' } else { 'No' }
                    Reason               = $reasons -join ' | '
                }
            }
        }
    }
    return ,@($rows)
}
