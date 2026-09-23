# QuotaRecon.Input.ps1
#
# Read + normalize the three input surfaces: subscriptions, regions, SKUs.
# Every normalizer returns objects shaped { Raw; Normalized; Valid; Reason }
# so the main script can emit both an Inputs sheet and an Errors sheet
# without losing any user data.

function Read-QRList {
    <#
    .SYNOPSIS
        Reads a single-column CSV *or* a comma-string into a de-duped list of
        raw strings. Blank lines and lines starting with '#' are ignored.
    #>
    [CmdletBinding()]
    param(
        [string]   $CsvPath,
        [string[]] $Inline,
        [string]   $ColumnName
    )

    $values = New-Object System.Collections.Generic.List[string]

    if ($Inline) {
        foreach ($v in $Inline) {
            if ($null -eq $v) { continue }
            foreach ($p in ($v -split '[,;\r\n]')) {
                $t = ($p.Trim() -replace '^"|"$','').Trim()
                if ($t -and -not $t.StartsWith('#')) { $values.Add($t) }
            }
        }
    }
    elseif ($CsvPath) {
        if (-not (Test-Path -LiteralPath $CsvPath)) {
            throw "Input file not found: $CsvPath"
        }
        # Support both CSV-with-header and raw text.
        $lines = Get-Content -LiteralPath $CsvPath
        $header = $null
        foreach ($ln in $lines) {
            # Strip a leading BOM and any surrounding double-quotes that
            # spreadsheet tools add when a value contains commas or quotes.
            $t = $ln.Trim() -replace '^\uFEFF',''
            $t = ($t -replace '^"|"$','').Trim()
            if (-not $t -or $t.StartsWith('#')) { continue }
            if (-not $header -and $ColumnName -and
                ($t -ieq $ColumnName -or $t.StartsWith("$ColumnName,"))) {
                $header = $t
                continue
            }
            # Handle a header row that's not a match too — first non-comment
            # line that contains no digit/GUID is treated as header.
            if (-not $header -and $t -notmatch '[0-9]') {
                $header = $t
                continue
            }
            foreach ($p in ($t -split ',')) {
                $x = ($p.Trim() -replace '^"|"$','').Trim()
                if ($x -and -not $x.StartsWith('#')) { $values.Add($x) }
            }
        }
    }

    return ,($values | Select-Object -Unique)
}

# ---------- Subscriptions ------------------------------------------------

function Resolve-QRSubscriptions {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]] $Raw)

    $guid = [regex]'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$'
    $out  = foreach ($r in $Raw) {
        $t = $r.Trim().ToLowerInvariant()
        if ($guid.IsMatch($t)) {
            [pscustomobject]@{
                Kind       = 'Subscription'
                Raw        = $r
                Normalized = $t
                Valid      = $true
                Reason     = ''
            }
        }
        else {
            [pscustomobject]@{
                Kind       = 'Subscription'
                Raw        = $r
                Normalized = ''
                Valid      = $false
                Reason     = 'Not a GUID.'
            }
        }
    }
    return ,$out
}

function Get-QRSubscriptionNames {
    <#
    .SYNOPSIS
        Returns a hashtable of subscriptionId -> displayName.
    .DESCRIPTION
        Strategy (fastest first, most reliable last):
          1. One bulk ARM list: GET /subscriptions?api-version=2022-12-01
             This returns every sub the current identity can see in the
             current tenant. In practice this resolves ~99% of cases in
             one round trip.
          2. For anything not found in the bulk list, try Get-AzSubscription,
             which handles cross-tenant lookups from the local Az context
             cache (falls back gracefully when other tenants require MFA).
          3. For anything still missing, direct ARM GET /subscriptions/{id}.
          4. Anything that STILL fails is marked '(inaccessible: <reason>)'
             so you can see why in the workbook rather than a bare
             '(inaccessible)'.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]] $SubscriptionIds)

    $map = @{}
    $remaining = @($SubscriptionIds | Select-Object -Unique)

    # --- 1. Bulk list ------------------------------------------------------
    try {
        $listed = Invoke-QRArm -Path "/subscriptions?api-version=2022-12-01"
        if ($listed.value) {
            foreach ($s in $listed.value) {
                if ($remaining -contains $s.subscriptionId -and -not $map.ContainsKey($s.subscriptionId)) {
                    $map[$s.subscriptionId] = $s.displayName
                }
            }
        }
    } catch {
        Write-Verbose "Bulk /subscriptions list failed: $($_.Exception.Message)"
    }
    $remaining = @($remaining | Where-Object { -not $map.ContainsKey($_) })

    # --- 2. Get-AzSubscription (silent - Az handles cross-tenant caching) --
    if ($remaining.Count -gt 0 -and (Get-Command Get-AzSubscription -ErrorAction SilentlyContinue)) {
        foreach ($id in $remaining) {
            try {
                $s = Get-AzSubscription -SubscriptionId $id -WarningAction SilentlyContinue -ErrorAction Stop
                if ($s -and $s.Name) { $map[$id] = $s.Name }
            } catch {
                # fall through to step 3
            }
        }
        $remaining = @($remaining | Where-Object { -not $map.ContainsKey($_) })
    }

    # --- 3. Direct ARM GET per remaining sub ------------------------------
    foreach ($id in $remaining) {
        try {
            $r = Invoke-QRArm -Path "/subscriptions/$id`?api-version=2022-12-01"
            if ($r.displayName) {
                $map[$id] = $r.displayName
            } else {
                $map[$id] = '(no displayName)'
            }
        } catch {
            $reason = $_.Exception.Message
            if ($reason.Length -gt 120) { $reason = $reason.Substring(0,120) + '...' }
            $map[$id] = "(inaccessible: $reason)"
        }
    }

    return $map
}

# ---------- Regions ------------------------------------------------------

$script:QRLocationCache = $null

function Get-QRLocationCatalog {
    <#
    .SYNOPSIS
        Returns a hashtable keyed by lowercased normalized form (both ARM
        name and display name, with spaces stripped) -> ARM name.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string] $SubscriptionId)

    if ($script:QRLocationCache) { return $script:QRLocationCache }

    $resp = Invoke-QRArm -Path "/subscriptions/$SubscriptionId/locations?api-version=2022-12-01"
    $map = @{}
    foreach ($l in $resp.value) {
        $arm    = $l.name.ToLowerInvariant()
        $disp   = $l.displayName
        $map[$arm] = $arm
        $map[$disp.ToLowerInvariant()] = $arm
        $map[($disp -replace '\s','').ToLowerInvariant()] = $arm
    }
    $script:QRLocationCache = $map
    return $map
}

function Resolve-QRRegions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]] $Raw,
        [Parameter(Mandatory)][string]   $ProbeSubscriptionId
    )

    $catalog = Get-QRLocationCatalog -SubscriptionId $ProbeSubscriptionId
    $keys    = $catalog.Keys

    $out = foreach ($r in $Raw) {
        $key = ($r -replace '\s','').ToLowerInvariant()
        if ($catalog.ContainsKey($key)) {
            [pscustomobject]@{
                Kind       = 'Region'
                Raw        = $r
                Normalized = $catalog[$key]
                Valid      = $true
                Reason     = ''
            }
        }
        else {
            # Fuzzy suggest: keys that share a prefix or contain the input.
            $suggest = $keys | Where-Object { $_ -like "*$key*" -or $key -like "*$_*" } | Select-Object -First 3
            [pscustomobject]@{
                Kind       = 'Region'
                Raw        = $r
                Normalized = ''
                Valid      = $false
                Reason     = if ($suggest) { "Unknown region. Did you mean: $($suggest -join ', ')?" } else { "Unknown region." }
            }
        }
    }
    return ,$out
}

# ---------- SKUs ---------------------------------------------------------

function ConvertTo-QRSkuLookupKey {
    <#
    .SYNOPSIS
        Reduces a raw user string to a comparison key: lowercase, no spaces,
        no underscores, no dashes. Used to match against size names AND
        family names uniformly.
    #>
    param([Parameter(Mandatory)][string] $Raw)
    return (($Raw -replace '[\s_\-]', '')).ToLowerInvariant()
}

function Get-QRFamilyLookupKey {
    <#
    .SYNOPSIS
        Reduces a resourceSkus[].family value to a compact key that matches
        what ConvertTo-QRSkuLookupKey produces for user input.

        Examples:
          "standardDDSv5Family"       -> "ddsv5"
          "Standard NCASv3_T4 Family" -> "ncasv3t4"
    #>
    param([Parameter(Mandatory)][string] $Family)
    $k = $Family -replace '(?i)^standard', ''
    $k = $k      -replace '(?i)family$', ''
    return (($k -replace '[\s_\-]', '')).ToLowerInvariant()
}

function Resolve-QRSkus {
    <#
    .SYNOPSIS
        Validates SKU input against the union of resourceSkus catalogs for
        every target region and *aggregates to family level*. Accepts full
        names ('Standard_D4s_v5'), short sizes ('D4s_v5', 'D4s v5',
        'D4sv5'), AND families ('Ddsv5', 'Lasv4'). Every valid entry becomes
        one family entry with a SampleSize used for restriction lookups.

    .OUTPUTS
        pscustomobject rows with:
            Kind, Raw, Family, FamilyKey, SampleSize, SampleSizeSource,
            Valid, Reason
        Where:
          - Family        : the resourceSkus family string ("standardDDSv5Family").
          - FamilyKey     : the short lookup key ("ddsv5"), useful for humans.
          - SampleSize    : one canonical size in the family, used to look up
                            restrictions/zones (per your feedback: quota is
                            per-family, but restriction calls need a size).
          - SampleSizeSource: 'user' if the user typed that specific size,
                              'auto' if we picked the smallest offered size.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]] $Raw,
        [Parameter(Mandatory)][string]   $ProbeSubscriptionId,
        [Parameter(Mandatory)][string[]] $Regions
    )

    # ---- Build a union catalog across every target region ----
    $sizeByKey    = @{}   # normalized key -> canonical size name
    $familyOfSize = @{}   # canonical size name -> raw family string
    $vcpuOfSize   = @{}   # canonical size name -> vCPU count (for smallest-size pick)
    $familyByKey  = @{}   # normalized family key -> [list of size names]

    foreach ($region in $Regions) {
        $catalog = Get-QRComputeSkuCatalog -SubscriptionId $ProbeSubscriptionId -Region $region
        foreach ($sku in $catalog) {
            if ($sku.resourceType -ne 'virtualMachines') { continue }
            $name = $sku.name

            # Size keys: 'Standard_D4s_v5' matches 'standardd4sv5' AND 'd4sv5'
            $k1 = ConvertTo-QRSkuLookupKey -Raw $name
            $k2 = ConvertTo-QRSkuLookupKey -Raw ($name -replace '^Standard_','' -replace '^Basic_','')
            foreach ($k in @($k1, $k2)) {
                if (-not $sizeByKey.ContainsKey($k)) { $sizeByKey[$k] = $name }
            }
            if ($sku.family -and -not $familyOfSize.ContainsKey($name)) {
                $familyOfSize[$name] = $sku.family
            }
            if (-not $vcpuOfSize.ContainsKey($name)) {
                $vCap = @($sku.capabilities | Where-Object { $_.name -eq 'vCPUs' } | Select-Object -First 1)
                $vcpuOfSize[$name] = if ($vCap) { [int]$vCap[0].value } else { [int]::MaxValue }
            }
            if ($sku.family) {
                $fk = Get-QRFamilyLookupKey -Family $sku.family
                if (-not $familyByKey.ContainsKey($fk)) {
                    $familyByKey[$fk] = New-Object System.Collections.Generic.List[string]
                }
                if ($familyByKey[$fk] -notcontains $name) {
                    $familyByKey[$fk].Add($name) | Out-Null
                }
            }
        }
    }

    function _PickSmallest([string[]] $sizes) {
        # Prefer the size with the fewest vCPUs; ties broken by name for stability.
        return ($sizes | Sort-Object @{Expression={$vcpuOfSize[$_]}}, @{Expression={$_}} | Select-Object -First 1)
    }

    $out = New-Object System.Collections.Generic.List[object]
    foreach ($r in $Raw) {
        if (-not $r) { continue }
        $key = ConvertTo-QRSkuLookupKey -Raw $r

        if ($sizeByKey.ContainsKey($key)) {
            # User typed a size. Aggregate to its family; that size is the sample.
            $size   = $sizeByKey[$key]
            $family = $familyOfSize[$size]
            $out.Add([pscustomobject]@{
                Kind             = 'Sku'
                Raw              = $r
                Family           = $family
                FamilyKey        = if ($family) { Get-QRFamilyLookupKey -Family $family } else { '' }
                SampleSize       = $size
                SampleSizeSource = 'user'
                Valid            = $true
                Reason           = ''
            })
        }
        elseif ($familyByKey.ContainsKey($key)) {
            # User typed a family. Pick the smallest offered size in that family
            # as the restriction/zone probe.
            $sizes  = @($familyByKey[$key])
            $sample = _PickSmallest $sizes
            $family = $familyOfSize[$sample]
            $out.Add([pscustomobject]@{
                Kind             = 'Sku'
                Raw              = $r
                Family           = $family
                FamilyKey        = $key
                SampleSize       = $sample
                SampleSizeSource = 'auto'
                Valid            = $true
                Reason           = "family probe uses size '$sample' (auto-picked, smallest of $($sizes.Count) sizes)"
            })
        }
        else {
            $sizeSug = @($sizeByKey.Keys   | Where-Object { $_ -like "*$key*" -or $key -like "*$_*" } | Select-Object -First 5 | ForEach-Object { $sizeByKey[$_] })
            $famSug  = @($familyByKey.Keys | Where-Object { $_ -like "*$key*" -or $key -like "*$_*" } | Select-Object -First 5)
            $reason  =
                if ($famSug -and $sizeSug) {
                    "Unknown. Close families: $($famSug -join ', '). Close sizes: $($sizeSug -join ', ')."
                } elseif ($famSug) {
                    "Unknown. Close families: $($famSug -join ', ')."
                } elseif ($sizeSug) {
                    "Unknown. Close sizes: $($sizeSug -join ', ')."
                } else {
                    "Unknown SKU or family in the target regions ($(($Regions) -join ', '))."
                }
            $out.Add([pscustomobject]@{
                Kind             = 'Sku'
                Raw              = $r
                Family           = ''
                FamilyKey        = ''
                SampleSize       = ''
                SampleSizeSource = ''
                Valid            = $false
                Reason           = $reason
            })
        }
    }
    return ,$out.ToArray()
}
