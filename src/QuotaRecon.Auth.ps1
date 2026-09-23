# QuotaRecon.Auth.ps1
#
# Thin wrapper over Az.Accounts + Invoke-RestMethod. Deliberately avoids
# `az rest` / `az.cmd` because those re-parse arguments through cmd.exe and
# break on `?` in URLs. We reuse the current Az context so callers just
# `Connect-AzAccount` once.

$script:QuotaReconArmBase = 'https://management.azure.com'
$script:QuotaReconTokenCache = @{}   # audience -> @{ Token = ...; ExpiresOn = ... }

function Get-QRArmToken {
    <#
    .SYNOPSIS
        Returns a valid ARM bearer token for the current Az context, caching
        it until ~5 minutes before expiry.
    #>
    [CmdletBinding()]
    param(
        [string] $Audience = 'https://management.azure.com/'
    )

    $entry = $script:QuotaReconTokenCache[$Audience]
    if ($entry -and $entry.ExpiresOn -gt [datetime]::UtcNow.AddMinutes(5)) {
        return $entry.Token
    }

    $ctx = Get-AzContext -ErrorAction Stop
    if (-not $ctx) {
        throw "No Az context. Run Connect-AzAccount before invoking QuotaRecon."
    }

    # Get-AzAccessToken changed in Az 12 to return a SecureString by default.
    # Handle both shapes.
    $raw = Get-AzAccessToken -ResourceUrl $Audience -ErrorAction Stop
    $tokenText =
        if ($raw.Token -is [System.Security.SecureString]) {
            [System.Net.NetworkCredential]::new('', $raw.Token).Password
        } else {
            [string]$raw.Token
        }

    $expires =
        if ($raw.ExpiresOn -is [System.DateTimeOffset]) {
            $raw.ExpiresOn.UtcDateTime
        } elseif ($raw.ExpiresOn) {
            [datetime]$raw.ExpiresOn
        } else {
            (Get-Date).AddMinutes(30)
        }

    $script:QuotaReconTokenCache[$Audience] = @{
        Token     = $tokenText
        ExpiresOn = $expires
    }
    return $tokenText
}

function Invoke-QRArm {
    <#
    .SYNOPSIS
        Calls an ARM endpoint using the cached bearer token. Returns the
        parsed JSON body. Errors are surfaced with URL + status so the caller
        can log them into the Errors sheet without unwinding the run.
    .PARAMETER Path
        Path-and-query beginning with '/subscriptions/...' or '/providers/...'.
        Do NOT include the host.
    .PARAMETER Method
        HTTP verb. Defaults to GET.
    .PARAMETER Body
        Optional object; will be JSON-serialized.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $Path,

        [ValidateSet('GET','POST','PUT','DELETE','PATCH')]
        [string] $Method = 'GET',

        [object] $Body,

        [int] $TimeoutSec = 60
    )

    $token = Get-QRArmToken
    $url   = $script:QuotaReconArmBase + $Path
    $headers = @{
        Authorization = "Bearer $token"
        Accept        = 'application/json'
    }

    $params = @{
        Uri             = $url
        Method          = $Method
        Headers         = $headers
        TimeoutSec      = $TimeoutSec
        ErrorAction     = 'Stop'
    }
    if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
        $params.Body        = ($Body | ConvertTo-Json -Depth 20 -Compress)
        $params.ContentType = 'application/json'
    }

    try {
        Write-Verbose "[QRArm] $Method $Path"
        $resp = Invoke-RestMethod @params
        return $resp
    }
    catch {
        $status = $null
        if ($_.Exception.Response) {
            try { $status = [int]$_.Exception.Response.StatusCode } catch { }
        }
        $msg = $_.Exception.Message
        # On failure, dump the response body if we can - ARM 4xx often carries
        # a helpful reason in the body. Only visible with -Verbose or transcript.
        try {
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                Write-Verbose "[QRArm] ERROR body: $($_.ErrorDetails.Message)"
            }
        } catch { }
        throw [System.Exception]::new(
            "ARM $Method $Path failed$(if($status){" (HTTP $status)"}): $msg",
            $_.Exception)
    }
}

function Invoke-QRArmPaged {
    <#
    .SYNOPSIS
        GET wrapper that follows nextLink and returns the flat list of
        `value` items.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $Path,
        [int]    $TimeoutSec = 60
    )

    $items = New-Object System.Collections.Generic.List[object]
    $next  = $Path
    while ($next) {
        # nextLink from ARM is an absolute URL; strip the host so Invoke-QRArm can
        # prepend it back consistently.
        if ($next.StartsWith('https://')) {
            $next = $next -replace '^https://[^/]+', ''
        }
        $page = Invoke-QRArm -Path $next -TimeoutSec $TimeoutSec
        if ($page.value) { $items.AddRange([object[]]$page.value) }
        $next = $page.nextLink
    }
    return ,$items.ToArray()
}
